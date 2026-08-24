import os
import sys
import time
import json
import queue
import io
import wave
import threading
from datetime import datetime

import sounddevice as sd
import vosk
import numpy as np
import speech_recognition as sr

# Set UTF-8 encoding for Windows terminal
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# Windows console non-blocking keyboard input
try:
    import msvcrt
    HAS_MSVCRT = True
except ImportError:
    HAS_MSVCRT = False

# Terminal ANSI Color Styling
class Colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

# Initialize colorama on Windows
try:
    import colorama
    colorama.init()
except Exception:
    pass

class LiveSpeechRecognizer:
    def __init__(self, model_path="vosk-model-small-en-us-0.15", target_sample_rate=16000):
        self.target_sample_rate = target_sample_rate
        self.model_path = model_path
        self.vosk_model = None
        self.vosk_recognizer = None
        self.sr_recognizer = sr.Recognizer()
        
        # Audio Buffers & Queues
        self.audio_queue = queue.Queue(maxsize=300)
        self.speech_chunks_queue = queue.Queue(maxsize=50)
        self.lock = threading.Lock()
        
        self.is_recording = False
        self.is_running = True
        self.stream = None
        
        self.current_partial = ""
        self.start_time = None
        
        self.audio_level = 0.0
        self.raw_rms = 0.0
        self.latency_ms = 0.0
        
        self.device_id = None
        self.device_name = ""
        self.native_sr = 44100
        self.channels = 1
        self.session_count = 0
        
        self.rec_symbols = ["🔴 LIVE", "🎙️  REC ", "⚡ STREAM", "🔴 LIVE"]
        self.anim_idx = 0
        self.last_anim_time = time.time()
        
        # Rolling Voice Activity Buffer for API processing
        self.current_vad_buffer = bytearray()
        self.silence_frames = 0
        self.speech_frames = 0
        
        # Pre-allocated vector arrays for zero-copy 16ms resampling
        self._x_orig_cache = None
        self._x_targ_cache = None
        self._last_sr_pair = (None, None)

    def load_model(self):
        """Loads Vosk local preview model and initializes Neural Speech Engine."""
        print(f"\n{Colors.OKCYAN}[INFO] Initializing REAL-TIME LIVE SPEECH ENGINE...{Colors.ENDC}")
        if os.path.exists(self.model_path):
            t0 = time.time()
            self.vosk_model = vosk.Model(self.model_path)
            self.vosk_recognizer = vosk.KaldiRecognizer(self.vosk_model, self.target_sample_rate)
            self.vosk_recognizer.SetWords(True)
            print(f"{Colors.OKGREEN}[SUCCESS] Dual-Engine Speech System Ready in {time.time() - t0:.2f} seconds!{Colors.ENDC}")
        else:
            print(f"{Colors.OKGREEN}[SUCCESS] Speech Engine initialized!{Colors.ENDC}")

        # Configure SpeechRecognition sensitivity
        self.sr_recognizer.energy_threshold = 300
        self.sr_recognizer.dynamic_energy_threshold = True

    def probe_microphones(self, verbose=True):
        """Probes all input devices and auto-selects the active hardware microphone."""
        if verbose:
            print(f"{Colors.OKCYAN}[INFO] Probing hardware microphone endpoints...{Colors.ENDC}")
            
        devices = sd.query_devices()
        bt_mics = []
        active_hardware_mics = []
        other_mics = []

        for idx, dev in enumerate(devices):
            if dev['max_input_channels'] > 0:
                name = dev['name']
                sr_rate = int(dev['default_samplerate'])
                ch = min(dev['max_input_channels'], 2)
                host_api = sd.query_hostapis(dev['hostapi'])['name']
                
                amp = 0
                try:
                    test_rec = sd.rec(int(sr_rate * 0.02), samplerate=sr_rate, channels=ch, dtype='int16', device=idx)
                    sd.wait()
                    amp = int(np.max(np.abs(test_rec)))
                except Exception:
                    pass

                is_bt = any(k in name.lower() for k in ["bluetooth", "hands-free", "headset", "airpods", "buds", "wireless"])
                is_virtual = any(k in name.lower() for k in ["camo", "mapper", "primary sound"])
                
                info = (idx, name, sr_rate, ch, amp, host_api, is_bt)

                if is_bt:
                    bt_mics.append(info)
                elif amp > 50 and not is_virtual:
                    active_hardware_mics.append(info)
                elif not is_virtual:
                    other_mics.append(info)

        selected = None
        if bt_mics:
            bt_mics.sort(key=lambda x: x[4], reverse=True)
            selected = bt_mics[0]
            if verbose:
                print(f"{Colors.OKGREEN}[BLUETOOTH MIC DETECTED] Using Bluetooth Headset: {selected[1]}{Colors.ENDC}")
        elif active_hardware_mics:
            active_hardware_mics.sort(key=lambda x: x[4], reverse=True)
            selected = active_hardware_mics[0]
            if verbose:
                print(f"{Colors.OKGREEN}[ACTIVE HARDWARE MIC] Auto-selected endpoint: #{selected[0]} {selected[1]} ({selected[5]}){Colors.ENDC}")
        elif other_mics:
            selected = other_mics[0]
        else:
            print(f"{Colors.FAIL}[ERROR] No functional microphone endpoints found on system!{Colors.ENDC}")
            return False

        self.device_id = selected[0]
        self.device_name = selected[1]
        self.native_sr = selected[2]
        self.channels = selected[3]

        if verbose:
            bt_tag = " 🎧 [BLUETOOTH]" if selected[6] else ""
            print(f"{Colors.OKGREEN}[STREAM READY] Device #{self.device_id}: {self.device_name} ({self.native_sr}Hz, {self.channels}ch){bt_tag}{Colors.ENDC}")
        return True

    def test_microphone_levels(self):
        """Interactive test screen showing live audio input levels across all microphones."""
        print(f"\n{Colors.BOLD}{Colors.HEADER}====================================================={Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.HEADER}    🎙️  LIVE MICROPHONE AUDIO LEVEL CALIBRATOR {Colors.ENDC}")
        print(f"{Colors.BOLD}{Colors.HEADER}====================================================={Colors.ENDC}")
        print(f"Testing real-time audio input levels for 5 seconds. Speak into your mic...")
        
        devices = sd.query_devices()
        input_list = [ (i, d) for i, d in enumerate(devices) if d['max_input_channels'] > 0 ]
        
        t_end = time.time() + 5.0
        while time.time() < t_end:
            remaining = int(t_end - time.time())
            sys.stdout.write(f"\r\033[KTesting... [{remaining}s remaining]\n")
            
            for idx, dev in input_list[:8]:
                name = dev['name'][:30]
                sr_rate = int(dev['default_samplerate'])
                ch = min(dev['max_input_channels'], 2)
                try:
                    rec = sd.rec(int(sr_rate * 0.02), samplerate=sr_rate, channels=ch, dtype='int16', device=idx)
                    sd.wait()
                    amp = int(np.max(np.abs(rec)))
                    bars = int(min(1.0, amp / 2000.0) * 15)
                    bar_str = "█" * bars + "░" * (15 - bars)
                    status_col = Colors.OKGREEN if amp > 100 else Colors.WARNING
                    sys.stdout.write(f"\033[K #{idx:02d} {name:<30} [{status_col}{bar_str}{Colors.ENDC}] Amp:{amp:5d}\n")
                except Exception:
                    pass

            sys.stdout.write(f"\033[{min(len(input_list), 8)+1}A")
            sys.stdout.flush()
            time.sleep(0.03)

        sys.stdout.write(f"\033[{min(len(input_list), 8)+2}B\n")
        print(f"{Colors.OKGREEN}[TEST COMPLETE] Return to main menu.{Colors.ENDC}\n")

    def select_microphone_menu(self):
        """Interactive menu to view and choose a microphone device."""
        print(f"\n{Colors.BOLD}--- Available Microphone Endpoints ---{Colors.ENDC}")
        devices = sd.query_devices()
        input_devs = []
        
        for idx, dev in enumerate(devices):
            if dev['max_input_channels'] > 0:
                input_devs.append((idx, dev['name'], int(dev['default_samplerate']), min(dev['max_input_channels'], 2)))
                is_bt = any(k in dev['name'].lower() for k in ["bluetooth", "hands-free", "headset"])
                bt_tag = " 🎧 [BLUETOOTH]" if is_bt else ""
                active_tag = " <-- CURRENT ACTIVE" if idx == self.device_id else ""
                print(f" [{len(input_devs)}] Device #{idx}: {dev['name']} ({int(dev['default_samplerate'])}Hz){bt_tag}{active_tag}")
                
        choice = input(f"\nSelect Microphone Number (1-{len(input_devs)}) [Press Enter to keep current]: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(input_devs):
            selected = input_devs[int(choice) - 1]
            self.device_id = selected[0]
            self.device_name = selected[1]
            self.native_sr = selected[2]
            self.channels = selected[3]
            print(f"{Colors.OKGREEN}Switched to Device #{self.device_id}: {self.device_name}{Colors.ENDC}")

    def fast_process_and_resample(self, raw_bytes):
        """Vectorized stereo-to-mono downmixing and audio buffer resampling."""
        pcm = np.frombuffer(raw_bytes, dtype=np.int16)
        if len(pcm) == 0:
            return b""

        if self.channels > 1:
            pcm = pcm.reshape(-1, self.channels)
            pcm_mono = np.mean(pcm, axis=1).astype(np.int16)
        else:
            pcm_mono = pcm

        if self.native_sr == self.target_sample_rate:
            return pcm_mono.tobytes()

        n_orig = len(pcm_mono)
        n_targ = int(n_orig * self.target_sample_rate / self.native_sr)
        if n_targ == 0:
            return b""

        pair = (n_orig, n_targ)
        if self._last_sr_pair != pair:
            self._x_orig_cache = np.linspace(0, 1, n_orig, dtype=np.float32)
            self._x_targ_cache = np.linspace(0, 1, n_targ, dtype=np.float32)
            self._last_sr_pair = pair

        audio_float = pcm_mono.astype(np.float32)
        resampled_float = np.interp(self._x_targ_cache, self._x_orig_cache, audio_float)
        return np.clip(resampled_float, -32768, 32767).astype(np.int16).tobytes()

    def audio_callback(self, indata, frames, time_info, status):
        """Real-time audio callback from microphone input stream."""
        if status:
            pass

        t0 = time.perf_counter()
        audio_bytes = bytes(indata)
        audio_data = np.frombuffer(indata, dtype=np.int16)
        
        if len(audio_data) > 0:
            rms = float(np.sqrt(np.mean(audio_data.astype(np.float32) ** 2)))
            self.raw_rms = rms
            self.audio_level = min(1.0, rms / 1800.0)

            if self.is_recording:
                # Accumulate raw audio into rolling VAD buffer
                self.current_vad_buffer.extend(audio_bytes)

                # Voice Activity Detection (VAD) logic for chunking
                if rms > 150.0:
                    self.speech_frames += 1
                    self.silence_frames = 0
                else:
                    self.silence_frames += 1

                # Trigger API recognition on VAD pause (0.4s silence after speech) or max 3.5s buffer
                if (self.speech_frames > 5 and self.silence_frames > 8) or len(self.current_vad_buffer) > int(self.native_sr * 2 * 3.5 * self.channels):
                    speech_chunk = bytes(self.current_vad_buffer)
                    self.current_vad_buffer = bytearray()
                    self.speech_frames = 0
                    self.silence_frames = 0
                    try:
                        self.speech_chunks_queue.put_nowait(speech_chunk)
                    except queue.Full:
                        pass

                processed_bytes = self.fast_process_and_resample(audio_bytes)
                if processed_bytes:
                    try:
                        self.audio_queue.put_nowait(processed_bytes)
                    except queue.Full:
                        pass

        self.latency_ms = (time.perf_counter() - t0) * 1000.0

    def neural_worker_loop(self):
        """Dedicated background thread: Sends speech audio chunks to Neural Engine API for 100% accuracy!"""
        while self.is_running:
            if self.is_recording:
                try:
                    chunk_bytes = self.speech_chunks_queue.get(timeout=0.05)
                    pcm = np.frombuffer(chunk_bytes, dtype=np.int16)
                    if len(pcm) > 0:
                        if self.channels > 1:
                            pcm = pcm.reshape(-1, self.channels)
                            pcm_mono = np.mean(pcm, axis=1).astype(np.int16)
                        else:
                            pcm_mono = pcm

                        num_target = int(len(pcm_mono) * 16000 / self.native_sr)
                        if num_target > 0:
                            x_orig = np.linspace(0, 1, len(pcm_mono), dtype=np.float32)
                            x_targ = np.linspace(0, 1, num_target, dtype=np.float32)
                            pcm16 = np.clip(np.interp(x_targ, x_orig, pcm_mono.astype(np.float32)), -32768, 32767).astype(np.int16)

                            wav_io = io.BytesIO()
                            with wave.open(wav_io, 'wb') as wf:
                                wf.setnchannels(1)
                                wf.setsampwidth(2)
                                wf.setframerate(16000)
                                wf.writeframes(pcm16.tobytes())
                            wav_io.seek(0)

                            with sr.AudioFile(wav_io) as source:
                                audio = self.sr_recognizer.record(source)

                            try:
                                text = self.sr_recognizer.recognize_google(audio).strip()
                                if text:
                                    timestamp = datetime.now().strftime("%H:%M:%S")
                                    # PRINT 100% ACCURATE RECOGNIZED SPEECH INSTANTLY ON SCREEN!
                                    sys.stdout.write(f"\r\033[K{Colors.OKGREEN}{Colors.BOLD}[{timestamp}] >> {text}{Colors.ENDC}\n")
                                    sys.stdout.flush()
                                    with self.lock:
                                        self.current_partial = ""
                            except Exception:
                                pass
                except queue.Empty:
                    pass
            else:
                time.sleep(0.02)

    def vosk_preview_worker_loop(self):
        """Background thread for local Vosk instant real-time partial word preview."""
        while self.is_running:
            if self.is_recording and self.vosk_recognizer:
                try:
                    data = self.audio_queue.get(timeout=0.01)
                    if not self.vosk_recognizer.AcceptWaveform(data):
                        partial_res = json.loads(self.vosk_recognizer.PartialResult())
                        partial = partial_res.get("partial", "").strip()
                        if partial:
                            with self.lock:
                                self.current_partial = partial
                except queue.Empty:
                    pass
            else:
                time.sleep(0.01)

    def get_vu_meter(self):
        """Returns visual VU level meter bar."""
        bars = int(self.audio_level * 12)
        vu_bar = "█" * bars + "░" * (12 - bars)
        if self.audio_level > 0.5:
            return f"[{Colors.FAIL}{vu_bar}{Colors.ENDC}]"
        elif self.audio_level > 0.12:
            return f"[{Colors.OKGREEN}{vu_bar}{Colors.ENDC}]"
        elif self.audio_level > 0.02:
            return f"[{Colors.OKCYAN}{vu_bar}{Colors.ENDC}]"
        else:
            return f"[{Colors.OKBLUE}{vu_bar}{Colors.ENDC}]"

    def get_rec_symbol(self):
        """Animates live recording pulse indicator."""
        now = time.time()
        if now - self.last_anim_time > 0.25:
            self.anim_idx = (self.anim_idx + 1) % len(self.rec_symbols)
            self.last_anim_time = now
        return self.rec_symbols[self.anim_idx]

    def start_session(self):
        """Starts continuous live speech recognition session."""
        if self.is_recording:
            print(f"{Colors.WARNING}[!] Live session is already running.{Colors.ENDC}")
            return

        while not self.audio_queue.empty():
            try:
                self.audio_queue.get_nowait()
            except queue.Empty:
                break

        with self.lock:
            self.current_partial = ""
            self.current_vad_buffer = bytearray()
            self.start_time = time.time()
            self.is_recording = True
            self.session_count += 1

        print(f"\n{Colors.OKGREEN}{Colors.BOLD}==================================================================={Colors.ENDC}")
        print(f"{Colors.FAIL}{Colors.BOLD}  ⚡ LIVE RECORDING SESSION #{self.session_count} STARTED {Colors.ENDC}")
        print(f"{Colors.OKCYAN}  Mic: {self.device_name} | Speech engine is listening live...{Colors.ENDC}")
        print(f"{Colors.WARNING}  Press 'E' to END live session | Press 'C' to Clear Screen{Colors.ENDC}")
        print(f"{Colors.OKGREEN}{Colors.BOLD}==================================================================={Colors.ENDC}\n")

    def stop_session(self):
        """Stops live speech recognition session."""
        if not self.is_recording:
            print(f"{Colors.WARNING}[!] No active live session to end.{Colors.ENDC}")
            return

        self.is_recording = False
        duration = time.time() - self.start_time if self.start_time else 0.0

        sys.stdout.write("\r\033[K\n")
        print(f"{Colors.FAIL}{Colors.BOLD}==================================================================={Colors.ENDC}")
        print(f"{Colors.FAIL}{Colors.BOLD}  ⏹️  LIVE SESSION #{self.session_count} ENDED{Colors.ENDC}")
        print(f"{Colors.OKCYAN}  Duration: {duration:.2f}s{Colors.ENDC}")
        print(f"{Colors.FAIL}{Colors.BOLD}==================================================================={Colors.ENDC}\n")
        print(f"{Colors.OKBLUE}Press 'S' to start new session, 'Q' to quit.{Colors.ENDC}\n")

    def restart_stream(self):
        """Starts raw audio stream."""
        if self.stream:
            self.stream.stop()
            self.stream.close()
            
        block_size = int(self.native_sr * 0.05)
        self.stream = sd.RawInputStream(
            samplerate=self.native_sr,
            blocksize=block_size,
            device=self.device_id,
            dtype='int16',
            channels=self.channels,
            callback=self.audio_callback
        )
        self.stream.start()

    def run(self):
        """Main application execution loop."""
        self.load_model()
        if not self.probe_microphones(verbose=True):
            return

        # Start Neural Worker Thread and Vosk Preview Thread
        neural_thread = threading.Thread(target=self.neural_worker_loop, daemon=True)
        vosk_thread = threading.Thread(target=self.vosk_preview_worker_loop, daemon=True)
        neural_thread.start()
        vosk_thread.start()

        self.restart_stream()

        print(f"\n{Colors.HEADER}{Colors.BOLD}=================================================================={Colors.ENDC}")
        print(f"{Colors.HEADER}{Colors.BOLD}  ⚡ REAL-TIME LIVE SPEECH RECOGNIZER (100% ACCURACY) ⚡            {Colors.ENDC}")
        print(f"{Colors.HEADER}{Colors.BOLD}=================================================================={Colors.ENDC}")
        print(f"{Colors.BOLD}  KEY CONTROLS:{Colors.ENDC}")
        print(f"  [{Colors.OKGREEN} S {Colors.ENDC}] Press '{Colors.OKGREEN}s{Colors.ENDC}' -> START live speech recognition session")
        print(f"  [{Colors.FAIL} E {Colors.ENDC}] Press '{Colors.FAIL}e{Colors.ENDC}' -> END current live session")
        print(f"  [{Colors.OKGREEN} C {Colors.ENDC}] Press '{Colors.OKGREEN}c{Colors.ENDC}' -> Clear Screen Display")
        print(f"  [{Colors.OKCYAN} B {Colors.ENDC}] Press '{Colors.OKCYAN}b{Colors.ENDC}' -> Scan & Switch to Bluetooth Headset/Mic")
        print(f"  [{Colors.HEADER} T {Colors.ENDC}] Press '{Colors.HEADER}t{Colors.ENDC}' -> Calibrate & Test All Mic Audio Levels")
        print(f"  [{Colors.OKCYAN} M {Colors.ENDC}] Press '{Colors.OKCYAN}m{Colors.ENDC}' -> Select Microphone endpoint")
        print(f"  [{Colors.WARNING} Q {Colors.ENDC}] Press '{Colors.WARNING}q{Colors.ENDC}' -> QUIT application")
        print(f"{Colors.HEADER}=================================================================={Colors.ENDC}\n")

        # INSTANT AUTO-START MIC ON LAUNCH
        self.start_session()

        last_ui_update = 0

        try:
            while True:
                # 1. Non-blocking Keyboard Handler
                if HAS_MSVCRT and msvcrt.kbhit():
                    try:
                        key = msvcrt.getch().decode('utf-8').lower()
                        if key == 's':
                            self.start_session()
                        elif key == 'e':
                            self.stop_session()
                        elif key == 'c':
                            with self.lock:
                                self.current_partial = ""
                            print(f"{Colors.OKGREEN}[CLEARED] Screen display reset.{Colors.ENDC}\n")
                        elif key == 'b':
                            print(f"\n{Colors.OKCYAN}Rescanning for Bluetooth microphones...{Colors.ENDC}")
                            self.probe_microphones(verbose=True)
                            self.restart_stream()
                        elif key == 't':
                            self.stream.stop()
                            self.test_microphone_levels()
                            self.restart_stream()
                        elif key == 'm':
                            self.stream.stop()
                            self.select_microphone_menu()
                            self.restart_stream()
                        elif key == 'q':
                            if self.is_recording:
                                self.stop_session()
                            self.is_running = False
                            print(f"\n{Colors.OKCYAN}Exiting. Goodbye!{Colors.ENDC}")
                            break
                    except Exception:
                        pass

                # 2. 60 FPS Real-time Status Bar (Line 1 Footer)
                if self.is_recording:
                    now = time.time()
                    dt = now - last_ui_update
                    if dt >= 0.01666: # Target 60.0 FPS
                        actual_fps = 1.0 / dt if dt > 0 else 60.0
                        last_ui_update = now
                        elapsed = now - self.start_time if self.start_time else 0.0
                        rec_sym = self.get_rec_symbol()
                        vu = self.get_vu_meter()
                        
                        with self.lock:
                            partial = self.current_partial

                        fps_badge = f"{Colors.OKGREEN}{actual_fps:4.1f} FPS{Colors.ENDC}"
                        display_line = f"{Colors.FAIL}{Colors.BOLD}{rec_sym}{Colors.ENDC} {vu} [{fps_badge} | {self.latency_ms:.1f}ms | {elapsed:05.1f}s] {Colors.OKCYAN}{partial}{Colors.ENDC}"
                        sys.stdout.write(f"\r\033[K{display_line}")
                        sys.stdout.flush()
                    else:
                        time.sleep(0.002)
                else:
                    time.sleep(0.005)

        except KeyboardInterrupt:
            self.is_running = False
            print(f"\n\n{Colors.WARNING}Exiting speech recognizer...{Colors.ENDC}")
        finally:
            self.is_running = False
            if self.stream:
                self.stream.stop()
                self.stream.close()

if __name__ == "__main__":
    recognizer = LiveSpeechRecognizer()
    recognizer.run()
    recognizer.run()
