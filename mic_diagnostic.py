import sounddevice as sd
import numpy as np
import sys

# Ensure UTF-8 output encoding for Windows terminal
sys.stdout.reconfigure(encoding='utf-8')

def run_mic_diagnostics():
    print("=" * 65)
    print("        [+] COMPREHENSIVE MICROPHONE HARDWARE DIAGNOSTICS")
    print("=" * 65)
    
    devices = sd.query_devices()
    input_devices = []
    
    for idx, d in enumerate(devices):
        if d['max_input_channels'] > 0:
            input_devices.append((idx, d))
            
    print(f"Total audio input endpoints detected: {len(input_devices)}\n")
    
    working_mics = []
    silent_mics = []
    failed_mics = []
    
    for idx, d in input_devices:
        name = d['name']
        max_ch = d['max_input_channels']
        native_sr = int(d['default_samplerate'])
        host_api = sd.query_hostapis(d['hostapi'])['name']
        
        print(f"Testing Device #{idx}: [{host_api}] {name}")
        print(f"  Max Channels: {max_ch} | Native Sample Rate: {native_sr}Hz")
        
        ch = min(max_ch, 2)
        success = False
        tested_sr = None
        max_amp = 0
        mean_amp = 0.0
        err_msg = ""
        
        for test_rate in [native_sr, 48000, 44100, 16000]:
            try:
                rec = sd.rec(int(test_rate * 0.3), samplerate=test_rate, channels=ch, dtype='int16', device=idx)
                sd.wait()
                audio_flat = rec.flatten()
                max_amp = int(np.max(np.abs(audio_flat)))
                mean_amp = float(np.mean(np.abs(audio_flat)))
                success = True
                tested_sr = test_rate
                break
            except Exception as e:
                err_msg = str(e)
                
        if success:
            if max_amp > 100:
                status = f"[WORKING] RECEIVING AUDIO (Max Amp: {max_amp}, Mean: {mean_amp:.1f})"
                working_mics.append((idx, name, host_api, tested_sr, ch, max_amp))
            else:
                status = f"[SILENT] MUTED / NO AUDIO SIGNAL (Max Amp: {max_amp})"
                silent_mics.append((idx, name, host_api, tested_sr, ch))
        else:
            status = f"[FAILED] INACCESSIBLE ({err_msg})"
            failed_mics.append((idx, name, host_api, err_msg))
            
        print(f"  Status: {status}\n")

    print("=" * 65)
    print("                DIAGNOSTIC SUMMARY REPORT")
    print("=" * 65)
    
    print(f"\n[+] ACTIVE / WORKING MICROPHONES ({len(working_mics)}):")
    if working_mics:
        for idx, name, api, sr, ch, amp in working_mics:
            is_bt = any(k in name.lower() for k in ["bluetooth", "hands-free", "headset", "airpods", "buds"])
            bt_tag = " [BLUETOOTH MIC]" if is_bt else ""
            print(f"  * Device #{idx}: {name} ({api}, {sr}Hz, {ch}ch) - Peak Amp: {amp}{bt_tag}")
    else:
        print("  None detected with active audio signal!")
        
    print(f"\n[-] SILENT / INACTIVE MICROPHONES ({len(silent_mics)}):")
    for idx, name, api, sr, ch in silent_mics:
        is_bt = any(k in name.lower() for k in ["bluetooth", "hands-free", "headset", "airpods", "buds"])
        bt_tag = " [BLUETOOTH MIC]" if is_bt else ""
        print(f"  * Device #{idx}: {name} ({api}){bt_tag}")
        
    if failed_mics:
        print(f"\n[!] INACCESSIBLE MICROPHONES ({len(failed_mics)}):")
        for idx, name, api, err in failed_mics:
            print(f"  * Device #{idx}: {name} ({api}) -> {err}")

if __name__ == "__main__":
    run_mic_diagnostics()
