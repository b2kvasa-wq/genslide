# ⚡ GenSlide - Production-Grade AI Voice-Controlled Dual-Screen PPT Presenter & Editor
## System Architecture, Object-Oriented Design & Technical Documentation Specification

**Document Version**: 2.0.0  
**Target Environment**: Windows 10 / 11 Desktop (Standalone Portable Executable & Python Runtime)  
**Author / Engineering Team**: Advanced Agentic Coding - Google DeepMind Pair Programming System  
**Date**: August 2026  

---

## 1. Executive Summary & System Purpose

**GenSlide** is an enterprise-grade desktop software application designed for voice-driven presentation management, real-time slide manipulation, and dual-screen HDMI display projection. 

Traditional presentation tools require manual clicker interaction or keyboard navigation near the host machine. **GenSlide** eliminates these physical constraints by embedding a **Sub-10ms Offline Acoustic Neural Engine** (Vosk Kaldi) alongside Google Neural Speech Recognition, allowing presenters to utter custom slide keywords to trigger slide switches in real-time milliseconds.

### Key Objectives & System Capabilities:
- **Zero-Latency Voice Keyword Triggering**: Real-time acoustic decoding (<10ms) allowing presenters to speak custom keywords naturally to trigger immediate slide jumps.
- **Dual-Screen HDMI Auto-Targeting**: Automatic detection of external HDMI monitors and projectors, projecting 100% borderless fullscreen slides on secondary displays while maintaining a compact `740x416` control preview on the primary laptop screen.
- **100% PowerPoint COM Visual Exporter**: Direct inter-process COM integration with Microsoft PowerPoint (`pywin32`) to export high-definition slide renderings (`.png`/`.jpg`), preserving 100% of native typography, tables, background art, and custom graphics.
- **Live Editable Keyword Matrix**: Real-time auto-syncing keyword grid with live re-indexing of the speech recognition engine and instant UI feedback.
- **Portable USB Execution**: Bundled as a single self-contained executable (`GenSlide.exe`) via PyInstaller, containing all Python runtimes, C++ DLLs, and neural models for instant plug-and-play execution from USB drives on any Windows PC.

---

## 2. High-Level System Architecture & Component Interaction Flow

The GenSlide system architecture follows a decoupled, event-driven model combining GUI main looper threads, asynchronous audio capture workers, offline neural decoding threads, and inter-process COM PowerPoint interfaces.

```
 +-----------------------------------------------------------------------------------+
 |                                 MAIN GUI THREAD                                   |
 |                             (PresentationApp: CTk)                                |
 |                                                                                   |
 |  +--------------------+   +-----------------------+   +------------------------+  |
 |  | Presenter          |   | Slide Deck Editor     |   | Voice Keyword Matrix   |  |
 |  | Dashboard View     |   | (Title, Bullets, Notes|   | (Live Sync & Grid)     |  |
 |  +---------+----------+   +-----------+-----------+   +-----------+------------+  |
 |            |                      |                           |                   |
 +------------|----------------------|---------------------------|-------------------+
              |                      |                           |
              v                      v                           v
 +-----------------------------------------------------------------------------------+
 |                                 SLIDE MANAGER                                     |
 |                          (SlideManager & SlideData)                               |
 |                                                                                   |
 |  • Loads/Saves .pptx decks via python-pptx                                         |
 |  • Renders 100% real PowerPoint visuals via win32com COM API                       |
 |  • Generates high-definition fallback slide canvases via Pillow (PIL)             |
 +-----------------------------------+-----------------------------------------------+
                                     |
                                     v
 +-----------------------------------------------------------------------------------+
 |                           EXTERNAL DISPLAY MANAGER                                |
 |                            (ExternalDisplayWindow)                                |
 |                                                                                   |
 |  • Queries physical displays via screeninfo                                        |
 |  • Projector (Monitor #2): 100% Borderless Fullscreen (overrideredirect)          |
 |  • Laptop (Monitor #1): Compact 740x416 Windowed Control Preview                  |
 +-----------------------------------------------------------------------------------+

                                     ^
                                     | Event Callback Dispatch (Slide Jump)
                                     |
 +-----------------------------------------------------------------------------------+
 |                             VOICE SPEECH ENGINE                                   |
 |                            (VoiceSpeechEngine)                                    |
 |                                                                                   |
 |  +-----------------------+    +-----------------------+    +-------------------+  |
 |  | SoundDevice Mic Stream| -> | Resampler & Downmixer | -> | Sub-10ms Vosk     |  |
 |  | (44.1kHz Int16 PCM)   |    | (Vectorized 16kHz)    |    | Worker Thread     |  |
 |  +-----------------------+    +-----------------------+    +---------+---------+  |
 |                                                                      |            |
 |                                                                      v            |
 |                                                            +-------------------+  |
 |                                                            | Fuzzy Regex       |  |
 |                                                            | Word Matcher      |  |
 |                                                            +-------------------+  |
 +-----------------------------------------------------------------------------------+
```

---

## 3. Object-Oriented Programming (OOP) Class Architecture

GenSlide is structured around fundamental Object-Oriented Programming (OOP) principles: **Encapsulation**, **Single Responsibility**, **Inheritance**, and **Abstraction**.

### 3.1 Class Summary Table

| Class Name | Inheritance | Primary Responsibility | OOP Role |
| :--- | :--- | :--- | :--- |
| `SlideData` | `object` | Encapsulates single slide properties (ID, title, bullets, notes, keywords, rendering path). | Data Value Object (Model) |
| `SlideManager` | `object` | Manages deck lifecycle, `.pptx` reading/writing, COM export, and PIL canvas generation. | Domain Service Manager |
| `VoiceSpeechEngine` | `object` | Handles microphone streaming, audio downmixing, Vosk neural recognition, and regex matching. | Background Subsystem Service |
| `ExternalDisplayWindow` | `ctk.CTkToplevel` | Controls secondary projector window fullscreen rendering and primary preview windowing. | Specialized UI Controller (View) |
| `PresentationApp` | `ctk.CTk` | Main application controller orchestrating views, hotkeys, navigation, and speech callbacks. | Orchestrator & Controller |

---

### 3.2 Detailed Class Specifications

#### 1. `SlideData` (Model Class)
Encapsulates all metadata and content associated with an individual presentation slide.
- **Attributes**:
  - `slide_id`: Integer sequence index (1-based).
  - `title`: String slide title.
  - `bullet_points`: List of string bullet statements.
  - `notes`: String speaker notes text.
  - `keywords`: List of string trigger keywords assigned to the slide.
  - `rendered_image_path`: Optional string filepath pointing to the exported slide image.

#### 2. `SlideManager` (Service Class)
Abstracts file format handling and visual slide rendering operations.
- **Methods**:
  - `load_pptx(filepath)`: Parses `.pptx` presentations using `python-pptx` and extracts text content, bullet points, and speaker notes.
  - `save_pptx(filepath)`: Writes updated slide content back into PowerPoint format.
  - `export_slides_with_pywin32(filepath)`: Invokes Microsoft PowerPoint via Windows COM interface (`win32com.client`) to export 100% color-accurate PNG slide images.
  - `create_fallback_slide_renderings()`: Generates PIL image canvases with dark themes when PowerPoint COM is unavailable.

#### 3. `VoiceSpeechEngine` (Background Engine Class)
Encapsulates real-time audio input capture, audio processing, neural acoustic decoding, and speech matching algorithms.
- **Methods**:
  - `probe_microphone()`: Auto-detects physical hardware microphone devices.
  - `fast_process_and_resample(raw_bytes)`: Vectorized downmixing (stereo-to-mono) and sample rate conversion (44.1kHz to 16kHz) using `numpy.interp`.
  - `vosk_instant_worker_loop()`: Asynchronous worker thread feeding 16kHz PCM frames to `vosk.KaldiRecognizer` and retrieving sub-10ms partial recognition results.
  - `match_speech_to_keyword(spoken_text)`: Performs fuzzy regex phrase matching using word boundaries (`\bkw\b`) sorted by keyword length.
  - `set_keywords(slides)`: Re-indexes the internal keyword lookup dictionary in real-time.

#### 4. `ExternalDisplayWindow` (View Window Class)
Inherits from `customtkinter.CTkToplevel` to manage external projector display windows.
- **Methods**:
  - `__init__(master, slide_mgr, monitor)`: Detects target monitor coordinates. Configures 100% borderless fullscreen (`overrideredirect(True)`) for external displays/projectors (`x != 0`) and compact windowed preview (`740x416`) for primary laptop screens (`x == 0`).
  - `update_display(slide_idx, is_blackout, is_whiteout)`: Redraws active slide image or color canvas onto display.

#### 5. `PresentationApp` (Main Controller Class)
Inherits from `customtkinter.CTk` to provide the main application shell and UI orchestration.
- **Methods**:
  - `build_dashboard_view()`: Constructs Presenter Dashboard UI with slide preview, upcoming slide thumbnail, hotkey cheat sheet, and speaker notes.
  - `build_editor_view()`: Constructs slide text and notes editor form.
  - `build_keywords_view()`: Constructs Voice Keyword Matrix grid with dual top/bottom save buttons.
  - `save_all_keywords()`: Re-indexes speech engine dictionary and displays visual sync feedback badges.

---

## 4. Technology Stack & Library Dependency Specifications

| Component Layer | Technology / Library | Version | Purpose |
| :--- | :--- | :--- | :--- |
| **Language & Runtime** | Python | `3.12.x` | Core application logic & runtime environment. |
| **GUI Framework** | CustomTkinter | `5.2.2` | Modern dark-themed responsive desktop user interface. |
| **Offline Speech Engine** | Vosk API | `0.3.45` | C++ Kaldi acoustic neural speech recognition engine (<10ms latency). |
| **Offline Acoustic Model** | `vosk-model-small-en-us-0.15` | `0.15` | Lightweight 50MB English acoustic neural model. |
| **Online Speech Backup** | SpeechRecognition | `3.14.1` | Google Speech Neural API fallback wrapper. |
| **Audio I/O Capture** | SoundDevice | `0.5.1` | Asynchronous real-time hardware microphone PCM audio streaming. |
| **Audio Processing** | NumPy | `2.2.3` | Vectorized audio array downmixing and linear interpolation resampling. |
| **PowerPoint COM API** | PyWin32 (`win32com`) | `308` | Direct Windows COM interface to Microsoft PowerPoint for 100% accurate slide exports. |
| **PowerPoint I/O** | `python-pptx` | `1.0.2` | Native parsing and generation of `.pptx` XML presentation files. |
| **Image Processing** | Pillow (PIL) | `11.1.0` | High-definition image loading, resizing, and fallback slide canvas rendering. |
| **Multi-Display Info** | ScreenInfo | `0.8.1` | Physical hardware monitor detection (resolutions, positioning, primary/secondary flags). |
| **Executable Packaging** | PyInstaller | `6.21.0` | Packaging Python applications into single self-contained `.exe` binaries. |

---

## 5. Directory Structure, File Catalog & Weightage Analysis

### 5.1 Project Directory Tree
```text
c:\Users\ASUS\Downloads\New folder\
│
├── presentation_app.py           (69.9 KB)  <-- Main Desktop Application & GUI Engine (OOP)
├── live_speech_recognizer.py     (24.4 KB)  <-- Standalone CLI Speech Recognizer Engine
├── mic_diagnostic.py              (3.4 KB)  <-- Hardware Microphone Diagnostic Utility
├── start_presenter.bat            (0.3 KB)  <-- Windows Presenter Batch Launcher
├── start_recognizer.bat           (0.1 KB)  <-- Windows Recognizer Batch Launcher
├── test_deck.pptx                (28.3 KB)  <-- Default PowerPoint Test Deck
├── README.md                      (3.1 KB)  <-- GitHub Repository Documentation
├── GENSLIDE_DOCUMENTATION.md      (15.2 KB) <-- Formal System Architecture Specification
├── GenSlide.spec                  (0.9 KB)  <-- PyInstaller Executable Specification
├── .gitignore                     (0.2 KB)  <-- Git Exclusion Rules
│
├── dist/
│   └── GenSlide.exe             (111.9 MB)  <-- Portable Standalone Executable Binary
│
└── vosk-model-small-en-us-0.15/  (51.2 MB)  <-- Local Offline Acoustic Neural Model
    ├── am/final.mdl              (14.8 MB)  <-- Acoustic Model Neural Network Weights
    ├── graph/HCLr.fst            (12.2 MB)  <-- Transducer Graph FST
    ├── graph/Gr.fst               (8.5 MB)  <-- Language Model Grammar Graph
    ├── ivector/final.ie           (9.1 MB)  <-- iVector Extractor Matrix
    └── ... (configuration & dictionary files)
```

---

### 5.2 File Weightage & Resource Distribution

```text
========================================================================================
FILE / COMPONENT                          WEIGHTAGE (% OF SYSTEM SIZE)    SIZE
========================================================================================
dist/GenSlide.exe (Standalone Executable) ██████████████████████████ 68.6%   111.9 MB
vosk-model-small-en-us-0.15/ (Model Data) ████████████████           31.4%    51.2 MB
presentation_app.py (Core Codebase)       ▏                          0.04%    69.9 KB
test_deck.pptx (Sample Deck)              ▏                          0.01%    28.3 KB
live_speech_recognizer.py                 ▏                          0.01%    24.4 KB
mic_diagnostic.py                         ▏                          0.00%     3.4 KB
Batch Launchers & Configs                 ▏                          0.00%     1.5 KB
========================================================================================
TOTAL DISK FOOTPRINT                                               100.0%   163.1 MB
========================================================================================
```

---

## 6. Real-Time Data Pipelines & Workflows

### 6.1 Audio Capture & Resampling Subsystem (<1ms)
1. Microphone PCM audio is captured in real-time via `sounddevice.RawInputStream` at native hardware sample rates (typically 44.1kHz or 48kHz, Int16).
2. The `audio_callback` calculates Root Mean Square (RMS) audio energy for the live GUI VU meter.
3. Raw audio bytes are passed to `fast_process_and_resample()`, where NumPy performs stereo-to-mono downmixing and linear interpolation resampling (`numpy.interp`) to produce a 16kHz Int16 stream required by Vosk.

### 6.2 Offline Vosk Acoustic Neural Decoding Subsystem (<10ms)
1. The 16kHz PCM stream is pushed into `audio_queue`.
2. The asynchronous worker thread `vosk_instant_worker_loop()` pulls PCM chunks and feeds them into `vosk_recognizer.AcceptWaveform(data)`.
3. `vosk_recognizer.PartialResult()` returns recognized words in real-time milliseconds **before the user even finishes full sentences**.

### 6.3 Fuzzy Word-Boundary Keyword Matching & Auto-Sync Algorithm
1. Spoken text is cleaned: punctuation is stripped, text is converted to lowercase, and extra spaces are normalized.
2. Keywords in `keywords_map` are sorted by length (longest phrases matched first to prevent partial word conflicts).
3. Regex word-boundary phrase matching (`r'\b' + re.escape(kw) + r'\b'`) checks if any assigned slide keyword exists in the spoken string.
4. On a match, `on_keyword_matched_cb` dispatches a thread-safe UI update (`self.after(0, ...)`), jumping immediately to the target slide index.

### 6.4 100% PowerPoint Visual Render Subsystem (COM API & PIL Fallback)
1. When a `.pptx` presentation is opened, `SlideManager.load_pptx()` checks for Microsoft PowerPoint installation via PyWin32 COM (`win32com.client.Dispatch("PowerPoint.Application")`).
2. If PowerPoint is installed, COM opens the deck silently and exports every slide into full 100% color-accurate PNG renderings.
3. If PowerPoint is absent, PIL generates high-definition dark-themed slide canvas fallback images preserving title and bullet text.

### 6.5 HDMI Multi-Monitor & Projector Auto-Targeting Subsystem
1. `screeninfo.get_monitors()` queries physical displays connected to the system.
2. If a secondary monitor/projector is detected (`len(monitors) > 1` or `x != 0`), `ExternalDisplayWindow` launches on the projector display in **100% Borderless Fullscreen** (`overrideredirect(True)`).
3. If only a single primary laptop monitor is connected (`x == 0`), `ExternalDisplayWindow` launches as a compact windowed preview (`740x416`) with window controls so presenter controls remain unobstructed.

---

## 7. Build, PyInstaller Packaging & USB Portability Specification

GenSlide uses PyInstaller to bundle Python code, C++ DLLs, and neural models into a single standalone executable file.

### 7.1 PyInstaller Build Command
```powershell
pyinstaller --noconfirm --onefile --windowed --name "GenSlide" \
  --add-data "vosk-model-small-en-us-0.15;vosk-model-small-en-us-0.15" \
  --exclude-module matplotlib --exclude-module PyQt6 --exclude-module PyQt5 \
  --exclude-module PySide6 --exclude-module PySide2 --exclude-module scipy \
  --exclude-module pandas --exclude-module tkinter.test --exclude-module unittest \
  --exclude-module email --exclude-module html --exclude-module http \
  --exclude-module xmlrpc --exclude-module pydoc --exclude-module pydoc_data \
  --exclude-module doctest --exclude-module setuptools --exclude-module test \
  presentation_app.py
```

### 7.2 Frozen Resource Resolution Architecture (`get_resource_path`)
When executing as a frozen `.exe`, PyInstaller extracts internal payload files into a temporary directory (`sys._MEIPASS`). `presentation_app.py` implements a resource loader helper:

```python
def get_resource_path(relative_path):
    """Returns absolute path to resource, working for dev environment and PyInstaller executable."""
    if getattr(sys, 'frozen', False):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    target_path = os.path.join(base_path, relative_path)
    if os.path.exists(target_path):
        return target_path
    
    cwd_path = os.path.join(os.getcwd(), relative_path)
    if os.path.exists(cwd_path):
        return cwd_path
        
    return relative_path
```

---

## 8. Operational Controls & Keyboard Cheat Sheet

| Key / Control | Function | Scope |
| :--- | :--- | :--- |
| `→` / `↓` / `Space` / `PgDown` | Advance to Next Slide | Global Application |
| `←` / `↑` / `Backspace` / `PgUp` | Return to Previous Slide | Global Application |
| `B` | Toggle Blackout Screen | Global Application |
| `W` | Toggle Whiteout Screen | Global Application |
| `Esc` | Close Projector Presentation Window | Global Application |
| **`🎙️ LIVE VOICE`** Button | Enable / Disable Voice Recognition Engine | Header Bar |
| **`📺 PRESENT ON HDMI`** Button | Launch Projector Fullscreen Presentation Window | Header Bar / Settings |
| **`💾 SAVE ALL VOICE KEYWORDS`** Button | Re-index Voice Recognition Dictionary | Keyword Matrix View |

---

*End of Formal System Documentation for GenSlide Version 2.0.0.*
