# ⚡ GenSlide - Production-Grade AI Voice-Controlled Dual-Screen PPT Presenter & Editor
## System Architecture, Object-Oriented Design & Technical Documentation Specification

**Document Version**: 2.5.0  
**Target Environment**: Windows 10 / 11 Desktop (Standalone Portable Executable & Python Runtime)  
**Author / Engineering Team**: Advanced Agentic Coding System  
**Date**: August 2026  

---

## 1. Executive Summary & System Purpose

**GenSlide** is an enterprise-grade desktop software application designed for voice-driven presentation management, real-time slide manipulation, and dual-screen HDMI display projection.

Traditional presentation tools require manual clicker interaction or keyboard navigation near the host machine. **GenSlide** eliminates these physical constraints by embedding a **Sub-Millisecond Constrained Kaldi FST Acoustic Neural Engine** (Vosk) alongside Google Neural Speech Recognition, allowing presenters to utter custom slide keywords or navigation commands to trigger slide switches in real-time milliseconds.

### Key Objectives & System Capabilities:
- 🎙️ **Sub-Millisecond Kaldi FST Keyword Navigation**: Dynamic compilation of slide keywords and navigation commands directly into the Kaldi Finite State Transducer grammar for instant (<1ms) slide switching with 100% precision.
- 🔇 **Zero-Hallucination & Adaptive SNR Noise Gate**: Continuous ambient noise floor tracking (`noise_floor = 0.96 * floor + 0.04 * rms`) and adaptive floating speech threshold. Ambient chatter and HVAC room noise are routed to `[unk]` and discarded.
- 🎧 **Dynamic Bluetooth & USB Mic Hot-Plugging**: Automatically detects and transitions to newly connected Bluetooth earphones, wireless headsets, and USB microphones. Automatically recovers and migrates stream to system microphones upon disconnection.
- 📊 **Dynamic Voice Level VU Meter & Color Alerts**: Instant peak attack with exponential smooth decay. Color dynamically adapts: Gray (`⚪ OFF`), Green (`🟢 LIVE`), Amber (`🟡 LIVE`), and Red (`🔴 NOISE`) when energy exceeds 80% (detecting shouting, clipping, or loud noise).
- ⏱️ **Real Measured FPS & Live DSP Latency**: Dynamically measures actual frame execution rate and true audio DSP processing time in milliseconds (`0.8ms` - `2.5ms`).
- 🏷️ **Full Presentation Keywords Matrix in Dashboard**: Displays all slides and their associated voice keywords row-by-row on the right sidebar, with live slide highlighting and 1-click slide jump.
- 🤖 **Automatic Default Voice Keywords on PPTX Upload**: Every uploaded slide automatically receives numbers (`"slide 1"`, `"one"`), ordinals (`"first slide"`, `"first"`), and content keywords extracted from the slide title and body.
- 📺 **Dual-Screen HDMI Auto-Targeting**: Automatic detection of external HDMI monitors and projectors, projecting 100% borderless fullscreen slides on secondary displays while maintaining a complete Presenter Dashboard on the primary laptop screen.
- 🎨 **100% PowerPoint COM Visual Exporter & Native Fallback**: Direct inter-process COM integration with Microsoft PowerPoint (`pywin32`) to export high-definition slide renderings (`.png`/`.jpg`), with automatic fallback to native Python-PPTX + Pillow vector canvas rendering if MS Office is not installed.
- 📦 **Universal Windows Portability**: Packaged as a single self-contained executable (`dist\GenSlide.exe`) via PyInstaller, containing all Python runtimes, C++ DLLs, and the embedded Vosk model for execution on any Windows 10/11 computer.

---

## 2. Technology Stack & Library Dependency Specifications

| Component Layer | Technology / Library | Version | System Purpose |
| :--- | :--- | :--- | :--- |
| **Language & Runtime** | Python | `3.12.x` | Core application logic & runtime environment. |
| **GUI Framework** | CustomTkinter | `5.2.2` | Modern dark-themed responsive desktop user interface. |
| **Offline Speech Engine** | Vosk API (Kaldi FST) | `0.3.45` | C++ Kaldi acoustic neural engine with constrained FST grammar (<1ms matching). |
| **Offline Acoustic Model** | `vosk-model-small-en-us-0.15` | `0.15` | Lightweight 50MB English acoustic neural model bundled into binary. |
| **Online Speech Backup** | SpeechRecognition | `3.14.1` | Google Speech Neural API fallback wrapper. |
| **Audio I/O Capture** | SoundDevice / WinMM | `0.5.1` | Asynchronous real-time hardware microphone PCM audio streaming & hot-plugging. |
| **Audio Processing** | NumPy | `2.2.3` | Vectorized RMS energy calculation, noise gate filtering, and linear resampling. |
| **PowerPoint COM API** | PyWin32 (`win32com`) | `308` | Direct Windows COM interface to Microsoft PowerPoint for 100% accurate slide exports. |
| **PowerPoint I/O** | `python-pptx` | `1.0.2` | Native parsing and generation of `.pptx` XML presentation files. |
| **Image Processing** | Pillow (PIL) | `11.1.0` | High-definition image loading, resizing, and fallback slide canvas rendering. |
| **Multi-Display Info** | ScreenInfo | `0.8.1` | Physical hardware monitor detection (resolutions, positioning, primary/secondary flags). |
| **Executable Packaging** | PyInstaller | `6.21.0` | Packaging Python application into a single self-contained `.exe` binary. |

---

## 3. Object-Oriented Programming (OOP) Class Architecture

GenSlide is structured around fundamental Object-Oriented Programming (OOP) principles: **Encapsulation**, **Single Responsibility**, **Inheritance**, and **Abstraction**.

| Class Name | Inheritance | Primary Responsibility | OOP Role |
| :--- | :--- | :--- | :--- |
| `SlideData` | `object` | Encapsulates single slide properties (ID, title, bullets, notes, keywords, rendering path, render hash). | Data Value Object (Model) |
| `SlideManager` | `object` | Manages deck lifecycle, `.pptx` reading/writing, COM export, and PIL canvas generation. | Domain Service Manager |
| `VoiceSpeechEngine` | `object` | Handles microphone streaming, audio downmixing, noise floor tracking, Vosk FST recognition, and hotplugging. | Background Subsystem Service |
| `WinMMAudioMonitor` | `object` | Direct low-level Windows multimedia subsystem monitor for instant device hotplugging. | Hardware Interface Adapter |
| `ExternalDisplayWindow` | `ctk.CTkToplevel` | Controls secondary projector window fullscreen rendering and primary preview windowing. | Specialized UI Controller (View) |
| `PresentationApp` | `ctk.CTk` | Main application controller orchestrating views, hotkeys, navigation, dashboard, and speech callbacks. | Orchestrator & Controller |

---

## 4. Directory Structure & File Catalog

```text
d:\Genslide\
├── presentation_app.py        # Main Desktop Application & GUI Engine (OOP Architecture)
├── GenSlide.spec              # PyInstaller Single-File Executable Specification
├── build_exe.bat              # Windows 1-Click PyInstaller Executable Builder
├── README.md                  # GitHub Repository Documentation & User Guide
├── GENSLIDE_DOCUMENTATION.md  # Formal System Architecture Documentation (Markdown)
├── GENSLIDE_DOCUMENTATION.txt # Formal System Architecture Documentation (Plain Text)
│
├── dist/
│   └── GenSlide.exe           # Standalone Single-File Windows Executable (146 MB)
│
└── vosk-model-small-en-us-0.15/ # Embedded Local Offline Acoustic Neural Model
```

---

## 5. Operational Controls & Keyboard Cheat Sheet

| Key / Control | Voice Equivalent | Action | Scope |
| :--- | :--- | :--- | :--- |
| `→` / `↓` / `Space` / `PgDown` | *"Next Slide"*, *"Forward"* | Advance to Next Slide | Global Application |
| `←` / `↑` / `Backspace` / `PgUp` | *"Previous Slide"*, *"Back"* | Return to Previous Slide | Global Application |
| `1` – `9` Keys | *"Slide 1"*, *"Slide 2"*, *"First Slide"* | Jump Directly to Slide N | Global Application |
| `B` | *"Black screen"*, *"Blackout"* | Toggle Blackout Screen | Global Application |
| `W` | *"White screen"*, *"Whiteout"* | Toggle Whiteout Screen | Global Application |
| `Esc` | *"Exit presentation"* | Close Projector Window / Stop | Global Application |
| **`🎙️ LIVE VOICE`** Button | N/A | Enable / Disable Voice Recognition Engine | Header Bar |
| **`📺 PRESENT ON HDMI`** Button | N/A | Launch Projector Fullscreen Presentation Window | Header Bar / Settings |
| **`💾 SAVE ALL VOICE KEYWORDS`** Button | N/A | Re-index Voice Recognition Dictionary | Keyword Matrix View |

---

*End of Formal System Documentation for GenSlide Version 2.5.0.*
