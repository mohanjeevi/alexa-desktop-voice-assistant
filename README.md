# 🎙️ Voice-Controlled Desktop Assistant (Alexa-Triggered)

This Python-based desktop assistant allows you to control your system using voice commands prefixed with "Alexa". It supports actions like opening/closing apps and folders, controlling volume/brightness, taking screenshots, and shutting down or restarting your PC.

---

## 🚀 Features

- 🎤 Voice command recognition with wake word **"Alexa"**
- 🗣️ Text-to-speech feedback
- 🗂️ Open and close popular apps and folders
- 🔊 Volume control (mute, unmute, set %)
- 💡 Brightness control
- 🖼️ Take screenshots
- 🖥️ System commands: shutdown, restart, lock
- 🪟 Close all folders/windows

---

## 🛠️ Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/mohanjeevi/alexa-desktop-voice-assistant.git
   cd alexa-desktop-voice-assistant
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the assistant**:
   ```bash
   python app.py
   ```

---

## 🧠 Usage Instructions

1. Speak clearly after saying the wake word **"Alexa"**.
2. Example commands:
   - `"Alexa open Chrome"`
   - `"Alexa close Notepad"`
   - `"Alexa set volume to 60%"`
   - `"Alexa shutdown"`
   - `"Alexa take a screenshot"`

> **Note:** Make sure your microphone is connected and working.

---

## 🧩 Supported Apps & Paths

The assistant supports opening:
- Word, Excel, PowerPoint, Notepad, Paint, Calculator, Chrome, Firefox, Edge, VS Code, etc.

---

## 🖥️ Platform Support

- ✅ Windows
- ✅ macOS
- ✅ Linux

---

## 📂 Folder Structure

```
desktop-voice-assistant/
│
├── main.py               # Main script with all voice assistant logic
├── requirements.txt      # Dependencies
└── README.md             # This file
```

---

## 📸 Screenshots

Captured screenshots will be saved to your **Pictures** folder as:
```
screenshot1.png, screenshot2.png, ...
```

---