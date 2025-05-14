# Voice-to-Command Application

A C# Windows Forms Application (.NET Framework) that enables **hands-free interaction** with your computer by converting voice input into executable system commands.

---

## Overview

This application listens to the user's **voice commands** and performs corresponding actions such as:

- **System Power Commands**: shutdown, restart, log off
- **Application Launching**: open browsers, notepad, calculator, etc.
- **Media Control**: increase/decrease volume, mute/unmute
- **Monitor Control**: turn off screen, lock screen
- **And other more commands**
- **Interface Features**:
  - "Always on top" mode
  - Minimize to system tray and continue listening in the background
  - Lightweight and user-friendly interface

---

## Technologies Used

- **Language**: C#
- **Framework**: .NET Framework (Windows Forms)
- **Speech Recognition**: System.Speech or Whisper integration

---

## Features

| Feature                    | Description                                               |
|----------------------------|-----------------------------------------------------------|
|    Voice Recognition       | Captures and transcribes user voice to text               |
|    Command Execution       | Executes matching system-level or app-level commands      |
|    Volume Management       | Adjust system sound volume or mute                        |
|    Display Actions         | Lock, turn off, or sleep monitor                          |
|    Other actions           | various ammount of commands
|    Background Mode         | Runs silently in background and responds to commands      |
|    Always on Top Mode      | Stays visible above other windows if enabled              |


---

## 🚀 How to Run

1. Clone the repository
2. Open the solution `.sln` in Visual Studio
3. Build the project (`Ctrl + Shift + B`)
4. Run the application (`F5`)
5. Speak a command like `"shutdown"` or `"open Chrome"`

---

## Author

Developed by **Mohammad nafiz sammour**

---

## License

This application is open-source and free to use. If you modify or redistribute it, please include credit or contact the author.

---

## Future Improvements

- Wake word support (e.g., "Hey PC")
- Command learning with ML.NET
- Language localization
- Scheduled voice tasks

---



