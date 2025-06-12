# Voice Based Volume Controller

## Overview

This Python script enables users to control the system audio volume on Windows using voice commands or a keyboard shortcut. It leverages the `pycaw` library to interact with the Windows Core Audio API and the `speech_recognition` library to process voice input via a microphone. The script supports commands to increase, decrease, or set the volume to a specific percentage, and it can be terminated with a keyboard shortcut.

## Features

- Recognizes voice commands: "volume up", "volume down", and "set volume to <number>" (e.g., "set volume to 60").
- Sets the master volume of the default audio device (e.g., speakers or headphones).
- Exits the program when the 'q' key is pressed.
- Provides feedback on recognized commands or when no speech is detected.

## Requirements

- **Operating System**: Windows 7 or later (Windows 10/11 recommended).
- **Python**: Version 3.6 or higher (tested with Python 3.12).
- **Dependencies**:
  - `pycaw`: For controlling Windows audio.
  - `comtypes`: For COM interface interactions.
  - `pywin32`: For COM initialization.
  - `speech_recognition`: For voice recognition.
  - `keyboard`: For detecting keyboard input.
- **Hardware**: A working microphone and an active audio output device (e.g., speakers or headphones).
- **Internet Connection**: Required for Google Speech Recognition API.

## Installation

1. **Install Python**:
   - Download and install Python 3.6+ from [python.org](https://www.python.org/downloads/).
   - Ensure `pip` is available.

2. **Install Dependencies**:
   - Run the following command to install required libraries:
     ```bash
     pip install pycaw comtypes pywin32 speechrecognition keyboard
     ```

3. **Verify Audio Devices**:
   - Ensure a microphone and an audio output device are connected and set as default in Windows Sound settings (Control Panel > Sound).

## Usage

1. **Save the Script**:
   - Save the script as `voice_volume_control.py`.

2. **Run the Script**:
   - Open a terminal or command prompt.
   - Navigate to the script's directory.
   - Run the script:
     ```bash
     python voice_volume_control.py
     ```
   - If permission issues occur, run the terminal as administrator.

3. **Voice Commands**:
   - Say "volume up" to set the volume to 100%.
   - Say "volume down" to set the volume to 10%.
   - Say "set volume to <number>" (e.g., "set volume to 60") to set the volume to a specific percentage (0-99).
   - The script will print feedback, e.g., `Command: "set volume to 60" Done` or `Nothing Heard!` if no speech is detected.

4. **Exit the Script**:
   - Press the `q` key to terminate the program (`Ended` will be printed).

## Code Explanation

- **COM Initialization**: Uses `pythoncom.CoInitialize()` and `CoUninitialize()` to manage COM resources for `pycaw`.
- **Audio Control**: Uses `pycaw` to access the default audio device's `IAudioEndpointVolume` interface for volume control.
- **Speech Recognition**: Uses `speech_recognition` to listen for voice commands via the default microphone, processed by Google's Speech Recognition API.
- **Command Processing**:
  - Detects "volume up" and "volume down" to set fixed volume levels (100% and 10%, respectively).
  - Uses a regex pattern (`set volume to (\d{1,2})`) to parse percentage values for "set volume to" commands.
- **Keyboard Input**: Uses the `keyboard` library to detect the `q` key for exiting.
- **Error Handling**: Catches exceptions for failed speech recognition, printing `Nothing Heard!` when no valid audio is detected.

## Notes

- **Ambient Noise**: The commented-out `recognizer.adjust_for_ambient_noise()` can be enabled to improve speech detection in noisy environments.
- **Limitations**:
  - The regex pattern only matches 1- or 2-digit numbers (e.g., "set volume to 100" is not supported).
  - Requires an active internet connection for Google Speech Recognition.
  - The `keyboard` library may require administrator privileges on some systems.
- **Troubleshooting**:
  - Ensure the microphone and audio output device are set as default in Windows Sound settings.
  - Check the internet connection for Google Speech Recognition.
  - Run the script as administrator if COM or microphone access fails.
  - Update dependencies with `pip install --upgrade pycaw comtypes pywin32 speechrecognition keyboard`.

## Example Output

```
Listening
set volume to 60
Command: "set volume to 60" Done
Listening
Nothing Heard!
Listening
Ended
```

## License

This script is provided as-is, with no warranty. Use and modify it freely under the MIT License (if applicable to dependencies).

## Acknowledgments

- Built using the `pycaw` library ([GitHub](https://github.com/AndreMiras/pycaw)).
- Uses Google Speech Recognition API via `speech_recognition`.
- Inspired by the need for hands-free audio control on Windows.
