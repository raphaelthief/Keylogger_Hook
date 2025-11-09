# Keylogger_Hook
This VB.NET WinForms application implements a low-level keyboard and input monitoring system. It captures keystrokes, modifier states, clipboard changes, active window context, and special key combinations in real time. The architecture emphasizes efficient, non-blocking hooks, accurate character resolution, and robust logging.

## Disclaimer
This software is a keylogger and input monitoring tool. It is capable of capturing keystrokes, keyboard shortcuts, clipboard content, window titles, and special key combinations. Use of this program can violate privacy laws and may be considered illegal if used without the explicit consent of all parties involved.
This code is provided for educational, research, or authorized administrative purposes only. The author and distributor do not accept any responsibility or liability for misuse of this software.
By using, running, or distributing this software, you assume full responsibility for compliance with local, national, and international laws.


## Program Customization & Live Display
The commented parts of the code can directly affect the program’s behavior, particularly regarding how text is displayed in the textbox, including truncated sections or the display of pressed keys.
Here, the save frequency (and consequently the clearing of the textbox) can be customized in this part of the code:

![Save_frequency](https://github.com/raphaelthief/Keylogger_Hook/blob/main/Pictures/textbox_display_and_save.png "Save_frequency")

The save options can be customized in these declarations:

![Save_options](https://github.com/raphaelthief/Keylogger_Hook/blob/main/Pictures/Save_options.png "Save_options")

Here is an example of how periodic saving works based on the configured settings:

![Files_saved](https://github.com/raphaelthief/Keylogger_Hook/blob/main/Pictures/files_saved.png "files_saved")

Here is an example of live log display in the textbox:

![Working](https://github.com/raphaelthief/Keylogger_Hook/blob/main/Pictures/Working.png "Working")


## 1. Windows API Integration
#### The application heavily relies on Win32 API functions to interact with the operating system at a low level:
- Keyboard Hook & Event Capture
- SetWindowsHookEx with WH_KEYBOARD_LL installs a global low-level keyboard hook.
- CallNextHookEx ensures proper chain execution.
- UnhookWindowsHookEx removes the hook when no longer needed.
- HookCallback processes WM_KEYDOWN and WM_SYSKEYDOWN events.
#### Keyboard State & Character Translation
- GetAsyncKeyState detects whether keys are currently pressed.
- GetKeyboardState retrieves a snapshot of all key states.
- ToUnicode and ToUnicodeEx convert virtual key codes to Unicode characters, taking into account Shift, CapsLock, AltGr, and keyboard layout.
- MapVirtualKey converts virtual key codes to scan codes.
- GetKeyState checks toggle key states (CapsLock, NumLock, ScrollLock).
#### Window Context
- GetForegroundWindow retrieves the handle of the active window.
- GetWindowText fetches the active window title for context-aware logging.


## 2. Keyboard Event Handling
#### Core Hook Logic
- StartHook initializes the low-level hook.
- StopHook removes the hook safely.
- SetHook binds the callback to the current process module.
#### Key Processing (ProcessKey)
- Checks for modifier keys: Ctrl, Alt, Shift, and Windows.
- Supports detection of AltGr (Right Ctrl + Right Alt) to avoid misinterpreting combinations.
- Updates toggle states for CapsLock, NumLock, ScrollLock.
- Resolves the actual character using GetActualCharacter with proper CapsLock and Shift handling.
- Supports special key logging (function keys, navigation keys, multimedia keys).
#### Special Keys & Combinations (CheckSpecialKeys)
- Tracks repeated key combinations using Static variables.
- Detects shortcuts:
  - Common: Ctrl+C, Ctrl+V, Ctrl+S, Ctrl+Shift+T
  - Windows: Win+D, Win+E, Win+L
  - System: Alt+Tab, Alt+F4
- Detects mouse clicks (Right and Middle).
- Handles Backspace sequences intelligently with a timing mechanism to aggregate presses.
- Maintains a HashSet of active combinations to avoid duplicate logging.


## 3. Clipboard Monitoring
- CheckClipboardChange periodically polls the clipboard every 3 seconds.
- Tracks clipboard text, images, and file drops.
- Associates clipboard events with user actions (e.g., Ctrl+C vs. right-click copy).
- Maintains lastClipboardText and lastCopyTime to avoid redundant entries.


## 4. Active Window Tracking
- GetActiveWindowTitle fetches the current foreground window’s title.
- The application logs window context changes, ensuring keystrokes are associated with the correct window.
- Detects window switches and flushes accumulated text accordingly.


## 5. Logging and File Management
- Text is displayed in a TextBox (TextBox1) with dynamic scrolling and buffer updates via BeginInvoke.
- Implements real-time logging with buffer size management:
  - AppendText appends to UI safely across threads.
  - SaveToFile writes buffer to a rotating log file system.
  - Log files are saved in %TEMP%\logsK by default, with a configurable maximum file size (default 0.1 MB).
  - Automatic file rotation ensures logs do not exceed defined thresholds.
- Each session starts with a timestamp and lock key status.
- Text buffer flushing occurs:
  - On window change
  - When TextBox exceeds 350 characters
  - On form closing


## 6. Timer-Based Updates
- Timer1 ticks every 100 ms to:
  - Check active window title changes.
  - Poll clipboard content.
  - Detect special key combinations and shortcuts.
- Ensures non-blocking real-time monitoring without overwhelming CPU resources.


## 7. Unicode & Modifier Handling
- GetActualCharacter accurately converts key codes to characters:
  - Handles Shift, CapsLock, AltGr, and keyboard layout.
  - Correctly interprets letters, numbers, symbols, and punctuation.
- UpdateLockStates and IsCapsLockActive maintain proper toggle key state tracking.
- Prevents duplicate logging for repeated key presses within 350 ms.


## 8. Form & UI
- Form1_Load initializes UI and starts monitoring.
- Uses a multiline TextBox with vertical scrolling and monospaced font (Consolas).
- Form closing triggers final buffer flush and hook removal.


## 9. Design Highlights
- Low-Level Efficiency: Uses Windows hooks for minimal latency.
- Accurate Character Resolution: Converts virtual key codes into actual Unicode text respecting modifier keys.
- Context Awareness: Associates keystrokes with window title and clipboard events.
- Event Deduplication: Avoids repeated logging for held keys or simultaneous shortcuts.
- Logging System: Rotating files prevent memory overflow and enable persistent data storage.
