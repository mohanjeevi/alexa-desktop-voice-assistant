import os
import subprocess
import re
import platform
import psutil
import logging
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pythoncom  # Required for COM initialization on Windows
import mss  # For taking screenshots
import mss.tools
import screen_brightness_control as sbc  # For controlling screen brightness
import speech_recognition as sr  # For voice recognition
import pyttsx3  # For text-to-speech

# Configure logging
logging.basicConfig(level=logging.INFO)  # Set logging level to INFO to suppress DEBUG messages

# Suppress comtypes debug logs
comtypes_logger = logging.getLogger('comtypes')
comtypes_logger.setLevel(logging.WARNING)

# Initialize the text-to-speech engine
engine = pyttsx3.init()

def speak(text):
    """Convert text to speech"""
    engine.say(text)
    engine.runAndWait()

def normalize_drive_name(folder_name):
    """Normalize misheard words to correct drive names"""
    folder_name = folder_name.lower()
    if folder_name in ["see drive", "seed drive", "se drive", "see"]:
        return "c"
    return folder_name

def open_folder(folder_name):
    if not folder_name:
        return "Folder name not provided."

    folder_name = normalize_drive_name(folder_name)

    # Handle cases like "open C" or "open D"
    if re.fullmatch(r"[a-z]", folder_name, re.I):
        drive_letter = folder_name.upper() + ":\\"
        if os.path.exists(drive_letter):
            folder_path = drive_letter
        else:
            return f"Drive '{drive_letter}' does not exist."
    else:
        drive_match = re.match(r"([a-z])\s*drive", folder_name, re.I)
        if drive_match:
            drive_letter = drive_match.group(1).upper() + ":\\"
            if os.path.exists(drive_letter):
                folder_path = drive_letter
            else:
                return f"Drive '{drive_letter}' does not exist."
        else:
            folder_path = os.path.join(os.path.expanduser('~'), folder_name)

    if not os.path.exists(folder_path):
        return f"Folder '{folder_path}' does not exist."

    try:
        system_name = platform.system()
        if system_name == "Windows":
            os.startfile(folder_path)
        elif system_name == "Darwin":  # macOS
            subprocess.run(['open', folder_path])
        else:  # Linux
            subprocess.run(['xdg-open', folder_path])
        return f"Opened: {folder_path}"
    except Exception as e:
        return f"Error: {str(e)}"

def close_folder(folder_name):
    folder_name = normalize_drive_name(folder_name)

    if folder_name.lower() == "all folders":
        if platform.system() == "Windows":
            try:
                import pygetwindow as gw
                all_windows = gw.getAllTitles()
                explorer_titles = [
                    "File Explorer", "This PC", "Quick access", 
                    "Documents", "Downloads", "Pictures", "Videos", "Music", "Desktop"
                ]
                explorer_windows = [win for win in all_windows if any(title.lower() in win.lower() for title in explorer_titles) or ":" in win]
                for win_title in explorer_windows:
                    try:
                        window = gw.getWindowsWithTitle(win_title)[0]
                        window.close()
                    except Exception as e:
                        logging.error(f"Error closing window: {win_title}, {str(e)}")
                return "Closed all open File Explorer windows."
            except ImportError:
                return "pygetwindow is not supported on this platform."
        else:
            return "Closing all folders is only supported on Windows."
    else:
        if platform.system() == "Windows":
            try:
                import pygetwindow as gw
                explorer_windows = gw.getWindowsWithTitle(folder_name)
                for window in explorer_windows:
                    window.close()
                return f"Closed: {folder_name}"
            except ImportError:
                return "pygetwindow is not supported on this platform."
        else:
            return "Closing specific folders is only supported on Windows."

def manage_app(app_name, action):
    """Opens or closes applications"""
    apps = {
        "word": {
            "open": r"C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
            "close": "WINWORD.EXE"
        },
        "excel": {
            "open": r"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
            "close": "EXCEL.EXE"
        },
        "powerpoint": {
            "open": r"C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
            "close": "POWERPNT.EXE"
        },
        "notepad": {
            "open": "notepad.exe",
            "close": "notepad.exe"
        },
        "calculator": {
            "open": "calc.exe",
            "close": "ApplicationFrameHost.exe"
        },
        "paint": {
            "open": "mspaint.exe",
            "close": "mspaint.exe"
        },
        "edge browser": {
            "open": r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
            "close": "msedge.exe"
        },
        "chrome browser": {
            "open": r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            "close": "chrome.exe"
        },
        "firefox browser": {
            "open": "firefox.exe",
            "close": "firefox.exe"
        },
        "vlc": {
            "open": r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\vlc.exe",
            "close": "vlc.exe"
        },
        "spotify": {
            "open": "Spotify.exe",
            "close": "Spotify.exe"
        },
        "adobe reader": {
            "open": "AcroRd32.exe",
            "close": "AcroRd32.exe"
        },
        "command prompt": {
            "open": "cmd.exe",
            "close": "cmd.exe"
        },
        "task manager": {
            "open": "Taskmgr.exe",
            "close": "Taskmgr.exe"
        },
        "windows media player": {
            "open": "wmplayer.exe",
            "close": "wmplayer.exe"
        },
        "photos": {
            "open": "Microsoft.Photos.exe",
            "close": "ApplicationFrameHost.exe"
        },
        "vs code": {
            "open": r"C:\\Users\\%USERNAME%\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
            "close": "Code.exe"
        },
        "visual studio code": {
            "open": r"C:\\Users\\%USERNAME%\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
            "close": "Code.exe"
        },
        "code": {
            "open": r"C:\\Users\\%USERNAME%\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
            "close": "Code.exe"
        }
    }
    
    if app_name not in apps:
        return f"Application '{app_name}' not recognized."

    try:
        if action == "open":
            # Replace %USERNAME% with actual username for VS Code path
            open_path = apps[app_name]["open"]
            if "%USERNAME%" in open_path:
                open_path = open_path.replace("%USERNAME%", os.getenv("USERNAME"))
            subprocess.Popen(open_path, shell=True)
            return f"Opened {app_name.capitalize()}"
        elif action == "close":
            process_name = apps[app_name]["close"]
            for process in psutil.process_iter(attrs=['pid', 'name']):
                if process.info['name'].lower() == process_name.lower():
                    os.kill(process.info['pid'], 9)
            return f"Closed {app_name.capitalize()}"
        else:
            return f"Invalid action '{action}'."
    except Exception as e:
        return f"Error {action}ing {app_name}: {str(e)}"

def shutdown_computer():
    try:
        if platform.system() == "Windows":
            os.system("shutdown /s /t 1")
        elif platform.system() == "Darwin" or platform.system() == "Linux":
            os.system("shutdown -h now")
        return "Shutting down the computer..."
    except Exception as e:
        return f"Error shutting down the computer: {str(e)}"

def restart_computer():
    try:
        if platform.system() == "Windows":
            os.system("shutdown /r /t 1")
        elif platform.system() == "Darwin" or platform.system() == "Linux":
            os.system("reboot")
        return "Restarting the computer..."
    except Exception as e:
        return f"Error restarting the computer: {str(e)}"

def lock_computer():
    try:
        if platform.system() == "Windows":
            os.system("rundll32.exe user32.dll,LockWorkStation")
        elif platform.system() == "Darwin":
            os.system("pmset displaysleepnow")
        elif platform.system() == "Linux":
            os.system("gnome-screensaver-command -l")
        return "Locking the computer..."
    except Exception as e:
        return f"Error locking the computer: {str(e)}"

def set_volume_windows(volume_level):
    """Set the system volume on Windows (0.0 to 1.0)"""
    try:
        pythoncom.CoInitialize()  # Initialize COM
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(volume_level, None)
        return f"Volume set to {int(volume_level * 100)}%"
    except Exception as e:
        return f"Error setting volume: {str(e)}"
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM

def set_volume_macos(volume_level):
    """Set the system volume on macOS (0 to 100)"""
    try:
        volume_level = int(volume_level * 100)
        subprocess.run(['osascript', '-e', f'set volume output volume {volume_level}'])
        return f"Volume set to {volume_level}%"
    except Exception as e:
        return f"Error setting volume: {str(e)}"

def set_volume_linux(volume_level):
    """Set the system volume on Linux (0 to 100)"""
    try:
        volume_level = int(volume_level * 100)
        subprocess.run(['amixer', 'set', 'Master', f'{volume_level}%'])
        return f"Volume set to {volume_level}%"
    except Exception as e:
        return f"Error setting volume: {str(e)}"

def adjust_volume(command):
    """Adjust the system volume based on the command"""
    # Handle mute/unmute
    if "mute" in command or "unmute" in command:
        if platform.system() == "Windows":
            try:
                pythoncom.CoInitialize()  # Initialize COM
                devices = AudioUtilities.GetSpeakers()
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                volume = cast(interface, POINTER(IAudioEndpointVolume))
                
                if "unmute" in command:
                    volume.SetMute(0, None)  # 0 = Unmute
                    return "Unmuted"
                else:
                    volume.SetMute(1, None)  # 1 = Mute
                    return "Muted"
            except Exception as e:
                return f"Error adjusting mute: {str(e)}"
            finally:
                pythoncom.CoUninitialize()  # Uninitialize COM
        elif platform.system() == "Darwin":
            if "unmute" in command:
                subprocess.run(['osascript', '-e', 'set volume output muted false'])
                return "Unmuted"
            else:
                subprocess.run(['osascript', '-e', 'set volume output muted true'])
                return "Muted"
        elif platform.system() == "Linux":
            if "unmute" in command:
                subprocess.run(['amixer', 'set', 'Master', 'unmute'])
                return "Unmuted"
            else:
                subprocess.run(['amixer', 'set', 'Master', 'mute'])
                return "Muted"
        else:
            return "Mute/Unmute is only supported on Windows, macOS, and Linux."

    # Handle volume up/down/set
    volume_match = re.search(r"(increase|decrease|set)\s*volume\s*(?:to)?\s*(\d+)%?", command)
    if volume_match:
        action = volume_match.group(1)
        volume_level = int(volume_match.group(2)) / 100
        if platform.system() == "Windows":
            return set_volume_windows(volume_level)
        elif platform.system() == "Darwin":
            return set_volume_macos(volume_level)
        elif platform.system() == "Linux":
            return set_volume_linux(volume_level)
        else:
            return "Volume control is only supported on Windows, macOS, and Linux."
    
    return "Volume command not recognized."

def set_brightness(brightness_level):
    """Set the screen brightness (0 to 100)"""
    try:
        sbc.set_brightness(brightness_level)
        return f"Brightness set to {brightness_level}%"
    except Exception as e:
        return f"Error setting brightness: {str(e)}"

def take_screenshot():
    """Take a screenshot and save it to the desktop with a unique name"""
    try:
        desktop_path = os.path.join(os.path.expanduser('~'), 'Pictures')
        base_name = "screenshot"
        extension = ".png"
        counter = 1

        # Find the next available filename
        while True:
            screenshot_path = os.path.join(desktop_path, f"{base_name}{counter}{extension}")
            if not os.path.exists(screenshot_path):
                break
            counter += 1

        # Capture the screenshot
        with mss.mss() as sct:
            screenshot = sct.grab(sct.monitors[1])
            mss.tools.to_png(screenshot.rgb, screenshot.size, output=screenshot_path)
        
        return f"Screenshot saved to {screenshot_path}"
    except Exception as e:
        return f"Error taking screenshot: {str(e)}"

def listen_for_command():
    """Listen for voice commands using the microphone"""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening for a command...")
        try:
            recognizer.adjust_for_ambient_noise(source)  # Adjust for ambient noise
            audio = recognizer.listen(source, timeout=5)  # Listen for up to 5 seconds
        except sr.WaitTimeoutError:
            return None  # Return None when no speech is detected

    try:
        command = recognizer.recognize_google(audio)  # Use Google Speech Recognition
        print(f"You said: {command}")
        return command.lower()
    except sr.UnknownValueError:
        return None  # Return None when speech is unintelligible
    except sr.RequestError:
        return "Sorry, there was an issue with the speech recognition service."

def handle_command(command):
    if command is None:
        return None  # No response when no command is detected
    
    if "alexa" not in command.lower():
        return None  # Ignore commands without the wake word
    
    logging.debug(f"Received Command: {command}")

    if command == "alexa stop":
        return "Stopping process... Goodbye!"
    
    match_open = re.search(r"alexa\s*open\s+(.+)", command)
    match_close = re.search(r"alexa\s*close\s+(.+)", command)
    
    if match_open:
        app_or_folder = match_open.group(1)
        recognized_apps = ["word", "excel", "powerpoint", "notepad", "calculator", "paint", "edge browser", 
                          "chrome browser", "firefox browser", "vlc", "spotify", "adobe reader", "command prompt", 
                          "task manager", "windows media player", "photos", "vs code", "visual studio code", "code"]
        if app_or_folder in recognized_apps:
            result = manage_app(app_or_folder, "open")
        else:
            result = open_folder(app_or_folder)
        return result
    
    if "close all windows" in command.lower() or "close all apps" in command.lower():
        if platform.system() == "Windows":
            try:
                import pygetwindow as gw
                all_windows = gw.getAllTitles()
                for win_title in all_windows:
                    if win_title:  # Skip empty titles
                        try:
                            window = gw.getWindowsWithTitle(win_title)[0]
                            window.close()
                        except Exception as e:
                            logging.error(f"Error closing window: {win_title}, {str(e)}")
                return "Closed all open windows."
            except ImportError:
                return "pygetwindow is not supported on this platform."
        else:
            return "Closing all windows is only supported on Windows."
    
    if "close all folders" in command:
        result = close_folder("all folders")
        return result
    
    if match_close:
        app_or_folder = match_close.group(1)
        recognized_apps = ["word", "excel", "powerpoint", "notepad", "calculator", "paint", "edge browser", 
                           "chrome browser", "firefox browser", "vlc", "spotify", "adobe reader", "command prompt", 
                           "task manager", "windows media player", "photos", "vs code", "visual studio code", "code"]
        if app_or_folder in recognized_apps:
            result = manage_app(app_or_folder, "close")
        else:
            result = close_folder(app_or_folder)
        return result
    
    if "shutdown" in command:
        result = shutdown_computer()
        return result
    
    if "restart" in command:
        result = restart_computer()
        return result
    
    if "lock" in command:
        result = lock_computer()
        return result
    
    if "volume" in command:
        result = adjust_volume(command)
        return result
    
    if "brightness" in command:
        brightness_match = re.search(r"(increase|decrease|set)\s*brightness\s*(?:to)?\s*(\d+)%?", command)
        if brightness_match:
            action = brightness_match.group(1)
            brightness_level = int(brightness_match.group(2))
            result = set_brightness(brightness_level)
            return result
    
    if "screenshot" in command:
        result = take_screenshot()
        return result
    
    return "Command not recognized."

if __name__ == '__main__':
    print("Voice Command Processor is running. Say 'Alexa stop' to exit.")
    while True:
        command = listen_for_command()
        if command:  # Only process if we got a command
            print(f"Processing command: {command}")
            response = handle_command(command)
            if response:  # Only speak if we have a response
                print(response)
                speak(response)
                if response == "Stopping process... Goodbye!":
                    break