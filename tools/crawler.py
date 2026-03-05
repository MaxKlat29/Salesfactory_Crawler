import time
import webbrowser
import pyautogui
import pyperclip

pyautogui.PAUSE = 0.15 

def _select_all_and_copy(sleep_after_open: float = 5.0, sleep_after_kbd: float = 0.4) -> str:
    time.sleep(sleep_after_open)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(sleep_after_kbd)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(sleep_after_kbd)
    return pyperclip.paste()

def _close_current_tab(delay: float = 0.2):
    time.sleep(delay)
    pyautogui.hotkey('ctrl', 'w')

def crawl(profile_url: str, activity_url: str) -> str:
    profile_dump = ""
    activity_dump = ""

    opened_tabs = 0

    try:
        if profile_url:
            webbrowser.open(profile_url, new=2)
            opened_tabs += 1
            profile_dump = _select_all_and_copy()
        else:
            profile_dump = ""

        if activity_url:
            webbrowser.open(activity_url, new=2) 
            opened_tabs += 1
            activity_dump = _select_all_and_copy()
        else:
            activity_dump = ""

        result = f"Lebenslauf: {profile_dump} LinkedIn Beiträge: {activity_dump}"
        return result

    finally:
        if opened_tabs >= 1:
            _close_current_tab()
        if opened_tabs >= 2:
            _close_current_tab()
        time.sleep(0.15)