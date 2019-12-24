import pyautogui
screenWidth, screenHeight=pyautogui.size()


if __name__ == '__main__':
    pyautogui.FAILSAFE=True
    pyautogui.click(clicks=20, interval=0.25)
    