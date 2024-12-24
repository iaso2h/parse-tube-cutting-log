from pynput import keyboard
from pynput.mouse import Button, Controller
from PIL import ImageGrab
import win32gui, win32process, win32api, win32con
import psutil
import time
import copy
import console

mouse = Controller()
mouseInterval = 0.075
print = console.print


def hotkeyAlignTube():
    pId = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[1]
    pName = psutil.Process(pId).name()
    if pName == "TubePro.exe":
        screenX = 123
        screenY = 124
        img = ImageGrab.grab()
        imgRGB = img.convert("RGB")
        startPixelPos    = [850, 1013]
        pausePixelPos    = [972, 1013]
        continuePixelPos = [850, 1077]
        stopPixelPos     = [972, 1077]
        targetCompletedPixel = imgRGB.getpixel((15, 1810))
        if targetCompletedPixel == (170, 170, 0) or targetCompletedPixel == (255, 155, 155):
            startPixelPos[2]    = startPixelPos[2] + 60
            pausePixelPos[2]    = pausePixelPos[2] + 60
            continuePixelPos[2] = continuePixelPos[2] + 60
            stopPixelPos[2]     = stopPixelPos[2] + 60

        startPixel    = imgRGB.getpixel(startPixelPos)
        pausePixel    = imgRGB.getpixel(pausePixelPos)
        continuePixel = imgRGB.getpixel(continuePixelPos)
        stopPixel     = imgRGB.getpixel(stopPixelPos)

        if startPixel == (0, 160, 45) and \
            pausePixel == (255, 167, 51) and \
            continuePixel == (196, 196, 196) and \
            stopPixel == (228, 28, 28):
                savedPosition = copy.copy(mouse.position)
                mouse.position = (screenX + 386, screenY + 123)
                mouse.press(Button.left)
                mouse.release(Button.left)
                time.sleep(mouseInterval)
                mouse.position = (screenX + 386, screenY + 199)
                mouse.press(Button.left)
                mouse.release(Button.left)
                time.sleep(mouseInterval)
                mouse.position = (screenX + 576, screenY + 223)
                mouse.press(Button.left)
                mouse.release(Button.left)
                mouse.position = savedPosition
        else:
            pass
            # return win32api.MessageBox(
            #         None,
            #         "激光机运行时无法进行此操作",
            #         "Warning",
            #         4096 + 0 + 16
            #     )
            # MB_SYSTEMMODAL==4096
            # Button Styles:
            # 0:OK  --  1:OK|Cancel -- 2:Abort|Retry|Ignore -- 3:Yes|No|Cancel -- 4:Yes|No -- 5:Retry|No -- 6:Cancel|Try Again|Continue
            # To also change icon, add these values to previous number
            # 16 Stop-sign  ### 32 Question-mark  ### 48 Exclamation-point  ### 64 Information-sign ('i' in a circle)



with keyboard.GlobalHotKeys({
        "<alt>+a": hotkeyAlignTube}) as h:
    import gui
    h.join()
