
import time
import random

import numpy
import win32api
import win32com
import win32con
import win32gui


#TODO: There are a lot of sleeps around mouse interactions that might not be necessary.
# I was running into issues where the interface seemed to be lagging and the mouse was
# moving/clicking too fast.
# TODO: double click is needed.

# Yanked from https://gist.github.com/chriskiehl/2906125
VK_CODE = {'backspace':0x08,
           'tab':0x09,
           'clear':0x0C,
           'enter':0x0D,
           'shift':0x10,
           'ctrl':0x11,
           'alt':0x12,
           'pause':0x13,
           'caps_lock':0x14,
           'esc':0x1B,
           'spacebar':0x20,
           'page_up':0x21,
           'page_down':0x22,
           'end':0x23,
           'home':0x24,
           'left_arrow':0x25,
           'up_arrow':0x26,
           'right_arrow':0x27,
           'down_arrow':0x28,
           'select':0x29,
           'print':0x2A,
           'execute':0x2B,
           'print_screen':0x2C,
           'ins':0x2D,
           'del':0x2E,
           'help':0x2F,
           '0':0x30,
           '1':0x31,
           '2':0x32,
           '3':0x33,
           '4':0x34,
           '5':0x35,
           '6':0x36,
           '7':0x37,
           '8':0x38,
           '9':0x39,
           'a':0x41,
           'b':0x42,
           'c':0x43,
           'd':0x44,
           'e':0x45,
           'f':0x46,
           'g':0x47,
           'h':0x48,
           'i':0x49,
           'j':0x4A,
           'k':0x4B,
           'l':0x4C,
           'm':0x4D,
           'n':0x4E,
           'o':0x4F,
           'p':0x50,
           'q':0x51,
           'r':0x52,
           's':0x53,
           't':0x54,
           'u':0x55,
           'v':0x56,
           'w':0x57,
           'x':0x58,
           'y':0x59,
           'z':0x5A,
           'numpad_0':0x60,
           'numpad_1':0x61,
           'numpad_2':0x62,
           'numpad_3':0x63,
           'numpad_4':0x64,
           'numpad_5':0x65,
           'numpad_6':0x66,
           'numpad_7':0x67,
           'numpad_8':0x68,
           'numpad_9':0x69,
           'multiply_key':0x6A,
           'add_key':0x6B,
           'separator_key':0x6C,
           'subtract_key':0x6D,
           'decimal_key':0x6E,
           'divide_key':0x6F,
           'F1':0x70,
           'F2':0x71,
           'F3':0x72,
           'F4':0x73,
           'F5':0x74,
           'F6':0x75,
           'F7':0x76,
           'F8':0x77,
           'F9':0x78,
           'F10':0x79,
           'F11':0x7A,
           'F12':0x7B,
           'F13':0x7C,
           'F14':0x7D,
           'F15':0x7E,
           'F16':0x7F,
           'F17':0x80,
           'F18':0x81,
           'F19':0x82,
           'F20':0x83,
           'F21':0x84,
           'F22':0x85,
           'F23':0x86,
           'F24':0x87,
           'num_lock':0x90,
           'scroll_lock':0x91,
           'left_shift':0xA0,
           'right_shift ':0xA1,
           'left_control':0xA2,
           'right_control':0xA3,
           'left_menu':0xA4,
           'right_menu':0xA5,
           'browser_back':0xA6,
           'browser_forward':0xA7,
           'browser_refresh':0xA8,
           'browser_stop':0xA9,
           'browser_search':0xAA,
           'browser_favorites':0xAB,
           'browser_start_and_home':0xAC,
           'volume_mute':0xAD,
           'volume_Down':0xAE,
           'volume_up':0xAF,
           'next_track':0xB0,
           'previous_track':0xB1,
           'stop_media':0xB2,
           'play/pause_media':0xB3,
           'start_mail':0xB4,
           'select_media':0xB5,
           'start_application_1':0xB6,
           'start_application_2':0xB7,
           'attn_key':0xF6,
           'crsel_key':0xF7,
           'exsel_key':0xF8,
           'play_key':0xFA,
           'zoom_key':0xFB,
           'clear_key':0xFE,
           '+':0xBB,
           ',':0xBC,
           '-':0xBD,
           '.':0xBE,
           '/':0xBF,
           '`':0xC0,
           ';':0xBA,
           '[':0xDB,
           '\\':0xDC,
           ']':0xDD,
           "'":0xDE,
           '`':0xC0}



class ScreenInteraction(object):
    def __init__(self, gamewindow):
        self.window = gamewindow

    def set_window_active(self):
        # For testing purposes, need to make the target window the active window.
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(self.window.get('targetwindow'))
        time.sleep(0.05)

    def set_focus(self, coords=(0,0)):
        # Click on the screen to set focus. Can't find a reliable way to do this.
        self.click(coords, "left")
        # This might be unnecessary.
        time.sleep(0.10)

    def click_duration(self, duration):
        mu = duration
        sigma = duration*0.20
        click_duration = numpy.random.normal(mu, sigma, 1)[0]
        return click_duration

    def click_create_coords(self, coords):
        # The coordinates provided will need to be adjusted to be in reference to the actual
        # desktop coordinates of the window we're interacting with.
        # In short, we'll be treating 0,0 as the top left corner of the game window, and 
        # ScreenInteraction will be responsible for adjusting the offsets.
        x_coord = coords[0] + self.window.get('left')
        y_coord = coords[1] + self.window.get('top')
        return (x_coord, y_coord)

    def click_event(self, click_down, click_up, duration, coords):
            win32api.mouse_event(click_down, coords[0], coords[1], 0, 0)
            time.sleep(duration)
            win32api.mouse_event(click_up, coords[0], coords[1], 0, 0)

    def click(self, coords, click_type="left", duration=0.200, double=False):
        # Perform a left click at coords with a duration between click and release.
        # Default duration of 75ms seemed reasonable.
        click_duration = self.click_duration(duration)
        coords = self.click_create_coords(coords)

        print(f"ScreenInteraction placing mouse at: {coords[0]}, {coords[1]} with duration: {click_duration}")
        if click_type == "left":
            click_down = win32con.MOUSEEVENTF_LEFTDOWN
            click_up = win32con.MOUSEEVENTF_LEFTUP
        elif click_type == "right":
            click_down = win32con.MOUSEEVENTF_RIGHTDOWN
            click_up = win32con.MOUSEEVENTF_RIGHTUP
        else:
            raise Exception(f"Unknown click type {click_type}")
        self.set_window_active()
        win32api.SetCursorPos(coords)
        # The UI seems to be VERY slow to respond to the mouse. When positioning the mouse, we have to wait a short period of time for the UI to register it.
        time.sleep(click_duration/2)
        if double:
            self.click_event(click_down, click_up, click_duration/4, coords)
            time.sleep(click_duration/4)
            self.click_event(click_down, click_up, click_duration/4, coords)            
        else:
            self.click_event(click_down, click_up, click_duration, coords)

    def calculate_point_from_bbox(self, bbox, offset_bbox=None):
        # tesseract-OCR provides bbox coords for words/lines. With that data we then
        # need to determine a singular point within that bbox to click.
        if offset_bbox:
            # Sometimes functions want to click on something on the screen, but they're operating with a limited snapshot of a portion of the screen. Convert the coordinates for their limited snapshot to the whole window.
            bbox["left"] = bbox.get("left") + offset_bbox.get("left")
            bbox["right"] = bbox.get("right") + offset_bbox.get("left")
            bbox["top"] = bbox.get("top") + offset_bbox.get("top")
            bbox["bottom"] = bbox.get("bottom") + offset_bbox.get("top")
        x_coord = random.randint(bbox.get('left'), bbox.get('right'))
        y_coord = random.randint(bbox.get('top'), bbox.get('bottom'))
        return (x_coord, y_coord)

    def downscale_bbox(self, bbox, resize_factor):
        # Sometimes we upscale images to help tesseract-OCR. Downscale the resultant bbox coords.
        bbox['left'] = int(bbox.get('left')/resize_factor)
        bbox['right'] = int(bbox.get('right')/resize_factor)
        bbox['top'] = int(bbox.get('top')/resize_factor)
        bbox['bottom'] = int(bbox.get('bottom')/resize_factor)
        return bbox

    def press(self, *args):
        '''
        one press, one release.
        accepts as many arguments as you want. e.g. press('left_arrow', 'a','b').
        '''
        for i in args:
            win32api.keybd_event(VK_CODE[i], 0,0,0)
            time.sleep(.05)
            win32api.keybd_event(VK_CODE[i],0 ,win32con.KEYEVENTF_KEYUP ,0)


    def pressAndHold(self, *args):
        '''
        press and hold. Do NOT release.
        accepts as many arguments as you want.
        e.g. pressAndHold('left_arrow', 'a','b').
        '''
        for i in args:
            win32api.keybd_event(VK_CODE[i], 0, 0, 0)
            time.sleep(.05)


    def pressHoldRelease(self, *args):
        '''
        press and hold passed in strings. Once held, release
        accepts as many arguments as you want.
        e.g. pressAndHold('left_arrow', 'a','b').

        this is useful for issuing shortcut command or shift commands.
        e.g. pressHoldRelease('ctrl', 'alt', 'del'), pressHoldRelease('shift','a')
        '''
        for i in args:
            win32api.keybd_event(VK_CODE[i], 0, 0, 0)
            time.sleep(.05)

        for i in args:
                win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(.1)


    def release(self, *args):
        '''
        release depressed keys
        accepts as many arguments as you want.
        e.g. release('left_arrow', 'a','b').
        '''
        for i in args:
               win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)
