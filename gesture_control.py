import cv2
import mediapipe as mp
import math
import numpy as np
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pyautogui
import time

# Initialize MediaPipe hands class.
mpHands = mp.solutions.hands
hands = mpHands.Hands()
mpDraw = mp.solutions.drawing_utils

# Initialize pycaw
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# Initialize pyautogui
pyautogui.PAUSE = 0.5

# Set up the webcam
cap = cv2.VideoCapture(0)

# Variables to track gestures
last_gesture_time = 0
thumb_down = False
scrolling = False
scroll_direction = 0

def percentage_to_db(percentage):
    return -96.0 + (percentage * (0 - (-96.0)) / 100)

# Default volume range if method is not available
try:
    min_vol = volume.GetMinLevel()
    max_vol = volume.GetMaxLevel()
except AttributeError:
    print("GetMinLevel or GetMaxLevel not available. Using default values.")
    min_vol = -96.0
    max_vol = 0.0

while True:
    success, image = cap.read()
    image = cv2.flip(image, 1)
    results = hands.process(image)
    if results.multi_hand_landmarks:
        for handLms in results.multi_hand_landmarks:
            thumb_x, thumb_y = None, None
            index_x, index_y = None, None
            for id, lm in enumerate(handLms.landmark):
                h, w, c = image.shape
                cx, cy = int(lm.x * w), int(lm.y * h)
                if id == 4:  # Thumb tip
                    thumb_x, thumb_y = cx, cy
                if id == 8:  # Index finger tip
                    index_x, index_y = cx, cy
            
            if thumb_x is not None and index_x is not None:
                length = math.hypot(index_x - thumb_x, index_y - thumb_y)
                vol = np.interp(length, [50, 220], [0, 100])
                vol_db = percentage_to_db(vol)
                vol_db = np.clip(vol_db, min_vol, max_vol)
                
                # Debugging output
                print(f"Volume level in dB: {vol_db}")
                
                try:
                    volume.SetMasterVolumeLevel(vol_db, None)
                except Exception as e:
                    print(f"Error setting volume: {e}")
                
                volPer = np.interp(length, [50, 220], [0, 100])
                volBar = 400
                cv2.rectangle(image, (50, 150), (85, 400), (0, 0, 0), 3)
                cv2.rectangle(image, (50, int(volBar - (volBar * volPer / 100))), (85, 400), (0, 0, 0), cv2.FILLED)
                cv2.putText(image, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 0, 0), 3)

                # Check for thumb down gesture
                if thumb_y > 300 and not thumb_down:
                    thumb_down = True
                    last_gesture_time = time.time()
                    pyautogui.press('playpause')
                elif thumb_y < 300 and thumb_down:
                    thumb_down = False

                # Check for next track gesture
                if length > 220 and abs(index_y - thumb_y) < 50 and time.time() - last_gesture_time > 1:
                    last_gesture_time = time.time()
                    pyautogui.press('nexttrack')

                # Check for previous track gesture
                if length < 50 and abs(index_y - thumb_y) < 50 and time.time() - last_gesture_time > 1:
                    last_gesture_time = time.time()
                    pyautogui.press('prevtrack')

                # Check for volume up gesture
                if index_y < thumb_y - 50 and time.time() - last_gesture_time > 1:
                    last_gesture_time = time.time()
                    current_vol = volume.GetMasterVolumeLevel()
                    volume.SetMasterVolumeLevel(min(current_vol + 1.0, max_vol), None)

                # Check for volume down gesture
                if index_y > thumb_y + 50 and time.time() - last_gesture_time > 1:
                    last_gesture_time = time.time()
                    current_vol = volume.GetMasterVolumeLevel()
                    volume.SetMasterVolumeLevel(max(current_vol - 1.0, min_vol), None)

                # Check for scrolling gesture
                if index_x > thumb_x + 50 and abs(index_y - thumb_y) < 50 and not scrolling:
                    scrolling = True
                    scroll_direction = 1
                elif index_x < thumb_x - 50 and abs(index_y - thumb_y) < 50 and not scrolling:
                    scrolling = True
                    scroll_direction = -1
                elif abs(index_x - thumb_x) > 50 or abs(index_y - thumb_y) > 50 and scrolling:
                    scrolling = False
                    if scroll_direction == 1:
                        pyautogui.scroll(100)
                    elif scroll_direction == -1:
                        pyautogui.scroll(-100)

    cv2.imshow('handDetector', image)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
