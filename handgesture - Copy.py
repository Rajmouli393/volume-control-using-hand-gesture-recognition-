from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL, GUID
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import cv2
import numpy as np
import mediapipe as mp
import math

# Manually define the interface identifier (iid)
IID_IAudioEndpointVolume = GUID("{5CDF2C82-841E-4546-9722-0CF74078229A}")

devices = AudioUtilities.GetSpeakers()

# Use the manually defined iid
interface = devices.Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, None)

volume = cast(interface, POINTER(IAudioEndpointVolume))

volumeinformation = volume.GetVolumeRange()

minVolume = volumeinformation[0]
maxVolume = volumeinformation[1]

cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)

mpHands = mp.solutions.hands
hands = mpHands.Hands()
mpDraw = mp.solutions.drawing_utils

while True:
    status, img = cap.read()
    img = cv2.flip(img, 1)
    imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    results = hands.process(imgRGB)
    multiLandMarks = results.multi_hand_landmarks

    if multiLandMarks:
        indexPoint = ()
        thumbPoint = ()
        for handLms in multiLandMarks:
            for idx, lm in enumerate(handLms.landmark):
                h, w, c = img.shape
                cx, cy = int(lm.x * w), int(lm.y * h)
                if idx == 4:
                    thumbPoint = (cx, cy)
                if idx == 8:
                    indexPoint = (cx, cy)

        # Calculate distance without taking the square root if possible
        length_squared = ((indexPoint[0] - thumbPoint[0]) ** 2) + ((indexPoint[1] - thumbPoint[1]) ** 2)
        length = math.sqrt(length_squared) if length_squared >= 0 else 0

        vol = np.interp(length, [20, 200], [minVolume, maxVolume])
        print(vol)
        volume.SetMasterVolumeLevel(vol, None)

    cv2.imshow("HandGesture", img)
    cv2.waitKey(1)
