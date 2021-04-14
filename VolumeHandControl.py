import cv2
import time
import numpy as np
import HandTrackingModule as htm
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


wCam, hCam = 640, 480


cap = cv2.VideoCapture(0)
cap.set(2, wCam)
cap.set(4, hCam)

pTime = 0
cTime = 0

detector = htm.handDetector(detectionCon=0.7)

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]

volBar = np.interp(volume.GetMasterVolumeLevel(), [minVol, maxVol], [400, 150])

while True:
    success, img = cap.read()
    img = detector.findHands(img)
    lmList = detector.findPosition(img, draw=False)
    if len(lmList) != 0:
        #print(lmList[4], lmList[8])

        x1, y1 = lmList[4][1], lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]
        cx, cy = (x1+x2)//2, (y1+y2)//2

        cv2.circle(img, (x1, y1), 8, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), 8, (255, 0, 255), cv2.FILLED)
        cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
        cv2.circle(img, (cx, cy), 8, (255, 0, 255), cv2.FILLED)

        length = math.hypot(x2-x1, y2-y1)

        # Hand Range 40 - 200
        # Volume Range -65 - 0

        vol = np.interp(length, [40, 200], [minVol, maxVol])
        volBar = np.interp(length, [40, 200], [400, 150])
        volume.SetMasterVolumeLevel(vol, None)

        if length < 40:
            cv2.circle(img, (cx, cy), 8, (0, 255, 0), cv2.FILLED)

    img = cv2.flip(img, 1)

    cv2.rectangle(img, (50, 150), (85, 400), 	(255, 0, 0), 3)
    cv2.rectangle(img, (50, int(volBar)), (80, 400), (255, 0, 0), cv2.FILLED)

    curVol = np.interp(volume.GetMasterVolumeLevel(),
                       [minVol, maxVol], [0, 100])

    cv2.putText(img, f'{int(curVol)}%', (40, 450),
                cv2.FONT_HERSHEY_PLAIN, 3, 	(255, 0, 0), 3)

    cTime = time.time()
    fps = 1/(cTime-pTime)
    pTime = cTime

    cv2.putText(img, f'FPS: {int(fps)}', (10, 70),
                cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 255), 3)

    cv2.imshow("image", img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
cap.release()
cv2.destroyAllWindows()
