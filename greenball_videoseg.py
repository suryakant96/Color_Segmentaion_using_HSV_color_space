import cv2
import numpy as np
from datetime import datetime
import time

import xlwt
from xlwt import Workbook

wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet v75')

dt1=datetime.now()



cam=cv2.VideoCapture("D:/Ball_project/MVI_1475.MP4")
i=0
width =640
height=480
window_capture_name="Original"
window_detection_name="ball"

Graph_window="Graph"
cv2.namedWindow(window_capture_name,cv2.WINDOW_NORMAL)
cv2.resizeWindow(window_capture_name, width, height)
cv2.namedWindow(window_detection_name,cv2.WINDOW_NORMAL)
cv2.resizeWindow(window_detection_name, width, height)
cv2.namedWindow(Graph_window,cv2.WINDOW_NORMAL)
cv2.resizeWindow(Graph_window, width, height)

ret, img0= cam.read()
graph = np.zeros(img0.shape, dtype="uint8")
ball=[]
j=0
i=0
while True:
    ret, img= cam.read()
    i+=1
    dt2 = datetime.now()


    txt = str(dt2.hour - dt2.hour) + ":" + str(dt2.minute - dt1.minute) + ":" + str(dt2.second - dt1.second) + "." + str(dt2.microsecond - dt1.microsecond)


    im=cv2.cvtColor(img,cv2.COLOR_BGR2HSV)
    #green
    #mask = cv2.inRange(im, (21,60,100), (180,140,255))


    #red
    mask = cv2.inRange(im, (20, 60, 50), (40, 160, 200))

    #mask = cv2.cvtColor(mask,cv2.COLOR_HSV2BGR)
    IMG=cv2.bitwise_and(img,img,mask=mask)

    gray_image=cv2.cvtColor(IMG,cv2.COLOR_BGR2GRAY)
    inv=cv2.bitwise_not(gray_image)
    ret, thresh = cv2.threshold(gray_image, 10, 255, 0)
    #print(ret)

    # time.sleep(0.1)
    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cX=[]
    cY=[]
    print("Contours : ",len(contours))
    for c in contours:
        M = cv2.moments(c)
        if M["m00"] >= 2000:
            cx = int(M["m10"] / M["m00"])
            cy = int(M["m01"] / M["m00"])
            # graph=graph+IMG
            cv2.circle(IMG, (cx, cy), 5, (255, 255, 255), -1)
            cv2.putText(IMG, "centroid: "+str(cx)+", "+str(cy), (cx - 25, cy - 25), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
            sheet1.write(j, 0, cx)
            sheet1.write(j, 1, cy)
            sheet1.write(j, 2, i)
            wb.save("D:/Ball_project/sample75.xls")
            j = j + 1
            graph = graph + IMG
            break
        else:
            cx, cy = 0, 0


    # cv2.imshow(window_capture_name,img)

    cv2.imshow(window_detection_name,IMG)
    cv2.imshow(Graph_window,graph)

    # cv2.imwrite(filename1, IMG)
    # cv2.imwrite(filename2, img)
    cv2.imwrite("segment_green75.png", graph)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cv2.imwrite("segment_green75.png",graph)
cam.release()
cv2.destroyAllWindows()
