import csv

from datetime import datetime
import time

import cv2
import face_recognition_models
import face_recognition
import numpy as np
from win32com.client import Dispatch
def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1+" is present")
def speak1(str2):
    speak1=Dispatch(("SAPI.SpVoice"))
    speak1.speak(str2)

video_capture = cv2.VideoCapture(0)

# load non faces
sumit_image = face_recognition.load_image_file("faces/sumit.jpg")
sumit_encoding = face_recognition.face_encodings(sumit_image)[0]


mohit_image = face_recognition.load_image_file("faces/mohit.jpg")
mohit_encoding = face_recognition.face_encodings(mohit_image)[0]

known_face_encodings = [sumit_encoding, mohit_encoding]
known_face_names = ["Sumit", "Mohit"]
# known_face_encodings=[sumit_encoding]
# known_face_names=["Sumit"]
# list of expected students
students = known_face_names.copy()
face_locations = []
face_encodings = []
# get the current date and time
now = datetime.now()
current_date = now.strftime("%Y-%m-%d")

f = open(f"{current_date}.csv", "w+", newline="")
lnwriter = csv.writer(f)

while True:
    _, frame = video_capture.read()
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
    rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

    # recognize faces
    face_locations = face_recognition.face_locations(rgb_small_frame)
    face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

    for face_encoding in face_encodings:
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        face_distance = face_recognition.face_distance(known_face_encodings, face_encoding)
        best_match_index = np.argmin(face_distance)

        if matches[best_match_index]:
            name = known_face_names[best_match_index]
        # add the text if a person is present
        if name in known_face_names:
            font=cv2.FONT_HERSHEY_SIMPLEX
            bottomLeftCornerOfText=(10,100)
            fontScale=1.5
            fontColor=(255,0,0)
            thickness=3
            lineType=2
            cv2.putText(frame,name+" is present ",bottomLeftCornerOfText,font,fontScale,fontColor,thickness,lineType)

            if name in students:
                students.remove(name)
                current_time=now.strftime("%H-%M:%S")
                lnwriter.writerow([name,current_time])

    cv2.imshow(" attendence ", frame)

    if cv2.waitKey(1) & 0xFF ==ord("a"):
        speak(name)
        speak1("attendence Taken")
        time.sleep(0)
    if cv2.waitKey(1) & 0xFF == ord("q"):
        break

video_capture.release()
cv2.destroyAllWindows()
f.close()