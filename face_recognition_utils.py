import cv2
import os
import openpyxl

def load_known_faces():
    known_faces = []
    known_names = []
    known_roll_nos = []
    users_path = "users"

    if os.path.exists('users.xlsx'):
        workbook = openpyxl.load_workbook('users.xlsx')
        sheet = workbook["Users"]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            roll_no, name = row
            img_path = os.path.join(users_path, f"{name}_{roll_no}.jpg")
            if os.path.exists(img_path):
                image = cv2.imread(img_path)
                known_faces.append(image)
                known_names.append(name)
                known_roll_nos.append(roll_no)

    return known_faces, known_names, known_roll_nos

def recognize_faces(frame, known_faces, known_names, known_roll_nos):
    names = []
    roll_nos = []
    for i, known_face in enumerate(known_faces):
        result = cv2.matchTemplate(frame, known_face, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        if max_val > 0.7:
            names.append(known_names[i])
            roll_nos.append(known_roll_nos[i])

    return names, roll_nos
