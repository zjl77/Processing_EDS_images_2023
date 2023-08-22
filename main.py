# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import cv2
from PIL import ImageFont, ImageDraw, Image
import numpy as np
import os
import pytesseract
import docx
import re
import cv2
import shutil

words_path = r"D:\数据\19号样品\6.28 2# EDS+EBSD\yanjing-ebsd-6.28-已调色\项目 1\reports\test\test\\"

pytesseract.pytesseract.tesseract_cmd = r'D:\software\Tesseract-OCR\tesseract.exe'


def get_pictures(word_path, result_path):
    file = docx.Document(word_path)
    dict_rel = file.part._rels
    i = 1
    for rel in dict_rel:
        rel = dict_rel[rel]
        if "image" in rel.target_ref:
            with open(f'{result_path}\\{i}.png', 'wb') as f:
                f.write(rel.target_part.blob)
                i += 1


def renameimage(image):
    img = cv2.imdecode(np.fromfile(image, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    # print(img[0, 0])
    if len(img[0, 0]) != 4:
        os.remove(image)
        return None
    height = img.shape[0]
    width = img.shape[1]
    for i in range(height):
        for j in range(width):
            if img[i, j][3] == 0:
                img[i, j] = [255, 255, 255, 255]
    # cv2.imshow('1', img)
    # cv2.waitKey(0)
    text = pytesseract.image_to_string(img, lang='eng', config='--psm 6')
    # print(text)
    new_name = re.match('^([A-Z]{1}[a-z]?)\s?(K|k|L|M)', text)
    if new_name:
        new_name_1 = new_name.group(1)
        if os.path.exists(f'{result_path}\\{new_name_1}.png'):
            img1 = cv2.imdecode(np.fromfile(f'{result_path}\\{new_name_1}.png', dtype=np.uint8), cv2.IMREAD_UNCHANGED)
            if img.shape[0] > img1.shape[0]:
                shutil.move(image, f'{result_path}\\{new_name_1}.png')
            else:
                os.remove(image)
        else:
            shutil.move(image, f'{result_path}\\{new_name_1}.png')
    else:
        os.remove(image)


def clipimage(image):
    img = cv2.imdecode(np.fromfile(image, dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    img = img[20:img.shape[0]-20, :]
    height = img.shape[0]
    width = img.shape[1]
    for i in range(height):
        for j in range(width):
            if img[i, j][3] == 0:
                img[i, j] = [255, 255, 255, 255]
    imgray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    ret, thresh = cv2.threshold(imgray, 127, 255, 0)
    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    for i in range(0, len(contours)):
    #     x, y, w, h = cv2.boundingRect(contours[i])
    #     cv2.rectangle(img, (x, y), (x+w, y+h), (153, 153, 0), 2)
    # cv2.imshow('2', img)
    # cv2.waitKey(0)
        peri = cv2.arcLength(contours[i], True)
        approx = cv2.approxPolyDP(contours[i], 0.04 * peri, True)
        if len(approx) == 4:
            x, y, w, h = cv2.boundingRect(contours[i])
            if w*h > 10000 and x != 0 and y != 0:
                img = img[y+2:y+h-2, x+2:x+w-2]
                break
    cv2.imencode('.png', img)[1].tofile(image)


def add_word(image):
    img = Image.open(image)
    width, height = img.size
    image_name = image.split("\\")[-1].split(".")[0]
    font = ImageFont.truetype(r"Arial.ttf", 60)
    image_draw = ImageDraw.Draw(img)
    font_width = image_draw.textlength(image_name, font=font)
    image_draw.text((width-font_width-24, 12), image_name, font=font, fill="white")
    img.save(image, format='png', subsampling=0, quality=100)


if __name__ == '__main__':
    for i in os.listdir(words_path):
        if '.docx' in i:
            word_path = words_path + i
            splitext = os.path.splitext(word_path)
            result_path = f'{splitext[0]}_EDS'
            if not os.path.exists(result_path):
                os.makedirs(result_path)
            get_pictures(word_path, result_path)
            for image in os.listdir(result_path):
                image_path = f'{result_path}\\{image}'
                renameimage(image_path)
            for image in os.listdir(result_path):
                image_path = f'{result_path}\\{image}'
                clipimage(image_path)
                add_word(image_path)
    # result_path = r"D:\数据\1号样品\8.1 1#场发射EDS —晶界夹杂成分\test\\"
    # for image in os.listdir(result_path):
    #     clipimage(image)
    #     add_word(f'{result_path}\\{image}')

    # for image in os.listdir(result_path):
    #     clipimage(image)
    #     add_word(f'{result_path}\\{image}')

