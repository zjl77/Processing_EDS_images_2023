from PIL import ImageFont, ImageDraw, Image
import numpy as np
import os
import pytesseract
import docx
import re
import cv2
import shutil

result_path = r"D:\数据\19号样品\6.28 2# EDS+EBSD\yanjing-ebsd-6.28-已调色\项目 1\reports\test\test\\"


def add_word(result_path):
    image = Image.open(result_path)
    width, height = image.size
    image_name = result_path.split("\\")[-1].split(".")[0]
    font = ImageFont.truetype(r"Arial.ttf", 60)
    image_draw = ImageDraw.Draw(image)
    font_width = image_draw.textlength(image_name, font=font)
    image_draw.text((width-font_width-24, 12), image_name, font=font, fill="white")
    image.save(f"{result_path}", format='png', subsampling=0, quality=100)

def clipimage(image):
    img = cv2.imdecode(np.fromfile(f'{result_path}\\{image}', dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    img = img[20:img.shape[0]-20, :img.shape[0]-5]
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
    cv2.imencode('.png', img)[1].tofile(f'{result_path}\\{image}')


if __name__ == "__main__":
    for image in os.listdir(result_path):
        clipimage(image)
        add_word(f'{result_path}\\{image}')
    # add_word(r'D:\数据\19号样品\8.1 7#8#EDS\zjl-eds(2)\项目 1\reports\test\项目 1_区 6_2023-08-01_18-47-43_EDS\新建文件夹\Co.png')
    # for i in os.listdir(words_path):
    #     if '.docx' in i:
    #         word_path = words_path + i
    #         splitext = os.path.splitext(word_path)
    #         result_path = f'{splitext[0]}_picture'
    #         if not os.path.exists(result_path):
    #             os.makedirs(result_path)
    #         get_pictures(word_path, result_path)
    # renameimage(r'D:\数据\19号样品\6.21 TEM\YJ\项目 1\reports\test\项目 1_区 5_2023-06-21_23-13-11_picture')
    # clipimage(r'D:\数据\19号样品\6.21 TEM\YJ\项目 1\reports\test\项目 1_区 5_2023-06-21_23-13-11_picture')
