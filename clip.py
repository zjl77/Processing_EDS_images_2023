from PIL import ImageFont, ImageDraw, Image
import numpy as np
import os
import cv2

pic_path = r"D:\数据\19号样品\7.11 4# EDS\ZJL-EDS-7.11已调色\项目 1\reports\项目 1_区 11_2023-08-18_11-29-37_EDS\\"
height_start = 0.25
height_stop = 0.75
width_start = 0.25
width_stop = 0.75

def clipimage(image, result_path):
    img = Image.open(f"{pic_path}\\{image}")
    width, height = img.size
    img_new = img.crop((int(width * width_start), int(height * height_start), int(width * width_stop), int(height * height_stop)))
    img_new.save(f'{result_path}\\{image}', format='png', subsampling=0, quality=100)

def add_word(image):
    img = Image.open(image)
    width, height = img.size
    image_name = image.split("\\")[-1].split(".")[0]
    font = ImageFont.truetype(r"Arial.ttf", 45)
    image_draw = ImageDraw.Draw(img)
    font_width = image_draw.textlength(image_name, font=font)
    image_draw.text((width-font_width-12, 12), image_name, font=font, fill="white")
    img.save(image, format='png', subsampling=0, quality=100)


if __name__ == '__main__':
    result_path = f'{pic_path}\\clip\\'
    if not os.path.exists(result_path):
        os.makedirs(result_path)
    for image in os.listdir(pic_path):
        if ".png" in image:
            clipimage(image, result_path)
            add_word(f'{result_path}\\{image}')

