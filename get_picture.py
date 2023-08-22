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

words_path = r"D:\数据\19号样品\6.28 2# EDS+EBSD\yanjing-ebsd-6.28-已调色\项目 1\reports\test\\"


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


if __name__ == '__main__':
    for i in os.listdir(words_path):
        if '.docx' in i:
            word_path = words_path + i
            splitext = os.path.splitext(word_path)
            result_path = f'{splitext[0]}_picture'
            if not os.path.exists(result_path):
                os.makedirs(result_path)
            get_pictures(word_path, result_path)

