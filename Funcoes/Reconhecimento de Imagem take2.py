import os

import pytesseract
import cv2
import pyautogui as pg
from PIL import Image
import re
from time import sleep
import matplotlib.pyplot as plt
import numpy as np
from pytesseract.pytesseract import Output

def a(cnpj_da_empresa):
    path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    os.chdir(r'C:\Users\Carvalho-Admin\Downloads\CSSSSS')


    impa = pg.screenshot('Imagens/Tela Atual.png',region=(910,480,130,175))

    img = cv2.imread('Imagens/Tela Atual.png')
    im = cv2.cvtColor(img,cv2.COLOR_RGB2GRAY)
    img = cv2.resize(img,None,fx=0.25,fy=0.25,interpolation=cv2.INTER_CUBIC)
    kernel = np.ones((1,1), np.uint8)
    img = cv2.dilate(img, kernel, iterations=1)
    img = cv2.erode(img, kernel, iterations=1)
    #
    # img = cv2.threshold(cv2.medianBlur(img, 3), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

    #Convertendo pra cinza





    #im = cv2.threshold(im,0,255,cv2.THRESH_BINARY +cv2.THRESH_OTSU)[1]

    plt.imshow(img)
    plt.show()



    pytesseract.pytesseract.tesseract_cmd = path
    hImg, wImg = im.shape

    boxes = pytesseract.image_to_data(im)
    posicao = 1
    for x, b in enumerate(boxes.splitlines()):

        if x != 0:
            b = b.split()

            if len(b) == 12:

                x, y, w, h = int(b[6]), int(b[7]), int(b[8]), int(b[9])

                if 600 > x > 40 and 850 > y > 50:

                    cnpj = b[11]
                    cv2.rectangle(im, (x, y), (w + x, h + y), (0, 0, 255), 2)
                    print(cnpj)
                    # print(posicao, cnpj)
                    if cnpj_da_empresa in cnpj:
                        # print(posicao,cnpj)
                        plt.imshow(im)
                        plt.show()
                        return posicao
                    else:
                        posicao = posicao + 1
    plt.imshow(im)
    plt.show()
    print('CNPJ não está na imagem')
    return None
    cv2.imshow('img', im)
    cv2.waitKey(0)

# #
# posicao_do_cnpj = a('14.239.847/0001-38')
# print(posicao_do_cnpj)