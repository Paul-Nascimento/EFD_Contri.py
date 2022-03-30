import pytesseract
import cv2
import pyautogui as pg
from PIL import Image
import re
from time import sleep
import matplotlib.pyplot as plt
import numpy as np
from pytesseract.pytesseract import Output


path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def buscando_posicao_cnpj(cnpj_da_empresa):
    sleep(4)

    impa = pg.screenshot('Imagens/Tela Atual.png',region=(910,480,130,175))

    im = cv2.imread('Imagens/Tela Atual.png')

    im = cv2.resize(im,None,fx=5,fy=5,interpolation=cv2.INTER_CUBIC)
    im = cv2.cvtColor(im,cv2.COLOR_BGR2GRAY)
    #im = cv2.threshold(im, , 0, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

    kernel = np.ones((1,1),np.uint8)

    sleep(1)

    pytesseract.pytesseract.tesseract_cmd = path
    hImg, wImg= im.shape

    boxes=pytesseract.image_to_data(im)
    posicao = 1
    for x,b in enumerate(boxes.splitlines()):

        if x != 0:
            b = b.split()

            if len(b) == 12:

                x,y,w,h = int(b[6]),int(b[7]),int(b[8]),int(b[9])

                if 600 > x > 35 and 850 > y > 50:


                    cnpj = b[11]
                    cv2.rectangle(im, (x, y), (w+x,h+y), (0, 0, 255), 2)
                    print(cnpj)
                    print(posicao, cnpj)
                    if cnpj_da_empresa in cnpj:
                        #print(posicao,cnpj)
                        plt.imshow(im)
                        plt.show()
                        return posicao
                    else:
                        posicao = posicao + 1
    plt.imshow(im)
    plt.show()
    print('CNPJ não está na imagem')


    return None
    #cv2.imshow('img', im)
    #cv2.waitKey(0)

#
#
# sleep(2)
#

# print(posicao_do_cnpj)
# buscando_posicao_cnpj('29.207.063/0001-03')
# print(posicao_do_cnpj)
# if posicao_do_cnpj == None:
#     print('Cnpj não encontrado')
#
# else:
#     pg.leftClick(753,(492+(15*(posicao_do_cnpj-1))))