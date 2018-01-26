#coding:utf-8

import sys
import xlrd
import pygame
import random
from pygame.locals import *

fullScreen = False

#解析 人员.xlsx 文件，得到人员名单列表
def get_name_list_from_excel(file_name):
    name_list = []
    excelFile = xlrd.open_workbook(file_name)
    sheet = excelFile.sheet_by_name('Sheet1')
    print sheet.name, sheet.nrows, sheet.ncols
    job_num  = sheet.cell(0, 0).value.encode('utf-8')
    job_name = sheet.cell(0, 1).value.encode('utf-8')
    for row in range(1, sheet.nrows):
        job_num  = sheet.cell(row, 0).value.encode('utf-8')
        job_name = sheet.cell(row, 1).value.encode('utf-8')
        #print job_num, job_name
        name_list.append((job_num, job_name))

    return job_num, job_name, name_list


def handle_mouse_event(index, pause_flag):
    if pygame.mouse.get_pressed()[0]:
        x, y = pygame.mouse.get_pos()
        x -= m.get_width() / 2
        y -= m.get_height() / 2
        if 1031 + 91 > x > 1031 and 436 + 68 > y > 436:
            pause_flag = True if pause_flag == False else False
            if not pause_flag:
                pygame.mixer.music.play()
                del name_list[index]
                print len(name_list)
                show_name_list(name_list)
            else:
                pygame.mixer.music.stop()

    return pause_flag


def show_name_list(name_list):
    '''调试接口，输出当前名单'''
    for index in range(0, len(name_list)):
        str = "%s %s" % (name_list[index][0], name_list[index][1])
        print(str.decode('utf-8'))

if __name__ == "__main__":
    job_num, job_name, name_list = get_name_list_from_excel(r'name_file.xlsx')
    print len(name_list)

    pygame.init()
    if fullScreen:
        screen = pygame.display.set_mode((1024, 768), FULLSCREEN, 32)
    else:
        screen = pygame.display.set_mode((1340, 670), 0, 32)

    pygame.display.set_caption("文思海辉年会抽奖程序")
    bg = 'bg.png'
    mg = 'cz.png'

    #newimg = pygame.transform.scale(bg, (1024, 480))

    b = pygame.image.load(bg).convert()
    m = pygame.image.load(mg).convert_alpha()
    screen.blit(b, (0, 0))
    screen.blit(m, (0, 0))

    pygame.mixer.init()
    pygame.mixer.music.load('9224.wav')
    #pygame.mixer.music.set_volume(0.5)
    pygame.mixer.music.play()

    index = random.randint(0, len(name_list) - 1)
    pause_flag = False
    while True:
        for event in pygame.event.get():
            if event.type == pygame.KEYDOWN and event.key == pygame.K_ESCAPE:
                pygame.quit()
            elif event.type == MOUSEBUTTONDOWN:
                pause_flag = handle_mouse_event(index, pause_flag)

        screen.blit(b, (0, 0))
        x, y = pygame.mouse.get_pos()
        x -= m.get_width() / 2
        y -= m.get_height() / 2
        pygame.mouse.set_visible(False)
        screen.blit(m, (x, y))

        if not pause_flag:
            index = random.randint(0, len(name_list)-1)

        text_context = '%s %s' % (name_list[index][0], name_list[index][1])
        #print text_context
        font = pygame.font.SysFont("SimHei", 60)
        text_obj = font.render(text_context.decode('utf-8'), True, (255, 255, 255), (0, 0, 0))
        text_pos = text_obj.get_rect()
        text_pos.center = (650, 130)
        screen.blit(text_obj, text_pos)

        font = pygame.font.SysFont("SimHei", 30)
        text_obj = font.render('Ericsson ODC', True, (255, 255, 255), (255, 0, 0))
        text_pos = text_obj.get_rect()
        text_pos.center = (650, 540)
        screen.blit(text_obj, text_pos)

        pygame.display.update()

