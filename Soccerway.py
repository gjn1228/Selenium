# coding=utf-8 #
# Author GJN #

import xlwings as xw
import os
from Soccerway_Func import Socc

SF = Socc()
sd = SF.soccerway_download
ee = SF.elist_to_epl
eu = SF.elist_to_unl
ep = SF.elist_to_pfl

# r = sd()
# ee(r)
ua = 'https://int.soccerway.com/international/europe/uefa-nations-league/20182019/league-a/r45718/'
ub = 'https://int.soccerway.com/international/europe/uefa-nations-league/20182019/league-b/r45719/'
uc = 'https://int.soccerway.com/international/europe/uefa-nations-league/20182019/league-c/r45720/'
ud = 'https://int.soccerway.com/international/europe/uefa-nations-league/20182019/league-d/r45721/'


def update():
    eu(sd())


def sd2(url):
    return SF.soccerway_download2(url)


def enl_wb():
    path = r'Y:\Departments\CSLRM\Level D\FB Odds Locals\National team\UEFA National League\\'
    for file in os.listdir(path):
        if file.startswith('UNL_') and file.endswith('xlsm'):
            xw.Book(path + file)


def up_enl():

    for u in [ua, ub, uc, ud]:
        eu(sd2(u))








