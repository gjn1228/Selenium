# coding=utf-8 #
# Author GJN #
import requests
import lxml.html
import xlwings as xw
from xlwings.constants import InsertShiftDirection
import datetime
import configparser
from Selenium_Func import Socc

SF = Socc()
sd = SF.soccerway_download
ee = SF.elist_to_epl

r = sd()
ee(r)












