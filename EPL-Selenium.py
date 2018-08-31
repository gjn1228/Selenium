# coding=utf-8 #
# Author GJN #
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import xlwings as xw
import lxml.html
from bs4 import BeautifulSoup
import json
import time
import pandas as pd
import datetime
import selenium
from Selenium_Func import Sele

SF = Sele()
times = time.clock()
chrome_options = Options()
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--disable-gpu')
driver = webdriver.Chrome(chrome_options=chrome_options)
# driver.get('https://www.whoscored.com/Matches/1190425/LiveStatistics/England-Premier-League-2017-2018-Leicester-Arsenal')
# driver.implicitly_wait(3)


wm = SF.whsc_match
rd = SF.r_to_df
dd = SF.df_to_draft2
de = SF.draft_to_epl
dp = SF.draft_to_pfl
wms = SF.whsc_match_summary
rds = SF.r_to_draft_summary
des = SF.draft_to_epl_summary
dps = SF.draft_to_pfl_summary
dec = SF.draft_to_epl_cup
dpc = SF.draft_to_pfl_cup


# # x = whsc_match(driver.page_source)
# t1 = 'https://www.whoscored.com/Matches/1190538/LiveStatistics/England-Premier-League-2017-2018-Leicester-Manchester-United'
# t2 = 'https://www.whoscored.com/Matches/1190197/LiveStatistics/England-Premier-League-2017-2018-Manchester-United-Leicester'
# t3 = 'https://www.whoscored.com/Matches/1080682/LiveStatistics/England-Premier-League-2016-2017-Leicester-Manchester-United'
# t4 = 'https://www.whoscored.com/Matches/1302108/LiveStatistics'
# t5 = 'https://www.whoscored.com/Matches/1284746/LiveStatistics'
# t6 = 'https://www.whoscored.com/Matches/1284750/LiveStatistics'
# t7 = 'https://www.whoscored.com/Matches/1201922/LiveStatistics'
# result = []
# for t in [t1, t2, t3, t4, t5, t6, t7]:
#     driver.get(t)
#     driver.implicitly_wait(3)
#     html = driver.page_source
#     result.append(whsc_match(html))
#
#
# timee = time.clock()
# print('+++++++++++++++++++++++++++++++++++++++++++++++++++++++')
# print('Times:%s' % (timee - times))
# print('+++++++++++++++++++++++++++++++++++++++++++++++++++++++')
#
# x = result[4]
# while True:
#     url = input('url:')
#     # url = 'https://www.whoscored.com/Matches/1301306/Live/Portugal-Liga-NOS-2018-2019-Portimonense-Boavista'
#     driver.get(url)
#     driver.implicitly_wait(3)
#     # elem = driver.find_element_by_xpath('/html/body/div[7]/div[1]/div[2]/div/div[2]/button[2]/span/strong')
#     # elem.click()
#     # driver.implicitly_wait(3)
#     # elem = driver.find_element_by_xpath('//*[@id="match-centre-stats"]/ul/li[1]/div[2]/span')
#     # elem.click()
#     # driver.implicitly_wait(3)
#     # time.sleep(1)
#     # elem = driver.find_element_by_xpath('//*[@id="sub-sub-navigation"]/ul/li[2]/a')
#     # elem.click()
#     # driver.implicitly_wait(3)
#     html = driver.page_source
#     # draft_to_epl(df_to_draft(r_to_df(whsc_match(html))))
#     draft_to_pfl(df_to_draft(r_to_df(whsc_match(html))))
#     print('Finished')
def me():
    dlist = driver.window_handles[1:]
    for d in dlist:
        driver.switch_to.window(d)
        # timeline
        tree = lxml.html.fromstring(driver.page_source)
        tr = tree.xpath('//span[@class="minute rc box"]')
        tt = [t.text_content() for t in tr]
        tl = [t.replace('\'', '') for t in tt]

        elem = driver.find_element_by_xpath('//*[@id="sub-sub-navigation"]/ul/li[2]/a')
        elem.click()
        driver.implicitly_wait(3)
        html = driver.page_source
        w = wm(html)
        r = rd(w)
        dfr = r[2]
        tlf = dfr[0].tolist()
        for ti in range(len(tlf)):
            if ti > 0:
                if tlf[ti] == tlf[ti - 1]:

                    if ti >= len(tl):
                        tl.insert(ti, tl[ti - 1])
                    else:
                        tll = tl[ti][:2]
                        tll2 = tl[ti - 1][:2]
                        if tll == '91':
                            tll = '90'
                        if tll2 == '91':
                            tll2 = '90'
                        if tll != tll2:
                            tl.insert(ti, tl[ti - 1])
                    print(tl[ti])
        if len(tlf) < len(tl):
            for ti in range(len(tlf)):
                if tlf[ti] > int(tl[ti].replace(',', '')[:2]):
                    tl.pop(ti)
        if len(tlf) < len(tl):
            for ti in range(len(tlf)):
                if tlf[ti] > int(tl[ti].replace(',', '')[:2]):
                    tl.pop(ti)
        if len(tl) < len(tlf):
            for ti in range(len(tlf)):
                tll = int(tl[ti].replace(',', '')[:2])
                if tlf[ti] < tll:
                    tl.insert(ti, str(tlf[ti]))
                    print(tlf[ti])
        dataf = dd(r, tl)
        de(dataf)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])


def mp():
    dlist = driver.window_handles[1:]
    for d in dlist:
        driver.switch_to.window(d)
        # timeline
        tree = lxml.html.fromstring(driver.page_source)
        tr = tree.xpath('//span[@class="minute rc box"]')
        tt = [t.text_content() for t in tr]
        tl = [t.replace('\'', '') for t in tt]

        elem = driver.find_element_by_xpath('//*[@id="sub-sub-navigation"]/ul/li[2]/a')
        elem.click()
        driver.implicitly_wait(3)
        html = driver.page_source
        w = wm(html)
        r = rd(w)
        dfr = r[2]
        tlf = dfr[0].tolist()
        for ti in range(len(tlf)):
            if ti > 0:
                if tlf[ti] == tlf[ti - 1]:

                    if ti >= len(tl):
                        tl.insert(ti, tl[ti - 1])
                    else:
                        tll = tl[ti][:2]
                        tll2 = tl[ti - 1][:2]
                        if tll == '91':
                            tll = '90'
                        if tll2 == '91':
                            tll2 = '90'
                        if tll != tll2:
                            tl.insert(ti, tl[ti - 1])
                    print(tl[ti])
        if len(tlf) < len(tl):
            for ti in range(len(tlf)):
                if tlf[ti] > int(tl[ti].replace(',', '')[:2]):
                    tl.pop(ti)
        if len(tlf) < len(tl):
            for ti in range(len(tlf)):
                if tlf[ti] > int(tl[ti].replace(',', '')[:2]):
                    tl.pop(ti)
        if len(tl) < len(tlf):
            for ti in range(len(tlf)):
                tll = int(tl[ti].replace(',', '')[:2])
                if tlf[ti] < tll:
                    tl.insert(ti, str(tlf[ti]))
                    print(tlf[ti])
        dataf = dd(r, tl)
        dp(dataf)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])


def cup():
    driver.switch_to.window(driver.window_handles[1])
    html = driver.page_source
    w = wms(html)
    r = rds(w)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return r


def cupe():
    dlist = driver.window_handles[1:]
    for d in dlist:
        driver.switch_to.window(d)
        html = driver.page_source
        w = wms(html)
        r = rds(w)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        des(r)


def cup_a():
    driver.switch_to.window(driver.window_handles[1])
    tree = lxml.html.fromstring(driver.page_source)
    tr = tree.xpath('//span[@class="minute rc box"]')
    tt = [t.text_content() for t in tr]
    tl = [t.replace('\'', '') for t in tt]

    elem = driver.find_element_by_xpath('//*[@id="sub-sub-navigation"]/ul/li[2]/a')
    elem.click()
    driver.implicitly_wait(3)
    html = driver.page_source
    w = wm(html)
    r = rd(w)
    dfr = r[2]
    tlf = dfr[0].tolist()
    for ti in range(len(tlf)):
        if ti > 0:
            if tlf[ti] == tlf[ti - 1]:
                if ti >= len(tl):
                    tl.insert(ti, tl[ti - 1])
                elif tl[ti][:2] != tl[ti - 1][:2]:
                    tl.insert(ti, tl[ti - 1])
                print(tl[ti])
    if len(tlf) < len(tl):
        for ti in range(len(tlf)):
            if tlf[ti] > int(tl[ti].replace(',', '')[:2]):
                tl.pop(ti)
    if len(tl) < len(tlf):
        for ti in range(len(tlf)):
            if tlf[ti] < int(tl[ti].replace(',', '')[:2]):
                print(tlf[ti])
                tl.insert(ti, str(tlf[ti]))

    dataf = dd(r, tl)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return dataf


# url = 'https://www.whoscored.com/Matches/1279505/Live/Portugal-Super-Cup-2017-2018-FC-Porto-Aves'
# url = 'https://www.whoscored.com/Matches/1325742/Live/Europe-UEFA-Champions-League-2018-2019-Vidi-FC-AEK-Athens'
# url = 'https://www.whoscored.com/Matches/1326684/Live/Europe-UEFA-Europa-League-2018-2019-Rapid-Wien-FC-FCSB'
# url = 'https://www.whoscored.com/Matches/1279481/Live/England-Community-Shield-2017-2018-Chelsea-Manchester-City'
# url = 'https://www.whoscored.com/Regions/252/Tournaments/2/Seasons/7361/England-Premier-League'
url = 'https://www.whoscored.com/Regions/252/Tournaments/2/Seasons/7361/England-Premier-League'
driver.get(url)
driver.implicitly_wait(3)
# html = driver.page_source
# w = wms(html)

# url = 'https://www.whoscored.com/Regions/252/Tournaments/2/Seasons/7361/England-Premier-League'
# driver.get(url)
# driver.implicitly_wait(3)



