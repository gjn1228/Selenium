# coding=utf-8 #
# Author GJN #
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import lxml.html
import pandas as pd
import numpy as np
from scipy.optimize import fsolve
import scipy.stats as st

# 初始化
chrome_options = Options()
driver = webdriver.Chrome(chrome_options=chrome_options)
url = 'https://www.sbobet.com/euro'
admin, pw = 'EURA0901223', 'csl00000'
driver.get(url)
driver.implicitly_wait(3)


def func(i, l1, l2, n1, n2):
    def matrix(sup, ttg, dadj=0.07, dsplit=0.5):
        home = {}
        away = {}
        for i in range(13):
            home[i] = st.poisson.pmf(i, (ttg + sup) / 2)
            away[i] = st.poisson.pmf(i, (ttg - sup) / 2)

        H, A, D = 0, 0, 0
        for i in home:
            for j in away:
                x = home[i] * away[j]
                if i > j:
                    H += x
                elif i == j:
                    D += x
                elif i < j:
                    A += x

        dx = 1 + dadj
        hx = (H - D * dadj * dsplit) / H
        ax = (A - D * dadj * (1 - dsplit)) / A

        score = {}
        for i in home:
            for j in away:
                x = home[i] * away[j]
                if i > j:
                    score[(i, j)] = x * hx
                elif i == j:
                    score[(i, j)] = x * dx
                elif i < j:
                    score[(i, j)] = x * ax
        return score

    score = matrix(i[0], i[1])

    def line(score, line):

        def over(x):
            o = 0
            for s in score:
                if s[0] - s[1] > -x:
                    o += score[s]
            return o

        def under(x):
            o = 0
            for s in score:
                if s[0] - s[1] < -x:
                    o += score[s]
            return o

        if line % 1 == 0.5:
            return [over(line), under(line)]
        if line % 1 == 0:
            x = over(line)
            y = under(line)
            z = x + y
            return [x / z, y / z]
        if line % 1 == 0.25:
            x = over(line - 0.25)
            y = under(line - 0.25)
            z = x + y
            u = y / z
            v = under(line + 0.25)
            w = 2 / (1 / u + 1 / v)
            return [1 - w, w]
        if line % 1 == 0.75:
            x = over(line + 0.25)
            y = under(line + 0.25)
            z = x + y
            u = x / z
            v = over(line - 0.25)
            w = 2 / (1 / u + 1 / v)
            return [w, 1 - w]

    def hline(score, line):
        def over(x):
            o = 0
            for s in score:
                if s[0] + s[1] > x:
                    o += score[s]
            return o

        def under(x):
            o = 0
            for s in score:
                if s[0] + s[1] < x:
                    o += score[s]
            return o

        if line % 1 == 0.5:
            return [over(line), under(line)]
        if line % 1 == 0:
            x = over(line)
            y = under(line)
            z = x + y
            return [x / z, y / z]
        if line % 1 == 0.25:
            x = over(line - 0.25)
            y = under(line - 0.25)
            z = x + y
            u = y / z
            v = under(line + 0.25)
            w = 2 / (1 / u + 1 / v)
            return [1 - w, w]
        if line % 1 == 0.75:
            x = over(line + 0.25)
            y = under(line + 0.25)
            z = x + y
            u = x / z
            v = over(line - 0.25)
            w = 2 / (1 / u + 1 / v)
            return [w, 1 - w]

    nh = np.array(line(score, l1))
    nhl = np.array(hline(score, l2))
    nxl = np.array([nh[0], nhl[0]])
    nx = np.array([n1, n2])

    return (nxl - nx).tolist()


def rarg(x):
    return x[0], x[2], 1 / x[1], 1 / x[3]


# 输入(ha-line, ha-home, ha-line, hilo-line, hilo-home, hilo-away),输出（sup, ttg)
def tost(x):
    hilo = x[4] * (1 / x[4] + 1 / x[5])
    try:
        hline = -float(x[0].split('/')[0])
    except AttributeError:
        hline = -x[0]
    handi = x[1] * (1 / x[1] + 1 / x[2])
    arg = (hline, handi, x[3], hilo)
    x0 = 0
    delta = 0
    na = np.array([x0, 2])
    f = fsolve(func, na, rarg(arg))
    # 180923部分大sup和ttg解不出来，优化了初始值
    while (f == na).all():
        delta += 1
        x1 = x0 + delta
        na = np.array([x1, 2])
        f = fsolve(func, np.array(na), rarg(arg))
        if (f == na).all():
            x1 = x0 - delta
            na = np.array([x1, 2])
            f = fsolve(func, np.array(na), rarg(arg))
    # if (f == na).all():
    #     x0 = 1
    #     na = np.array([x0, 3.5])
    #     f = fsolve(Rfunc.func, na, Rfunc.rarg(arg))
    #     while (f == na).all():
    #         x0 += 1
    #         na = np.array([x0, 3.5])
    #         f = fsolve(Rfunc.func, np.array([0, 3.5]), Rfunc.rarg(arg))
    return f


def get_df():
    tree = lxml.html.fromstring(driver.page_source)
    t = tree.xpath('//*[@id="odds-display-nonlive"]/table')[0]
    # name = t.xpath('tbody/tr/td[@class="BLN team-name-column"]')
    # n = [na.xpath('span') for na in name]
    # n1 = [n2.text_content()for n3 in n for n2 in n3]
    row = t.xpath('tbody/tr')
    col = [r.xpath('td') for r in row]
    text = [[c1.text_content() for c1 in c2] for c2 in col]
    head = [h.text_content() for h in t.xpath('tbody/tr/th')]
    head.insert(2, 'team')
    h2 = []
    i = 0
    for h in head:
        if i in [10, 11, 13, 14]:
            h2.append('Half-' + h)
        else:
            h2.append(h)
        i += 1
    h2[1] = 'League'
    tout = []
    time, league, team = '', '', ''
    for tx in text:
        if len(tx) == 1:
            league = tx[0]
        if len(tx) == 16:
            if tx[0].strip() != '':
                time = tx[0].strip()
                tx[0] = time
            else:
                tx[0] = time
            if tx[1].strip() == '':
                tx[1] = league
            if tx[2].strip() != '':
                team = tx[2].strip()
                tx[2] = team
            else:
                tx[2] = team
            tout.append(tx)

    df = pd.DataFrame(tout, columns=h2)
    return df


def get_st():
    tree = lxml.html.fromstring(driver.page_source)
    t = tree.xpath('//*[@id="odds-display-nonlive"]/table')[0]
    # name = t.xpath('tbody/tr/td[@class="BLN team-name-column"]')
    # n = [na.xpath('span') for na in name]
    # n1 = [n2.text_content()for n3 in n for n2 in n3]
    row = t.xpath('tbody/tr')
    col = [r.xpath('td') for r in row]
    favor = [r.xpath('td[@class="BLN team-name-column"]/span') for r in row]
    text = [[c1.text_content() for c1 in c2] for c2 in col]
    redblue = [[c1.get('class') for c1 in c2] for c2 in favor]
    team = [[c1.text_content() for c1 in c2] for c2 in favor]
    head = [h.text_content() for h in t.xpath('tbody/tr/th')]
    head.insert(2, 'team')
    h2 = []
    i = 0
    for h in head:
        if i in [10, 11, 13, 14]:
            h2.append('Half-' + h)
        else:
            h2.append(h)
        i += 1
    h2[1] = 'League'
    tout = []
    time, league, team = '', '', ''
    for tx in text:
        if len(tx) == 1:
            league = tx[0]
        if len(tx) == 16:
            if tx[0].strip() != '':
                time = tx[0].strip()
                tx[0] = time
            else:
                tx[0] = time
            if tx[1].strip() == '':
                tx[1] = league
            if tx[2].strip() != '':
                team = tx[2].strip()
                tx[2] = team
            else:
                tx[2] = team
            tout.append(tx)

    df = pd.DataFrame(tout, columns=h2)
    return df


ts = tost([0.5, 1.9, 2.05, 2.75, 2.12, 1.81])

