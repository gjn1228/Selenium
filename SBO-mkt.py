# coding=utf-8 #
# Author GJN #
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import lxml.html
import pandas as pd
import numpy as np
from scipy.optimize import fsolve
import scipy.stats as st
import xlwings as xw
import datetime as dt
import os

# 初始化
chrome_options = Options()
driver = webdriver.Chrome(chrome_options=chrome_options)
url = 'https://www.sbobet.com/euro'
# admin, pw = 'EURA0706019', 'csl00000'
admin, pw = 'EURA0706015', 'csl00000'
# admin, pw = 'EURA0901223', 'csl00000'
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
    text = [[c1.text_content() for c1 in c2] for c2 in col]

    # 获取让球方

    def get_favor(row):
        favor = [r.xpath('td[@class="BLN team-name-column"]/span') for r in row]
        redblue = [[c1.get('class') for c1 in c2] for c2 in favor]
        team = [[c1.text_content() for c1 in c2] for c2 in favor]
        teamrb = {}
        teamname = []
        for i in range(len(team)):
            t = team[i]
            rb = redblue[i]
            if len(t) > 1:
                if len(t[0]) > 1:
                    teamname = tuple(t)
            j = 3
            if len(rb) > 1:
                if 'Red' in rb:
                    j = rb.index('Red')
                else:
                    j = 2
            if j == 3:
                continue
            if teamname not in teamrb:
                teamrb[teamname] = []
            td = teamrb[teamname]
            if j == 2:
                td.append('PK')
            elif j == 0:
                td.append('Home')
            else:
                td.append('Away')
        return teamrb

    teamrb = get_favor(row)
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
            # print(league)
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
    name_rb_dic = {}
    output = {}
    for tn in list(set(df.team.tolist())):
        for tr in teamrb:
            if tr[0] in tn and tr[1] in tn:
                name_rb_dic[tn] = tr
        df2 = df[df.team == tn].reset_index()
        hllist = []
        for i, j in df2.iterrows():
            tr = name_rb_dic[tn]
            rb = teamrb[tr][i]
            if tr not in output:
                output[tr] = []
            op = output[tr]

            def to_num(x):
                if '-' in str(x):
                    y = x.split('-')
                    z = (float(y[0]) + float(y[1])) / 2
                    return z
                else:
                    return float(x)

            if rb == 'Home':
                haline = to_num(j.HDP)
            elif rb == 'Away':
                haline = - to_num(j.HDP)
            else:
                haline = to_num(j.HDP)
            if i == 0:
                hllist = [to_num(j.Goal), float(j.Over), float(j.Under)]

            op.append([tost([haline, float(j.Home), float(j.Away), hllist[0], hllist[1], hllist[2]]).tolist(),
                      [haline, float(j.Home), float(j.Away), hllist[0], hllist[1], hllist[2]]])

    return output


def read_input_form():
    datedic = {'Mon': '1', 'Tue': '2', 'Wed': '3', 'Thu': '4', 'Fri': '5', 'Sat': '6', 'Sun': '7'}

    def inputform():
        ipfile = []
        for file in os.listdir(os.curdir):
            if file.endswith('.xls') and file.startswith('FB_BOCC Input Form'):
                ipfile.append(file)

        allgame = []
        for ipf in ipfile:
            wb = xw.Book(ipf)
            sht = wb.sheets['Pool']

            # 查找空行
            blank_r = 0
            firstrow = 1
            while blank_r < 4:
                if sht.range(firstrow, 1).value == None:
                    blank_r += 1
                    firstrow += 1
                else:
                    blank_r = 0
                    firstrow += 1
            blank_row = firstrow - 4
            # 检索比赛
            for row in range(1, blank_row):
                game = []
                try:
                    datecode = sht.range(row, 2).value[:3]
                    code2 = sht.range(row, 2).value[3:]
                    if datecode in datedic:
                        code = datedic[datecode] + code2
                        league = sht.range(row, 5).value
                        game.append(code)
                        game.append(league)
                        game.append(sht.range(row, 8).value)
                        game.append(sht.range(row, 10).value)
                        allgame.append(game)
                except:
                    pass
            wb.close()

        agdic = {}
        for ag in allgame:
            league = ag[1]
            if league not in agdic:
                agdic[league] = []
            agdic[league].append(ag)
        gamecode = {}
        for l in agdic:
            games = agdic[l]
            for g in games:
                gamecode[g[0]] = g
        return gamecode

    gamecodedic = inputform()

    def read_oa(gcdic):
        # read oa
        with open('output.txt', 'r', encoding='mbcs') as w:
            data = w.read()
        ylist = data.split('\n')

        for y in ylist:
            if y == '':
                continue
            z = y.split('\t')
            code = datedic[z[1]] + z[2]
            if code in gcdic:
                gcdic[code].append(z[6])

        gamed = []
        for j in gcdic.values():
            if len(j) == 4:
                j.append('')
            gamed.append(j)
        gamed.sort()

        df = pd.DataFrame(gamed)
        df.to_csv('ipoutput.csv', index=False)

        return df

    df1 = read_oa(gamecodedic)

    def select_league(dfs1):
        # 输入联赛
        league = input('Leagues:')
        if '+' in league:
            ls = [l.strip() for l in league.split('+')]
        elif ' ' in league:
            ls = [l.strip() for l in league.strip().split(' ')]
        else:
            ls = [league.strip()]

        dfs2 = dfs1[dfs1[1].isin(ls)]

        weekday = input('Weekday:')
        if ' ' in weekday:
            ws = [l.strip() for l in weekday.strip().split(' ')]
        else:
            ws = [weekday.strip()]

        dfs3 = dfs2[dfs2[0].str[0].isin(ws)].reset_index().iloc[:, 1:]

        return dfs3

    df2 = select_league(df1)
    df2.to_csv('output.csv', index=False)
    return df2


def write():
    wb = xw.Book()
    st = wb.sheets[0]
    df = pd.read_csv('output.csv')
    gap1, gap2 = 3, 9
    o2 = get_st()
    result = []
    for games in o2:
        dfg1 = df[df['2'] == games[0]]
        dfg2 = df[df['3'] == games[1]]
        number0 = ''
        if dfg1.shape[0] == 1:
            number0 = str(dfg1.iloc[0, 0])
        elif dfg2.shape[0] == 1:
            number0 = str(dfg2.iloc[0, 0])
        number = input(' vs '.join(games) + ': ' + number0)
        if number == '':
            number = number0
        if len(number) == 4:
            dfn = df[df['0'] == int(number)]
            if dfn.shape[0] == 0:
                print('%s not in DF' % number)
                continue
            r = dfn.iloc[0]
            result.append(str(r[x]) for x in [str(y) for y in range(5)])



    row = 0
    for s in o2:
        o = o2[s]
        for p in o:
            row += 1
            st.range(row, 1).value = s
            st.range(row, 3).value = p[0]
            st.range(row, 5).value = p[1]






