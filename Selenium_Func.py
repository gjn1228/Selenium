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
import requests
from xlwings.constants import InsertShiftDirection
import configparser


class Sele(object):

    def row_determine(self, date, game, awayteam, ha, at):
        gdf = pd.DataFrame(game, columns=['D', 'T', 'HA', 'C'])
        gdf['Date'] = gdf['D'].map(datetime.datetime.date)
        gdf2 = gdf[datetime.timedelta(-3) < gdf['Date'] - date][
            gdf['Date'] - date < datetime.timedelta(3)]
        row = -1
        if gdf2.shape[0] == 1:
            if gdf2.iloc[0, 1] == awayteam:
                if gdf2.iloc[0, 2] == ['H', 'A'][at.index(ha)]:
                    row = gdf2.index.tolist()[0] + 10
            else:
                ip = input('%s %s %s %s' % (str(gdf2.iloc[0, 4]), gdf2.iloc[0, 1], gdf2.iloc[0, 2], gdf2.iloc[0, 3]))
                if ip == '':
                    row = gdf2.index.tolist()[0] + 10
        elif gdf2.shape[0] > 1:
            gdf3 = gdf2[gdf2['HA'] == ['H', 'A'][at.index(ha)]][gdf2['T'] == awayteam]
            if gdf3.shape[0] == 1:
                row = gdf3.index.tolist()[0] + 10
            else:
                print('Date: %s, Against team: %s, H/A: %s' % (date, awayteam, ['H', 'A'][at.index(ha)]))
                for i in range(gdf2.shape[0]):
                    print('%s :%s %s %s %s' % (i + 1, str(gdf2.iloc[i, 4]), gdf2.iloc[i, 1],
                                               gdf2.iloc[i, 2], gdf2.iloc[i, 3]))
                ip = input('which game')
                try:
                    print()
                    row = gdf2.index.tolist()[int(ip) - 1] + 10
                except:
                    pass
        print('row: %s' % row)

        if row == -1:
            row = int(input('%s row:' % ha))

        return row

    def whoscored_url(slef, url):


        def whoscored_url_team(driver, xpathid):
            i = 1
            Matches = []
            while True:
                try:
                    result = driver.find_element_by_xpath(
                        '//*[@id="%s"]/tbody/tr[%s]/td[@class="result"]/a' % (xpathid, i)).text
                    if result != 'vs':
                        date = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]/td[@class="date"]' % (xpathid, i)).text
                        hteam = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]/td[6]/a' % (xpathid, i)).text
                        ateam = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]/td[8]/a' % (xpathid, i)).text
                        address = driver.find_element_by_xpath(
                            '//*[@id="%s"]/tbody/tr[%s]/td[@class="result"]/a' % (xpathid, i)).get_attribute('href')
                        Matches.append([date, hteam, result, ateam, address])
                    i += 1
                except selenium.common.exceptions.NoSuchElementException:
                    break
                except Exception as inst:
                    print('Game Error')
                    print(type(inst))
                    print(inst.args)
                    i += 1

            return Matches

        def whoscored_url_tournaments(driver, xpathid):

            def whoscored_url_tournaments_main(driver, xpathid):
                i = 1
                Matches = []
                date = ''
                while i < 100:
                    try:
                        tr = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]' % (xpathid, i)).get_attribute(
                            'class')
                        if tr == 'rowgroupheader':
                            date = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]/th' % (xpathid, i)).text
                        else:
                            result = driver.find_element_by_xpath(
                                '//*[@id="%s"]/tbody/tr[%s]/td[@class="result"]/a' % (xpathid, i)).text
                            if result != 'vs':
                                hteam = driver.find_element_by_xpath(
                                    '//*[@id="%s"]/tbody/tr[%s]/td[4]/a' % (xpathid, i)).text
                                ateam = driver.find_element_by_xpath(
                                    '//*[@id="%s"]/tbody/tr[%s]/td[6]/a' % (xpathid, i)).text
                                address = driver.find_element_by_xpath(
                                    '//*[@id="%s"]/tbody/tr[%s]/td[@class="result"]/a' % (xpathid, i)).get_attribute('href')
                                Matches.append([date, hteam, result, ateam, address])
                        i += 1
                    except selenium.common.exceptions.NoSuchElementException:
                        break
                    except Exception as inst:
                        print('Game Error')
                        print(type(inst))
                        print(inst.args)
                        i += 1
                return Matches

            def whoscored_tournaments_preweek(driver):
                # 打开View previous week页面
                starttime = datetime.datetime.now()
                elem = driver.find_element_by_xpath('//*[@id="date-controller"]/a[1]/span')
                elem.click()
                driver.implicitly_wait(3)
                driver.implicitly_wait(4)
                endtime = datetime.datetime.now()
                print('Get Previous Week by ' + str((endtime - starttime).seconds) + 's')
                return driver

            Matches = whoscored_url_tournaments_main(driver, xpathid)
            # 没有比赛时，往回找一页
            if len(Matches) == 0:
                print('No Game, View Previous Week')
                driver2 = whoscored_tournaments_preweek(driver)
                Matches = whoscored_url_tournaments_main(driver2, xpathid)
            pre = 'P'
            print(Matches)
            # 手动往回找
            while pre == 'P':
                pre = input('Input "P" to View Previous Week, Else Continue: ')
                if pre == 'P':
                    print('Loading')
                    driver2 = whoscored_tournaments_preweek(driver)
                    Matches = whoscored_url_tournaments_main(driver2, xpathid)
            return Matches


        matches = []
        games = []
        if url.find('Matches') != -1:
            games = [url]
        else:
            starttime = datetime.datetime.now()
            # Chrome Headless
            chrome_options = Options()
            #chrome_options.add_argument('--headless')
            chrome_options.add_argument('--disable-gpu')
            #chrome_options.add_extension('Adblock-Plus_v1.8.12.crx')
            driver = webdriver.Chrome(chrome_options=chrome_options)
            driver.get(url)
            driver.implicitly_wait(3)
            endtime = datetime.datetime.now()
            print('Urltime = ' + str((endtime - starttime).seconds) + 's')
            if url.find('Teams') != -1:
                if url.find('Fixtures') != -1:
                    xpathid = 'team-fixtures'
                    matches = whoscored_url_team(driver, xpathid)
                elif url.find('Show') != -1:
                    xpathid = 'team-fixtures-summary'
                    matches = whoscored_url_team(driver, xpathid)
            elif url.find('Tournaments') != -1:
                xpathid = 'tournament-fixture'
                matches = whoscored_url_tournaments(driver, xpathid)
            else:
                return 'Error'
        if games != []:
            return games
        else:
            print(matches)
            print('-------------------------------------------')
            for match in matches:
                print('No. ' + str(matches.index(match) + 1) + ' | ' + match[0] + ' ' + match[1] + ' ' + match[2] + ' ' + match[3])
            print('-------------------------------------------')
            start_num = input('Begin by Which Game(No.):')
            end_num = input('End by Which Game(No.):')

            if start_num == '':
                start_num = 1
            if end_num == '':
                end_num = len(matches) + 2
            for match in matches[int(start_num) - 1:int(end_num)]:
                games.append(match[4])
            return games



        endtime = datetime.datetime.now()
        print('Alltime = ' + str((endtime - starttime).seconds) + 's')

        driver.quit()

    # 181122,读取页面太慢会导致信息不全，目前采用检查长度，如果信息不全就隔2秒重新读一遍网页
    def check_html(self, driver):
        yy = []
        times = 0
        while len(yy) < 1:
        # while times < 4:
            if times != 0:
                print('Retrying')
                time.sleep(2)
            times += 1
            yy = []
            html = driver.page_source
            tree = lxml.html.fromstring(html)
            bs = BeautifulSoup(html, 'html.parser')
            x = bs.find_all('script', type="text/javascript")
            for xxx in x:
                if 30000 > len(xxx.text) > 8000:
                    yy.append(xxx)
            # 181122 目前用长度定位，留下格式信息备用
            # print(len(yy))
            # print([len(xxx.text) for xxx in yy])
            # print([z.text[:100] for z in yy])
            # [5266, 57363, 10895]
            # [5266, 57345, 11283]
            # ["\n    var matchStats = [[[167,32,'Manchester City','Manchester United','11/11/2018 16:30:00','11/11/2"]

        return html

    def whsc_match(self, htmltext):
        tree = lxml.html.fromstring(htmltext)
        bs = BeautifulSoup(htmltext, 'html.parser')
        x = bs.find_all('script', type="text/javascript")
        yy = []
        # 181112 经过观察，目标字段长度在10000-12000之间，缩小范围以精确定位
        for xxx in x:
            if 30000 > len(xxx.text) > 8000:
            # if len(xxx.text) > 5000:
                yy.append(xxx)
        y = yy[0].text
        z = y.split(';')[0][23:]
        u = z.replace('\'', '"').replace(',,', ',null,').replace(',,', ',null,').replace(',]', ',null]')
        uu = ']'.join(u.split(']')[:-2]) + ']'
        v = json.loads(uu)
        # # 181022 似乎多了一行，y定位从1改到-1
        # try:
        #     print(1)
        #     y = yy[-1].text
        #     z = y.split(';')[0][23:]
        #     u = z.replace('\'', '"').replace(',,', ',null,').replace(',,', ',null,').replace(',]', ',null]')
        #     uu = ']'.join(u.split(']')[:-2]) + ']'
        #     v = json.loads(uu)
        # except:
        #     print(2)
        #     y = yy[1].text
        #     z = y.split(';')[0][23:]
        #     u = z.replace('\'', '"').replace(',,', ',null,').replace(',,', ',null,').replace(',]', ',null]')
        #     uu = ']'.join(u.split(']')[:-2]) + ']'
        #     v = json.loads(uu)

        tr = tree.xpath('//*[@id="breadcrumb-nav"]/a')[0].text.replace('\r\n', '').strip()
        return v, tr

    def whsc_match_summary(self, htmltext):
        tree = lxml.html.fromstring(htmltext)
        bs = BeautifulSoup(htmltext, 'html.parser')
        x = bs.find_all('script', type="text/javascript")
        yy = []
        for xxx in x:
            if len(xxx.text) > 2000:
                yy.append(xxx)
        y = yy[-1].text
        z = y.split('initialMatchDataForScrappers = ')[1].split(';')[0]
        u = z.replace('\'', '"').replace(',,', ',null,').replace(',,', ',null,').replace(',]', ',null]')
        v = json.loads(u)
        tr = tree.xpath('//*[@id="breadcrumb-nav"]/a')[0].text.replace('\r\n', '').strip()
        return v, tr

    def r_to_draft_summary(self, w):
        league = w[1]

        gdic = {}
        glist = w[0][0][0]
        gdic['hteam'] = glist[2]
        gdic['ateam'] = glist[3]
        gdic['datetime'] = glist[4]
        gdic['date'] = glist[5]
        gdic['gametime'] = glist[7]
        gdic['score'] = glist[9]
        gdic['halftime-score'] = glist[8]

        tl = w[0][0][1][0]
        tl2 = []
        for event in tl:
            if event[3] == 1:
                ttime = [event[0]]
                te = event[1]
                ttime.append('H')
                for tee in te:
                    ttl = []
                    ttl.extend(ttime)
                    ttl.extend(tee)
                    tl2.append(ttl)
            if event[4] == 1:
                ttime = [event[0]]
                te = event[2]
                ttime.append('A')
                for tee in te:
                    ttl = []
                    ttl.extend(ttime)
                    ttl.extend(tee)
                    tl2.append(ttl)
        dft = pd.DataFrame(tl2)

        pl = w[0][0][2]
        hteam = glist[2]
        ateam = glist[3]
        hscore = gdic['score'].split(':')[0].strip()
        ascore = gdic['score'].split(':')[1].strip()
        hco, aco, hcolor, acolor, hog, aog = [], [], [], [], [], []
        score = (0, 0)
        for it in dft.iterrows():
            red = 10
            i = it[1]
            time = i[7]
            if i[1] == 'H':
                r = hco
                c = hcolor
                og = hog
            else:
                r = aco
                c = acolor
                og = aog
                score = (score[1], score[0])
            if i[4] == 'red':
                event = '%s men @ %s\' when %s-%s' % (red, time, score[0], score[1])
                r.append(event)
                c.append('red')
                red -= 1
            elif i[4] == 'secondyellow':
                event = '%s men @ %s\' when %s-%s' % (red, time, score[0], score[1])
                r.append(event)
                c.append('red')
                red -= 1
            elif i[4] == 'penalty-goal':
                event = 'PG @ %s\' when %s-%s' % (time, score[0], score[1])
                r.append(event)
                score = (i[5][1], i[5][3])
            elif i[4] == 'penalty-missed':
                event = 'PK Missed @ %s\' when %s-%s' % (time, score[0], score[1])
                r.append(event)
            elif i[4] == 'owngoal':
                event = 'OG @ %s\' when %s-%s' % (time, score[0], score[1])
                r.append(event)
                og.append(1)
                score = (i[5][1], i[5][3])
            elif i[4] == 'goal':
                if 85 <= time <= 90:
                    event = 'LG @ %s\' when %s-%s' % (time, score[0], score[1])
                    r.append(event)
                score = (i[5][1], i[5][3])

        if gdic['gametime'] == 'AET':
            hco.append('Extratime %s' % glist[10])
        elif gdic['gametime'] != 'FT':
            print('gametime error: %s' % gdic['gametime'])

        hplayer, aplayer = {}, {}
        for player in [[hplayer, pl[9], pl[11], hteam], [aplayer, pl[10], pl[12], ateam]]:
            pdic = player[0]
            for pla in player[1]:
                name = pla[0]
                xs = ['X']
                font = ['', '', '', '']
                comment = []
                goal = 0
                exgoal = 0
                ytime = 0
                if len(pla[2]) > 0:
                    for eve in pla[2]:
                        t = eve[5]
                        if eve[2] == 'subst-out':
                            xs.insert(1, '-')
                            if t <= 45:
                                comment.append('Early Subout @ %s\'' % t)
                        elif eve[2] == 'goal':
                            font[0] = 'Bold'
                            if t > 90:
                                exgoal += 1
                            else:
                                goal += 1
                        elif eve[2] == 'penalty-goal':
                            if t > 90:
                                comment.append('Extra time PG @ %s\'' % t)
                                exgoal += 1
                            else:
                                font[0] = 'Bold'
                                comment.append('PG @ %s\'' % t)
                                goal += 1
                        elif eve[2] == 'goalown':
                            comment.append('OG @ %s\'' % t)
                        elif eve[2] == 'yellow':
                            ytime = t
                            font[2] = 'yellow'
                        elif eve[2] == 'secondyellow':
                            font[2] = 'red'
                            comment.append('2Y @ %s\',%s\'' % (ytime, t))
                        elif eve[2] == 'red':
                            font[2] = 'red'
                            comment.append('1R @ %s\'' % t)
                        else:
                            print('event error')
                            print(eve)
                if goal > 0:
                    xs.append(str(goal))

                if exgoal > 0:
                    comment.append('Extratime %s Goal' % exgoal)
                pdic[name] = [player[3], xs, font, comment]
            for pla in player[2]:
                name = pla[0]
                xs = ['b']
                font = ['', '', '', '']
                comment = []
                goal = 0
                ytime = 0
                if len(pla[2]) > 0:
                    for eve in pla[2]:
                        t = eve[5]
                        if eve[2] == 'subst-in':
                            xs[0] = 's'
                            if t <= 45:
                                xs[0] = 'S'
                                font[1] = 'red'
                        elif eve[2] == 'subst-out':
                            xs[0] = 'S'
                            xs.insert(1, '-')
                            font[1] = 'red'
                        elif eve[2] == 'goal':
                            font[0] = 'Bold'
                            if t > 90:
                                exgoal += 1
                            else:
                                goal += 1
                        elif eve[2] == 'penalty-goal':
                            if t > 90:
                                comment.append('Extra time PG @ %s\'' % t)
                                exgoal += 1
                            else:
                                font[0] = 'Bold'
                                comment.append('PG @ %s\'' % t)
                                goal += 1
                        elif eve[2] == 'goalown':
                            comment.append('OG @ %s\'' % t)
                        elif eve[2] == 'yellow':
                            ytime = t
                            font[2] = 'yellow'
                        elif eve[2] == 'secondyellow':
                            font[2] = 'red'
                            comment.append('2Y @ %s\',%s\'' % (ytime, t))
                        elif eve[2] == 'red':
                            font[2] = 'red'
                            comment.append('1R @ %s\'' % t)
                        else:
                            print('event error')
                            print(eve)
                    if goal > 0:
                        xs.append(str(goal))

                    if exgoal > 0:
                        comment.append('Extratime %s Goal' % exgoal)

                pdic[name] = [player[3], xs, font, comment]

        return [[hteam, ateam, league, gdic['datetime']], [[[hscore, hco, hcolor], [ascore, aco, acolor]], hplayer, len(hog)],
                [[[ascore, aco, acolor], [hscore, hco, hcolor]], aplayer, len(aog)]]

    def r_to_df(self, x):
        # 联赛名
        league = x[1]

        y = x[0]
        # 比赛信息
        glist = y[0]
        gdic = {}
        gdic['hteam'] = glist[2]
        gdic['ateam'] = glist[3]
        gdic['datetime'] = glist[4]
        gdic['date'] = glist[5]
        gdic['gametime'] = glist[7]
        gdic['score'] = glist[9]
        # timeline
        tl = y[1][0]
        tl2 = []
        for event in tl:
            if event[3] == 1:
                ttime = [event[0]]
                te = event[1]
                ttime.append('H')
                for tee in te:
                    ttl = []
                    ttl.extend(ttime)
                    ttl.extend(tee)
                    tl2.append(ttl)
            if event[4] == 1:
                ttime = [event[0]]
                te = event[2]
                ttime.append('A')
                for tee in te:
                    ttl = []
                    ttl.extend(ttime)
                    ttl.extend(tee)
                    tl2.append(ttl)

        dft = pd.DataFrame(tl2)
        # 球员
        pl = y[2]
        player = []
        for p1 in pl:
            HA = ['H', 'A'][pl.index(p1)]
            team = p1[1]
            dfs = pd.DataFrame(p1[3][0]).set_index(0).to_dict()[1]
            fm = p1[5][0]
            dfp = pd.DataFrame(p1[4])
            player.append([HA, team, fm, dfs, dfp])
        return league, gdic, dft, player

    def df_to_draft(self, x):
        hteam = x[1]['hteam']
        ateam = x[1]['ateam']
        hscore = x[1]['score'].split(':')[0].strip()
        ascore = x[1]['score'].split(':')[1].strip()
        hco, aco, hcolor, acolor = [], [], [], []
        score = (0, 0)
        pe = {}
        for it in x[2].iterrows():
            red = 10
            i = it[1]
            time = i[7]
            ply_eve = []
            if i[1] == 'H':
                r = hco
                c = hcolor
            else:
                r = aco
                c = acolor
                score = (score[1], score[0])
            if i[4] == 'red':
                event = '%s men @ %s\' when %s-%s' % (red, time, score[0], score[1])
                r.append(event)
                c.append('red')
                ply_eve.append('red')
                ply_eve.append('1R @  %s\' when %s-%s'% (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [(ply_eve)]
                red -= 1
            elif i[4] == 'secondyellow':
                event = '%s men @ %s\' when %s-%s' % (red, time, score[0], score[1])
                ply_eve.append('red')
                ytime = x[2][x[2][2] == i[2]][x[2][4] == 'yellow'][0].tolist()[0]
                ply_eve.append('2Y @  %s\', %s\' when %s-%s'% (ytime, time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [ply_eve]
                r.append(event)
                c.append('red')
                red -= 1
            elif i[4] == 'penalty-goal':
                event = 'PG @ %s\' when %s-%s' % (time, score[0], score[1])
                ply_eve.append('PG')
                ply_eve.append('PG @  %s\', when %s-%s' % (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [ply_eve]
                pe[i[2]].append(['goal'])
                r.append(event)
                score = (i[5][1], i[5][3])
            elif i[4] == 'penalty-missed':
                event = 'PK Missed @ %s\' when %s-%s' % (time, score[0], score[1])
                ply_eve.append('PM')
                ply_eve.append('PK Missed @  %s\', when %s-%s' % (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [ply_eve]
                r.append(event)
            elif i[4] == 'owngoal':
                event = 'OG @ %s\' when %s-%s' % (time, score[0], score[1])
                ply_eve.append('OG')
                ply_eve.append('OG @  %s\', when %s-%s' % (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [ply_eve]
                r.append(event)
                score = (i[5][1], i[5][3])
            elif i[4] == 'goal':
                if time >= 85:
                    event = 'LG @ %s\' when %s-%s' % (time, score[0], score[1])
                    r.append(event)
                if i[2] in pe:
                    pe[i[2]].append(['goal'])
                else:
                    pe[i[2]] = [['goal']]
                if i[3] is not None:
                    if i[3] in pe:
                        pe[i[3]].append(['assist'])
                    else:
                        pe[i[3]] = [['assist']]
                score = (i[5][1], i[5][3])
            elif i[4] == 'subst':
                if i[2] in pe:
                    pe[i[2]].append(['out', time])
                else:
                    pe[i[2]] = [['out', time]]
                if i[3] in pe:
                    pe[i[3]].append(['in', time])
                else:
                    pe[i[3]] = [['in', time]]
                if time < 45:
                    if i[2] in pe:
                        pe[i[2]].append(['early out', time])
                    else:
                        pe[i[2]] = [['early out', time]]
                    if i[3] in pe:
                        pe[i[3]].append(['early in', time])
                    else:
                        pe[i[3]] = [['early in', time]]

        pee = {}
        for pp in pe:
            ppe = []
            for ee in pe[pp]:
                ppe.append(ee[0])
            pee[pp] = ppe
        dfp = pd.concat([x[3][0][4], x[3][1][4]])
        dfp2 = dfp.sort_values(by=2, ascending=False).reset_index().iloc[:5]
        top = dfp2[1].tolist()
        mvp = dfp2.loc[0, 1]
        hp, ap = {}, {}
        pid = {}
        for tm in x[3]:
            pld = [hp, ap][x[3].index(tm)]
            for it in tm[4].iterrows():
                i = it[1]
                name = i[1]
                pdic = {}
                xs = []
                font = ['', '', '', '']
                comment = []
                for j in i[3][0]:
                    pdic[j[0]] = j[1][0]

                if pdic['formation_place'] != 0:
                    xs.append('X')
                elif name in pee:
                    if 'in' in pee[name]:
                        if 'early in ' in pee[name] or 'out' in pee[name]:
                            xs.append('S')
                            font[1] = 'red'
                        else:
                            xs.append('s')
                    else:
                        xs.append('b')
                else:
                    xs.append('b')

                if name in pee:
                    if 'out' in pee[name]:
                        xs.append('-')
                        if 'in' in pee[name]:
                            tin = pe[name][pee[name].index('in')][1]
                            tout = pe[name][pee[name].index('out')][1]
                            comment.append('Sub in @ %s, Sub out @ %s' % (tin, tout))
                        elif 'early out' in pee[name]:
                            tout = pe[name][pee[name].index('out')][1]
                            comment.append('Early sub @ %s' % tout)

                if 'goals' in pdic:
                    xs.append(str(pdic['goals']))
                    font[0] = 'Bold'
                    if 'PG' in pee[name]:
                        comment.append(pe[name][pee[name].index('PG')][1])

                if 'goal_assist' in pdic:
                    xs.append('A')
                    if pdic['goal_assist'] > 1:
                        xs.append(str(pdic['goal_assist']))

                if 'yellow_card' in pdic:
                    font[2] = 'yellow'

                if name in pee:
                    if 'red' in pee[name]:
                        font[2] = 'red'
                        comment.append(pe[name][pee[name].index('red')][1])

                    if 'OG' in pee[name]:
                        comment.append(pe[name][pee[name].index('OG')][1])

                    if 'PM' in pee[name]:
                        comment.append(pe[name][pee[name].index('PM')][1])

                if name == mvp:
                    font[0] = 'Bold'
                    font[3] = 'I'
                    font[1] = 'red'
                elif name in top:
                    font[1] = 'blue'
                pld[name] = [i[2], xs, font, comment]
                pid[name] = i[6]

        hshot, ashot = [], []
        for shot in x[3]:
            st = [hshot, ashot][x[3].index(shot)]
            sdic = shot[3]
            if 'ontarget_scoring_att' in sdic:
                so = int(sdic['ontarget_scoring_att'][0])
            else:
                so = 0
            if 'blocked_scoring_att' in sdic:
                sb = int(sdic['blocked_scoring_att'][0])
            else:
                sb = 0
            if 'shot_off_target' in sdic:
                sf = int(sdic['shot_off_target'][0])
            else:
                sf = 0
            st.append(so+ sb + sf)
            st.append(so)
            if int(sdic['possession_percentage'][0]) - sdic['possession_percentage'][0] == 0.5:
                st.append(int(sdic['possession_percentage'][0]) + x[3].index(shot))
            else:
                st.append(round(sdic['possession_percentage'][0]))
            st.append(shot[2])
        hss, hso, hpo, ass, aso, apo, hf, af = hshot[0], hshot[1], hshot[2], ashot[0], ashot[1], ashot[2], hshot[3], ashot[3]
        homeshot = [hss, ass, hso, aso, hpo, apo, hf]
        awayshot = [ass, hss, aso, hso, apo, hpo, af]

        return [[hteam, ateam, x[0], x[1]['datetime']], [[[hscore, hco, hcolor], [ascore, aco, acolor]], homeshot, [hp, pid]],
                [[[ascore, aco, acolor], [hscore, hco, hcolor]], awayshot, [ap, pid]]]

    # 180903调整og顺序， 180923客队红牌人数
    def df_to_draft2(self, x, tl):
        hteam = x[1]['hteam']
        ateam = x[1]['ateam']
        hscore = x[1]['score'].split(':')[0].strip()
        ascore = x[1]['score'].split(':')[1].strip()
        hco, aco, hcolor, acolor = [], [], [], []
        hred, ared = 10, 10
        score = (0, 0)
        pe = {}
        df = x[2]
        df[7] = tl
        for it in df.iterrows():
            i = it[1]
            time = i[7]
            time2 = int(time.split('+')[0])
            ply_eve = []
            if i[1] == 'H':
                r = hco
                c = hcolor
            else:
                r = aco
                c = acolor
                score = (score[1], score[0])
            if i[4] == 'red':
                if i[1] == 'H':
                    red = hred
                    hred -= 1
                else:
                    red = ared
                    ared -= 1
                event = '%s men @ %s\' when %s-%s' % (red, time, score[0], score[1])
                r.append(event)
                c.append('red')
                ply_eve.append('red')
                ply_eve.append('1R @  %s\' when %s-%s' % (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [(ply_eve)]
            elif i[4] == 'secondyellow':
                if i[1] == 'H':
                    red = hred
                    hred -= 1
                else:
                    red = ared
                    ared -= 1
                event = '%s men @ %s\' when %s-%s' % (red, time, score[0], score[1])
                ply_eve.append('red')
                ytime = x[2][x[2][2] == i[2]][x[2][4] == 'yellow'][0].tolist()[0]
                ply_eve.append('2Y @  %s\', %s\' when %s-%s' % (ytime, time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [ply_eve]
                r.append(event)
                c.append('red')
                red -= 1
            elif i[4] == 'penalty-goal':
                event = 'PG @ %s\' when %s-%s' % (time, score[0], score[1])
                ply_eve.append('PG')
                ply_eve.append('PG @  %s\', when %s-%s' % (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [ply_eve]
                pe[i[2]].append(['goal'])
                r.append(event)
                score = (i[5][1], i[5][3])
            elif i[4] == 'penalty-missed':
                event = 'PK Missed @ %s\' when %s-%s' % (time, score[0], score[1])
                ply_eve.append('PM')
                ply_eve.append('PK Missed @  %s\', when %s-%s' % (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [ply_eve]
                r.append(event)
            elif i[4] == 'owngoal':
                event = 'OG @ %s\' when %s-%s' % (time, score[0], score[1])
                ply_eve.append('OG')
                ply_eve.append('OG @  %s\', when %s-%s' % (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [ply_eve]
                r.append(event)
                score = (i[5][1], i[5][3])
            elif i[4] == 'goal':
                if time2 >= 85:
                    event = 'LG @ %s\' when %s-%s' % (time, score[0], score[1])
                    r.append(event)
                if i[2] in pe:
                    pe[i[2]].append(['goal'])
                else:
                    pe[i[2]] = [['goal']]
                if i[3] is not None:
                    if i[3] in pe:
                        pe[i[3]].append(['assist'])
                    else:
                        pe[i[3]] = [['assist']]
                score = (i[5][1], i[5][3])
            elif i[4] == 'subst':
                if i[2] in pe:
                    pe[i[2]].append(['out', time])
                else:
                    pe[i[2]] = [['out', time]]
                if i[3] in pe:
                    pe[i[3]].append(['in', time])
                else:
                    pe[i[3]] = [['in', time]]
                if time2 < 45:
                    if i[2] in pe:
                        pe[i[2]].append(['early out', time])
                    else:
                        pe[i[2]] = [['early out', time]]
                    if i[3] in pe:
                        pe[i[3]].append(['early in', time])
                    else:
                        pe[i[3]] = [['early in', time]]

        pee = {}
        for pp in pe:
            ppe = []
            for ee in pe[pp]:
                ppe.append(ee[0])
            pee[pp] = ppe
        dfp = pd.concat([x[3][0][4], x[3][1][4]])
        dfp2 = dfp.sort_values(by=2, ascending=False).reset_index().iloc[:5]
        top = dfp2[1].tolist()
        mvp = dfp2.loc[0, 1]
        hp, ap = {}, {}
        pid = {}
        allog = []
        for tm in x[3]:
            pld = [hp, ap][x[3].index(tm)]
            og = 0
            for it in tm[4].iterrows():
                i = it[1]
                name = i[1]
                pdic = {}
                xs = []
                font = ['', '', '', '']
                comment = []
                for j in i[3][0]:
                    pdic[j[0]] = j[1][0]

                if pdic['formation_place'] != 0:
                    xs.append('X')
                elif name in pee:
                    if 'in' in pee[name]:
                        if 'early in ' in pee[name] or 'out' in pee[name]:
                            xs.append('S')
                            font[1] = 'red'
                        else:
                            xs.append('s')
                    else:
                        xs.append('b')
                else:
                    xs.append('b')

                if name in pee:
                    if 'out' in pee[name]:
                        xs.append('-')
                        if 'in' in pee[name]:
                            tin = pe[name][pee[name].index('in')][1]
                            tout = pe[name][pee[name].index('out')][1]
                            comment.append('Sub in @ %s, Sub out @ %s' % (tin, tout))
                        elif 'early out' in pee[name]:
                            tout = pe[name][pee[name].index('out')][1]
                            comment.append('Early sub @ %s' % tout)

                if 'goals' in pdic:
                    xs.append(str(pdic['goals']))
                    font[0] = 'Bold'
                    if 'PG' in pee[name]:
                        comment.append(pe[name][pee[name].index('PG')][1])

                if 'goal_assist' in pdic:
                    xs.append('A')
                    if pdic['goal_assist'] > 1:
                        xs.append(str(pdic['goal_assist']))


                if 'yellow_card' in pdic:
                    font[2] = 'yellow'

                if name in pee:
                    if 'red' in pee[name]:
                        font[2] = 'red'
                        comment.append(pe[name][pee[name].index('red')][1])

                    if 'OG' in pee[name]:
                        og += 1
                        comment.append(pe[name][pee[name].index('OG')][1])

                    if 'PM' in pee[name]:
                        comment.append(pe[name][pee[name].index('PM')][1])

                if name == mvp:
                    font[0] = 'Bold'
                    font[3] = 'I'
                    font[1] = 'red'
                elif name in top:
                    font[1] = 'blue'
                pld[name] = [i[2], xs, font, comment]
                pid[name] = i[6]
            allog.append(og)

        hshot, ashot = [], []
        for shot in x[3]:
            st = [hshot, ashot][x[3].index(shot)]
            sdic = shot[3]
            if 'ontarget_scoring_att' in sdic:
                so = int(sdic['ontarget_scoring_att'][0])
            else:
                so = 0
            if 'blocked_scoring_att' in sdic:
                sb = int(sdic['blocked_scoring_att'][0])
            else:
                sb = 0
            if 'shot_off_target' in sdic:
                sf = int(sdic['shot_off_target'][0])
            else:
                sf = 0
            st.append(so + sb + sf)
            st.append(so)
            if int(sdic['possession_percentage'][0]) - sdic['possession_percentage'][0] == 0.5:
                st.append(int(sdic['possession_percentage'][0]) + x[3].index(shot))
            else:
                st.append(round(sdic['possession_percentage'][0]))
            st.append(shot[2])
        hss, hso, hpo, ass, aso, apo, hf, af = hshot[0], hshot[1], hshot[2], ashot[0], ashot[1], ashot[2], hshot[3], ashot[
            3]
        homeshot = [hss, ass, hso, aso, hpo, apo, hf]
        awayshot = [ass, hss, aso, hso, apo, hpo, af]
        hog = allog[1]
        aog = allog[0]

        return [[hteam, ateam, x[0], x[1]['datetime']], [[[hscore, hco, hcolor], [ascore, aco, acolor]], homeshot, [hp, pid], hog],
                [[[ascore, aco, acolor], [hscore, hco, hcolor]], awayshot, [ap, pid], aog]]

    def draft_to_epl(self, x):
        # path = r'E:\Company\League\EPL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Whoscored', 'Soccerway']].set_index('Whoscored').to_dict()['Soccerway']
        wsc_ps_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal_Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Man_City',
     'Manchester United': 'Man_Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West_Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wsc_game_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Manchester City',
     'Manchester United': 'Manchester Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wb = xw.Book('EPL_Playersheet_2018-2019.xlsm')
        hteam = x[0][0]
        ateam = x[0][1]
        at = [hteam, ateam]
        for ha in at:
            st = wb.sheets[wsc_ps_dic[ha]]
            game = st.range((10, 1), (80, 4)).value
            row = -1
            for g in game:
                if g[3] == 'PRL':
                    if g[2] == ['H', 'A'][at.index(ha)]:
                        if g[1] == wsc_game_dic[at[1 - at.index(ha)]]:
                            row = game.index(g) + 10
            hal = x[at.index(ha) + 1]

            self.to_sheet_all(st, row, hal, ha)

    def draft_to_pfl(self, x):
        # path = r'E:\Company\League\EPL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Whoscored', 'Soccerway']].set_index('Whoscored').to_dict()['Soccerway']
        # psd = df[['Whoscored', 'PS']].set_index('Whoscored').to_dict()['PS']
        wsc_ps_dic = {'Aves': 'Desportivo_Aves',
     'Belenenses': 'Belenenses',
     'Belenenses SAD': 'Belenenses',
     'Benfica': 'Benfica',
     'Boavista': 'Boavista',
     'Braga': 'Braga',
     'Chaves': 'Chaves',
     'FC Porto': 'Porto',
     'Feirense': 'Feirense',
     'Maritimo': 'Marítimo',
     'Moreirense': 'Moreirense',
     'Nacional': 'Nacional',
     'Portimonense': 'Portimonense',
     'Rio Ave': 'Rio_Ave',
     'Santa Clara': 'Santa_Clara',
     'Sporting CP': 'Sporting_CP',
     'Tondela': 'Tondela',
     'Vitoria de Guimaraes': 'Vitória_Guimarães',
     'Vitoria de Setubal': 'Vitória_Setúbal'}
        wsc_game_dic = {'Aves': 'Desportivo Aves',
     'Belenenses': 'Belenenses',
     'Belenenses SAD': 'Belenenses',
     'Benfica': 'Benfica',
     'Boavista': 'Boavista',
     'Braga': 'Sporting Braga',
     'Chaves': 'Chaves',
     'FC Porto': 'Porto',
     'Feirense': 'Feirense',
     'Maritimo': 'Marítimo',
     'Moreirense': 'Moreirense',
     'Nacional': 'Nacional',
     'Portimonense': 'Portimonense',
     'Rio Ave': 'Rio Ave',
     'Santa Clara': 'Santa Clara',
     'Sporting CP': 'Sporting CP',
     'Tondela': 'Tondela',
     'Vitoria de Guimaraes': 'Vitória Guimarães',
     'Vitoria de Setubal': 'Vitória Setúbal'}
        wb = xw.Book('PFL_Player_Sheet_18-19.xlsm')
        hteam = x[0][0]
        ateam = x[0][1]
        at = [hteam, ateam]
        for ha in at:
            st = wb.sheets[wsc_ps_dic[ha]]
            game = st.range((10, 1), (80, 4)).value
            row = -1
            for g in game:
                if g[3] == 'PRL':
                    if g[2] == ['H', 'A'][at.index(ha)]:
                        if g[1] == wsc_game_dic[at[1 - at.index(ha)]]:
                            row = game.index(g) + 10

            hal = x[at.index(ha) + 1]
            self.to_sheet_pfl(st, row, hal, ha)

    def draft_to_ps(self, x):
        # path = r'E:\Company\League\EPL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Whoscored', 'Soccerway']].set_index('Whoscored').to_dict()['Soccerway']
        wsc_ps_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal_Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Man_City',
     'Manchester United': 'Man_Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West_Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wsc_game_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Manchester City',
     'Manchester United': 'Manchester Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wb = xw.Book('EPL_Playersheet_2018-2019.xlsm')
        hteam = x[0][0]
        ateam = x[0][1]
        at = [hteam, ateam]
        for ha in at:
            st = wb.sheets[wsc_ps_dic[ha]]
            game = st.range((10, 1), (80, 4)).value
            row = -1
            for g in game:
                if g[3] == 'PRL':
                    if g[2] == ['H', 'A'][at.index(ha)]:
                        if g[1] == wsc_game_dic[at[1 - at.index(ha)]]:
                            row = game.index(g) + 10
            plist = st.range((8, 29), (8, 80)).value
            nlist = st.range((6, 29), (6, 80)).value
            result = st.range((row, 29), (row, 80)).value
            st.range((row, 29), (row, 80)).api.Font.Bold = False
            st.range((row, 29), (row, 80)).api.Font.ColorIndex = 1
            hal = x[at.index(ha) + 1]
            home = hal[0][0]
            away = hal[0][1]
            st.range(row, 5).value = home[0]
            st.range(row, 6).value = away[0]
            if len(home[1]) > 0:
                try:
                    st.range(row, 5).api.AddComment('\n'.join(home[1]))
                except:
                    st.range(row, 5).api.Comment.Text('\n'.join(home[1]))
                st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
            if len(away[1]) > 0:
                try:
                    st.range(row, 6).api.AddComment('\n'.join(away[1]))
                except:
                    st.range(row, 6).api.Comment.Text('\n'.join(away[1]))
                st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
            if home[2] == ['red']:
                st.range(row, 5).api.Interior.Color = 255
            if away[2] == ['red']:
                st.range(row, 6).api.Interior.Color = 255

            st.range((row, 19), (row, 25)).value = hal[1]

            for name in hal[2][0]:
                ind = 0
                data = hal[2][0][name]
                if name in plist:
                    ind = plist.index(name)
                elif hal[2][1][name] in nlist:
                    ind = nlist.index(hal[2][1][name])

                if ind == 0:
                    print(ha, '-', name, 'error')
                else:
                    result[ind] = ''.join(data[1])
                    col = ind + 29
                    if data[2][0] == 'Bold':
                        st.range(row, col).api.Font.Bold = True
                    if data[2][1] == 'red':
                        st.range(row, col).api.Font.ColorIndex = 3
                    if data[2][1] == 'blue':
                        st.range(row, col).api.Font.ColorIndex = 5
                    if data[2][2] == 'yellow':
                        st.range(row, col).api.Interior.Color = 65535
                    if data[2][2] == 'red':
                        st.range(row, col).api.Interior.Color = 255
                    if data[2][3] == 'I':
                        st.range(row, col).api.Font.Italic = True
                    if len(data[3]) > 0:
                        try:
                            st.range(row, col).api.AddComment('\n'.join(data[3]))
                        except:
                            st.range(row, col).api.Comment.Text('\n'.join(data[3]))
                        st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                        st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'

            st.range((row, 29), (row, 80)).value = result
            print('%s Finished' % ha)

    def draft_to_pfl_summary(self, x):
        # path = r'E:\Company\League\EPL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Whoscored', 'Soccerway']].set_index('Whoscored').to_dict()['Soccerway']
        # psd = df[['Whoscored', 'PS']].set_index('Whoscored').to_dict()['PS']
        wsc_ps_dic = {'Aves': 'Desportivo_Aves',
     'Belenenses': 'Belenenses',
     'Benfica': 'Benfica',
     'Boavista': 'Boavista',
     'Braga': 'Braga',
     'Chaves': 'Chaves',
     'FC Porto': 'Porto',
     'Feirense': 'Feirense',
     'Maritimo': 'Marítimo',
     'Moreirense': 'Moreirense',
     'Nacional': 'Nacional',
     'Portimonense': 'Portimonense',
     'Rio Ave': 'Rio_Ave',
     'Santa Clara': 'Santa_Clara',
     'Sporting CP': 'Sporting_CP',
     'Tondela': 'Tondela',
     'Vitoria de Guimaraes': 'Vitória_Guimarães',
     'Vitoria de Setubal': 'Vitória_Setúbal'}
        wsc_game_dic = {'Aves': 'Desportivo Aves',
     'Belenenses': 'Belenenses',
     'Benfica': 'Benfica',
     'Boavista': 'Boavista',
     'Braga': 'Sporting Braga',
     'Chaves': 'Chaves',
     'FC Porto': 'Porto',
     'Feirense': 'Feirense',
     'Maritimo': 'Marítimo',
     'Moreirense': 'Moreirense',
     'Nacional': 'Nacional',
     'Portimonense': 'Portimonense',
     'Rio Ave': 'Rio Ave',
     'Santa Clara': 'Santa Clara',
     'Sporting CP': 'Sporting CP',
     'Tondela': 'Tondela',
     'Vitoria de Guimaraes': 'Vitória Guimarães',
     'Vitoria de Setubal': 'Vitória Setúbal'}
        wb = xw.Book('PFL_Player_Sheet_18-19.xlsm')
        hteam = x[0][0]
        ateam = x[0][1]
        at = [hteam, ateam]
        for ha in at:
            if ha in wsc_ps_dic:
                st = wb.sheets[wsc_ps_dic[ha]]

                date = datetime.datetime.strptime(x[0][3], '%d/%m/%Y %H:%M:%S').date()
                awayteam = at[1 - at.index(ha)]
                if awayteam in wsc_game_dic:
                    awayteam = wsc_game_dic[awayteam]
                game = st.range((10, 1), (80, 4)).value

                row = self.row_determine(date, game, awayteam, ha, at)

                plist = st.range((8, 29), (8, 80)).value
                nlist = st.range((6, 29), (6, 80)).value
                pldic = {'ú': 'u', 'é': 'e', 'É': 'E'}

                def pl_2(l):
                    l2 = []
                    for ll in l:
                        ll2 = str(ll)
                        for p in pldic:
                            ll2 = ll2.replace(p, pldic[p])
                        l2.append(ll2)
                    return l2

                pl2 = pl_2(plist)
                result = st.range((row, 29), (row, 80)).value
                st.range((row, 29), (row, 80)).api.Font.Bold = False
                st.range((row, 29), (row, 80)).api.Font.ColorIndex = 1
                hal = x[at.index(ha) + 1]
                home = hal[0][0]
                away = hal[0][1]
                st.range(row, 5).value = home[0]
                st.range(row, 6).value = away[0]
                if len(home[1]) > 0:
                    try:
                        st.range(row, 5).api.AddComment('\n'.join(home[1]))
                    except:
                        st.range(row, 5).api.Comment.Text('\n'.join(home[1]))
                    st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                if len(away[1]) > 0:
                    try:
                        st.range(row, 6).api.AddComment('\n'.join(away[1]))
                    except:
                        st.range(row, 6).api.Comment.Text('\n'.join(away[1]))
                    st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                if 'red' in home[2]:
                    st.range(row, 5).api.Interior.Color = 255
                if 'red' in away[2]:
                    st.range(row, 6).api.Interior.Color = 255

                for name in hal[1]:
                    ind = 0
                    data = hal[1][name]
                    if name in plist:
                        ind = plist.index(name)
                    elif name in pl2:
                        ind = pl2.index(name)

                    if ind == 0:
                        print(ha, '-', name, 'error')
                    else:
                        result[ind] = ''.join(data[1])
                        col = ind + 29
                        if data[2][0] == 'Bold':
                            st.range(row, col).api.Font.Bold = True
                        if data[2][1] == 'red':
                            st.range(row, col).api.Font.ColorIndex = 3
                        if data[2][1] == 'blue':
                            st.range(row, col).api.Font.ColorIndex = 5
                        if data[2][2] == 'yellow':
                            st.range(row, col).api.Interior.Color = 65535
                        if data[2][2] == 'red':
                            st.range(row, col).api.Interior.Color = 255
                        if data[2][3] == 'I':
                            st.range(row, col).api.Font.Italic = True
                        if len(data[3]) > 0:
                            try:
                                st.range(row, col).api.AddComment('\n'.join(data[3]))
                            except:
                                st.range(row, col).api.Comment.Text('\n'.join(data[3]))
                            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                st.range((row, 29), (row, 80)).value = result
                og = hal[2]
                if og > 0:
                    st.range(row, plist.index('OG') + 29).value = og
                print('%s Finished' % ha)

            else:
                print(ha, ' not in PS')

    def draft_to_epl_summary(self, x):
        # path = r'E:\Company\League\EPL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Whoscored', 'Soccerway']].set_index('Whoscored').to_dict()['Soccerway']
        wsc_ps_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal_Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Man_City',
     'Manchester United': 'Man_Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West_Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wsc_game_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Manchester City',
     'Manchester United': 'Manchester Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wb = xw.Book('EPL_Playersheet_2018-2019.xlsm')
        hteam = x[0][0]
        ateam = x[0][1]
        at = [hteam, ateam]
        for ha in at:
            if ha in wsc_ps_dic:
                st = wb.sheets[wsc_ps_dic[ha]]

                date = datetime.datetime.strptime(x[0][3], '%d/%m/%Y %H:%M:%S').date()
                awayteam = at[1 - at.index(ha)]
                if awayteam in wsc_game_dic:
                    awayteam = wsc_game_dic[awayteam]
                game = st.range((10, 1), (80, 4)).value
                row = self.row_determine(date, game, awayteam, ha, at)

                plist = st.range((8, 29), (8, 80)).value
                nlist = st.range((6, 29), (6, 80)).value
                result = st.range((row, 29), (row, 80)).value
                st.range((row, 29), (row, 80)).api.Font.Bold = False
                st.range((row, 29), (row, 80)).api.Font.ColorIndex = 1
                hal = x[at.index(ha) + 1]
                home = hal[0][0]
                away = hal[0][1]
                st.range(row, 5).value = home[0]
                st.range(row, 6).value = away[0]
                if len(home[1]) > 0:
                    try:
                        st.range(row, 5).api.AddComment('\n'.join(home[1]))
                    except:
                        st.range(row, 5).api.Comment.Text('\n'.join(home[1]))
                    st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                if len(away[1]) > 0:
                    try:
                        st.range(row, 6).api.AddComment('\n'.join(away[1]))
                    except:
                        st.range(row, 6).api.Comment.Text('\n'.join(away[1]))
                    st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                if home[2] == ['red']:
                    st.range(row, 5).api.Interior.Color = 255
                if away[2] == ['red']:
                    st.range(row, 6).api.Interior.Color = 255

                for name in hal[1]:
                    ind = 0
                    data = hal[1][name]
                    if name in plist:
                        ind = plist.index(name)

                    if ind == 0:
                        print(ha, '-', name, 'error')
                    else:
                        result[ind] = ''.join(data[1])
                        col = ind + 29
                        if data[2][0] == 'Bold':
                            st.range(row, col).api.Font.Bold = True
                        if data[2][1] == 'red':
                            st.range(row, col).api.Font.ColorIndex = 3
                        if data[2][1] == 'blue':
                            st.range(row, col).api.Font.ColorIndex = 5
                        if data[2][2] == 'yellow':
                            st.range(row, col).api.Interior.Color = 65535
                        if data[2][2] == 'red':
                            st.range(row, col).api.Interior.Color = 255
                        if data[2][3] == 'I':
                            st.range(row, col).api.Font.Italic = True
                        if len(data[3]) > 0:
                            try:
                                st.range(row, col).api.AddComment('\n'.join(data[3]))
                            except:
                                st.range(row, col).api.Comment.Text('\n'.join(data[3]))
                            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                st.range((row, 29), (row, 80)).value = result
                og = hal[2]
                if og > 0:
                    st.range(row, plist.index('OG') + 29).value = og
                print('%s Finished' % ha)

            else:
                print(ha, ' not in PS')

    def draft_to_epl_cup(self, x):
        # path = r'E:\Company\League\EPL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Whoscored', 'Soccerway']].set_index('Whoscored').to_dict()['Soccerway']
        wsc_ps_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal_Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Man_City',
     'Manchester United': 'Man_Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West_Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wsc_game_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Manchester City',
     'Manchester United': 'Manchester Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wb = xw.Book('EPL_Playersheet_2018-2019.xlsm')
        hteam = x[0][0]
        ateam = x[0][1]
        at = [hteam, ateam]
        for ha in at:
            if ha in wsc_ps_dic:
                st = wb.sheets[wsc_ps_dic[ha]]

                date = datetime.datetime.strptime(x[0][3], '%d/%m/%Y %H:%M:%S').date()
                awayteam = at[1 - at.index(ha)]
                if awayteam in wsc_game_dic:
                    awayteam = wsc_game_dic[awayteam]
                game = st.range((10, 1), (80, 4)).value
                row = self.row_determine(date, game, awayteam, ha, at)
                hal = x[at.index(ha) + 1]

                self.to_sheet_all(st, row, hal, ha)

            else:
                print(ha, ' not in PS')

    def draft_to_unl_cup(self, x, unl):
        # path = r'E:\Company\League\EPL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Whoscored', 'Soccerway']].set_index('Whoscored').to_dict()['Soccerway']
        wsc_ps_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal_Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Man_City',
     'Manchester United': 'Man_Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West_Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        wsc_game_dic = {'Arsenal': 'Arsenal',
     'Bournemouth': 'Bournemouth',
     'Brighton': 'Brighton',
     'Burnley': 'Burnley',
     'Cardiff': 'Cardiff',
     'Chelsea': 'Chelsea',
     'Crystal Palace': 'Crystal Palace',
     'Everton': 'Everton',
     'Fulham': 'Fulham',
     'Huddersfield': 'Huddersfield',
     'Leicester': 'Leicester',
     'Liverpool': 'Liverpool',
     'Manchester City': 'Manchester City',
     'Manchester United': 'Manchester Utd',
     'Newcastle United': 'Newcastle',
     'Southampton': 'Southampton',
     'Tottenham': 'Tottenham',
     'Watford': 'Watford',
     'West Ham': 'West Ham',
     'Wolverhampton Wanderers': 'Wolves'}
        league = 'UNL' + unl
        wbdic = {'ISA': 'ISA Player Sheet  2018-19.xlsm', 'SFL': 'SFL_Player_Sheet_2018-19.xlsm',
                 'PFL': 'PFL_Player_Sheet_18-19.xlsm', 'EPL': 'EPL_Playersheet_2018-2019.xlsm',
                 'UNLA': 'UNLA_Playersheet_2018-2019.xlsm', 'UNLB': 'UNLB_Playersheet_2018-2019.xlsm',
                 'UNLC': 'UNLC_Playersheet_2018-2019.xlsm', 'UNLD': 'UNLD_Playersheet_2018-2019.xlsm'}
        wb = xw.Book(wbdic[league])
        snames = [sheet.name for sheet in wb.sheets]
        hteam = x[0][0]
        ateam = x[0][1]
        at = [hteam, ateam]
        for ha in at:
            if ha in snames:
                st = wb.sheets[ha]

                date = datetime.datetime.strptime(x[0][3], '%d/%m/%Y %H:%M:%S').date()
                awayteam = at[1 - at.index(ha)]
                game = st.range((10, 1), (80, 4)).value
                row = self.row_determine(date, game, awayteam, ha, at)
                hal = x[at.index(ha) + 1]

                self.to_sheet_nonumber(st, row, hal, ha)

            else:
                print(ha, ' not in PS')

    def draft_to_pfl_cup(self, x):
        # path = r'E:\Company\League\EPL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # psd = df[['Whoscored', 'Soccerway']].set_index('Whoscored').to_dict()['Soccerway']
        # psd = df[['Whoscored', 'PS']].set_index('Whoscored').to_dict()['PS']
        wsc_ps_dic = {'Aves': 'Desportivo_Aves',
     'Belenenses': 'Belenenses',
     'Benfica': 'Benfica',
     'Boavista': 'Boavista',
     'Braga': 'Braga',
     'Chaves': 'Chaves',
     'FC Porto': 'Porto',
     'Feirense': 'Feirense',
     'Maritimo': 'Marítimo',
     'Moreirense': 'Moreirense',
     'Nacional': 'Nacional',
     'Portimonense': 'Portimonense',
     'Rio Ave': 'Rio_Ave',
     'Santa Clara': 'Santa_Clara',
     'Sporting CP': 'Sporting_CP',
     'Tondela': 'Tondela',
     'Vitoria de Guimaraes': 'Vitória_Guimarães',
     'Vitoria de Setubal': 'Vitória_Setúbal'}
        wsc_game_dic = {'Aves': 'Desportivo Aves',
     'Belenenses': 'Belenenses',
     'Benfica': 'Benfica',
     'Boavista': 'Boavista',
     'Braga': 'Sporting Braga',
     'Chaves': 'Chaves',
     'FC Porto': 'Porto',
     'Feirense': 'Feirense',
     'Maritimo': 'Marítimo',
     'Moreirense': 'Moreirense',
     'Nacional': 'Nacional',
     'Portimonense': 'Portimonense',
     'Rio Ave': 'Rio Ave',
     'Santa Clara': 'Santa Clara',
     'Sporting CP': 'Sporting CP',
     'Tondela': 'Tondela',
     'Vitoria de Guimaraes': 'Vitória Guimarães',
     'Vitoria de Setubal': 'Vitória Setúbal'}
        wb = xw.Book('PFL_Player_Sheet_18-19.xlsm')
        hteam = x[0][0]
        ateam = x[0][1]
        at = [hteam, ateam]
        for ha in at:
            if ha in wsc_ps_dic:
                st = wb.sheets[wsc_ps_dic[ha]]

                date = datetime.datetime.strptime(x[0][3], '%d/%m/%Y %H:%M:%S').date()
                awayteam = at[1 - at.index(ha)]
                if awayteam in wsc_game_dic:
                    awayteam = wsc_game_dic[awayteam]
                game = st.range((10, 1), (80, 4)).value
                row = self.row_determine(date, game, awayteam, ha, at)

                hal = x[at.index(ha) + 1]

                self.to_sheet_pfl(st, row, hal, ha)

            else:
                print(ha, ' not in PS')

    def to_sheet_all(self, st, row, hal, ha):
        plist = st.range((8, 29), (8, 80)).value
        nlist = st.range((6, 29), (6, 80)).value
        result = st.range((row, 29), (row, 80)).value
        st.range((row, 29), (row, 80)).api.Font.Bold = False
        st.range((row, 29), (row, 80)).api.Font.ColorIndex = 1

        home = hal[0][0]
        away = hal[0][1]
        st.range(row, 5).value = home[0]
        st.range(row, 6).value = away[0]
        if len(home[1]) > 0:
            try:
                st.range(row, 5).api.AddComment('\n'.join(home[1]))
            except:
                st.range(row, 5).api.Comment.Text('\n'.join(home[1]))
            st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
            st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        if len(away[1]) > 0:
            try:
                st.range(row, 6).api.AddComment('\n'.join(away[1]))
            except:
                st.range(row, 6).api.Comment.Text('\n'.join(away[1]))
            st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
            st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        if home[2] == ['red']:
            st.range(row, 5).api.Interior.Color = 255
        if away[2] == ['red']:
            st.range(row, 6).api.Interior.Color = 255

        st.range((row, 19), (row, 25)).value = hal[1]

        for name in hal[2][0]:
            ind = 0
            data = hal[2][0][name]
            if name in plist:
                ind = plist.index(name)
            elif hal[2][1][name] in nlist:
                ind = nlist.index(hal[2][1][name])

            if ind == 0:
                print(ha, '-', name, 'error')
            else:
                result[ind] = ''.join(data[1])
                col = ind + 29
                if data[2][0] == 'Bold':
                    st.range(row, col).api.Font.Bold = True
                if data[2][1] == 'red':
                    st.range(row, col).api.Font.ColorIndex = 3
                if data[2][1] == 'blue':
                    st.range(row, col).api.Font.ColorIndex = 5
                if data[2][2] == 'yellow':
                    st.range(row, col).api.Interior.Color = 65535
                if data[2][2] == 'red':
                    st.range(row, col).api.Interior.Color = 255
                if data[2][3] == 'I':
                    st.range(row, col).api.Font.Italic = True
                if len(data[3]) > 0:
                    try:
                        st.range(row, col).api.AddComment('\n'.join(data[3]))
                    except:
                        st.range(row, col).api.Comment.Text('\n'.join(data[3]))
                    st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        st.range((row, 29), (row, 80)).value = result
        og = hal[3]
        if og > 0:
            st.range(row, plist.index('OG') + 29).value = og
        print('%s Finished' % ha)

    def to_sheet_nonumber(self, st, row, hal, ha):
        plist = st.range((8, 29), (8, 80)).value
        # nlist = st.range((6, 29), (6, 80)).value
        result = st.range((row, 29), (row, 80)).value
        st.range((row, 29), (row, 80)).api.Font.Bold = False
        st.range((row, 29), (row, 80)).api.Font.ColorIndex = 1

        home = hal[0][0]
        away = hal[0][1]
        st.range(row, 5).value = home[0]
        st.range(row, 6).value = away[0]
        if len(home[1]) > 0:
            try:
                st.range(row, 5).api.AddComment('\n'.join(home[1]))
            except:
                st.range(row, 5).api.Comment.Text('\n'.join(home[1]))
            st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
            st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        if len(away[1]) > 0:
            try:
                st.range(row, 6).api.AddComment('\n'.join(away[1]))
            except:
                st.range(row, 6).api.Comment.Text('\n'.join(away[1]))
            st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
            st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        if home[2] == ['red']:
            st.range(row, 5).api.Interior.Color = 255
        if away[2] == ['red']:
            st.range(row, 6).api.Interior.Color = 255

        st.range((row, 19), (row, 25)).value = hal[1]

        for name in hal[2][0]:
            ind = 0
            data = hal[2][0][name]
            if name in plist:
                ind = plist.index(name)
            # elif hal[2][1][name] in nlist:
            #     ind = nlist.index(hal[2][1][name])

            if ind == 0:
                print(ha, '-', name, 'error')
            else:
                result[ind] = ''.join(data[1])
                col = ind + 29
                if data[2][0] == 'Bold':
                    st.range(row, col).api.Font.Bold = True
                if data[2][1] == 'red':
                    st.range(row, col).api.Font.ColorIndex = 3
                if data[2][1] == 'blue':
                    st.range(row, col).api.Font.ColorIndex = 5
                if data[2][2] == 'yellow':
                    st.range(row, col).api.Interior.Color = 65535
                if data[2][2] == 'red':
                    st.range(row, col).api.Interior.Color = 255
                if data[2][3] == 'I':
                    st.range(row, col).api.Font.Italic = True
                if len(data[3]) > 0:
                    try:
                        st.range(row, col).api.AddComment('\n'.join(data[3]))
                    except:
                        st.range(row, col).api.Comment.Text('\n'.join(data[3]))
                    st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        st.range((row, 29), (row, 80)).value = result
        og = hal[3]
        if og > 0:
            st.range(row, plist.index('OG') + 29).value = og
        print('%s Finished' % ha)

    def to_sheet_pfl(self, st, row, hal, ha):
        plist = st.range((8, 29), (8, 80)).value
        nlist = st.range((6, 29), (6, 80)).value
        pldic = {'ú': 'u', 'é': 'e', 'É': 'E'}

        def pl_2(l):
            l2 = []
            for ll in l:
                ll2 = str(ll)
                for p in pldic:
                    ll2 = ll2.replace(p, pldic[p])
                l2.append(ll2)
            return l2

        pl2 = pl_2(plist)
        result = st.range((row, 29), (row, 80)).value
        st.range((row, 29), (row, 80)).api.Font.Bold = False
        st.range((row, 29), (row, 80)).api.Font.ColorIndex = 1
        home = hal[0][0]
        away = hal[0][1]
        st.range(row, 5).value = home[0]
        st.range(row, 6).value = away[0]
        if len(home[1]) > 0:
            try:
                st.range(row, 5).api.AddComment('\n'.join(home[1]))
            except:
                st.range(row, 5).api.Comment.Text('\n'.join(home[1]))
            st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
            st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        if len(away[1]) > 0:
            try:
                st.range(row, 6).api.AddComment('\n'.join(away[1]))
            except:
                st.range(row, 6).api.Comment.Text('\n'.join(away[1]))
            st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
            st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        if 'red' in home[2]:
            st.range(row, 5).api.Interior.Color = 255
        if 'red' in away[2]:
            st.range(row, 6).api.Interior.Color = 255

        st.range((row, 19), (row, 25)).value = hal[1]

        for name in hal[2][0]:
            ind = 0
            data = hal[2][0][name]
            if name in plist:
                ind = plist.index(name)
            elif name in pl2:
                ind = pl2.index(name)
            elif hal[2][1][name] in nlist:
                ind = nlist.index(hal[2][1][name])

            if ind == 0:
                print(ha, '-', name, 'error')
            else:
                result[ind] = ''.join(data[1])
                col = ind + 29
                if data[2][0] == 'Bold':
                    st.range(row, col).api.Font.Bold = True
                if data[2][1] == 'red':
                    st.range(row, col).api.Font.ColorIndex = 3
                if data[2][1] == 'blue':
                    st.range(row, col).api.Font.ColorIndex = 5
                if data[2][2] == 'yellow':
                    st.range(row, col).api.Interior.Color = 65535
                if data[2][2] == 'red':
                    st.range(row, col).api.Interior.Color = 255
                if data[2][3] == 'I':
                    st.range(row, col).api.Font.Italic = True
                if len(data[3]) > 0:
                    try:
                        st.range(row, col).api.AddComment('\n'.join(data[3]))
                    except:
                        st.range(row, col).api.Comment.Text('\n'.join(data[3]))
                    st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        st.range((row, 29), (row, 80)).value = result
        og = hal[3]
        if og > 0:
            st.range(row, plist.index('OG') + 29).value = og
        print('%s Finished' % ha)



