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

    def whsc_match(self, htmltext):
        tree = lxml.html.fromstring(htmltext)
        bs = BeautifulSoup(htmltext, 'html.parser')
        x = bs.find_all('script', type="text/javascript")
        yy = []
        for xxx in x:
            if len(xxx.text) > 5000:
                yy.append(xxx)
        y = yy[1].text
        z = y.split(';')[0][23:]
        u = z.replace('\'', '"').replace(',,', ',null,').replace(',,', ',null,').replace(',]', ',null]')
        uu = ']'.join(u.split(']')[:-2]) + ']'
        v = json.loads(uu)
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

    def df_to_draft2(self, x, tl):
        hteam = x[1]['hteam']
        ateam = x[1]['ateam']
        hscore = x[1]['score'].split(':')[0].strip()
        ascore = x[1]['score'].split(':')[1].strip()
        hco, aco, hcolor, acolor = [], [], [], []
        red = 10
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
                event = '%s men @ %s\' when %s-%s' % (red, time, score[0], score[1])
                r.append(event)
                c.append('red')
                ply_eve.append('red')
                ply_eve.append('1R @  %s\' when %s-%s' % (time, score[0], score[1]))
                if i[2] in pe:
                    pe[i[2]].append(ply_eve)
                else:
                    pe[i[2]] = [(ply_eve)]
                red -= 1
            elif i[4] == 'secondyellow':
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
        hog = allog[0]
        aog = allog[1]

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
            og = hal[3]
            if og > 0:
                st.range(row, plist.index('OG') + 29).value = og

    def draft_to_pfl(self, x):
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
                row = int(input('%s row:' % ha))
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
                row = int(input('%s row:' % ha))
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
                row = int(input('%s row:' % ha))
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
                og = hal[3]
                if og > 0:
                    st.range(row, plist.index('OG') + 29).value = og

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
                row = int(input('%s row:' % ha))
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

            else:
                print(ha, ' not in PS')


class Socc(object):
    def sow_matchs(self, league):
        # 网址
        headers = {
            'User-Agent': "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1"}
        all_url = league
        html = requests.get(all_url, headers=headers)
        # Soup = BeautifulSoup(html.text, 'lxml')
        tree = lxml.html.fromstring(html.text)
        AllGame = []
        # 判断是单场比赛还是比赛列表
        header = all_url.split('/')[3]
        if header == 'matches':
            try:
                AllGame.append(self.getsoccerway(all_url))
                print(self.getsoccerway(all_url)[0])
            except Exception as inst:
                print('error in ' + all_url)
                print(type(inst))
                print(inst.args)

        else:
            page = tree.xpath('/html/body ')
            games_a = tree.xpath(
                '//td[@class="score-time score"]/a ')
            games = [('https://int.soccerway.com/' + elem.get('href')) for elem in games_a]
            gamesss = [('No.' + str(games.index(elem) + 1) + elem[25:]) for elem in games]
            # print('\n'.join(gamesss))
            for x in gamesss:
                try:
                    y = x.split('/')
                    print('-----------------------------------------------------------------------')
                    print(
                        '|' + y[0] + '|' + '/'.join([y[3], y[4], y[5]]) + '| |' + y[7] + '| |' + y[8] + '| |' + y[9] + '|')
                except Exception as inst:
                    print('-----------------------------------------------------------------------')
                    print(x)
            print('-----------------------------------------------------------------------')
            start_num = input('Begin by Which Game(No.):')
            end_num = input('End by Which Game(No.):')

            if start_num == '':
                start_num = 1
            if end_num == '':
                end_num = len(games) + 2
            for match in games[int(start_num) - 1:int(end_num)]:
                try:
                    AllGame.append(self.getsoccerway(match))
                    number = games.index(match) + 1
                    print('No.' + str(number) + ' ' + self.getsoccerway(match)[0])
                except Exception as inst:
                    print('error in ' + match)
                    print(type(inst))
                    print(inst.args)

                    break

        return AllGame

    def list_to_elist(self, Game):
        x = Game
        elist = []

        for h in range(0, len(x)):
            match = [[], []]
            # 赛事信息
            info = [[], []]
            for i in [0, 1]:
                infoteam = []
                infoscore = []
                info[i].append(x[h][1][0][0][i])
                infoteam.append(x[h][1][0][0][i])
                infoteam.append(x[h][1][0][0][1 - i])
                info[i].append(infoteam)
                info[i].append(x[h][1][0][1])
                info[i].append(x[h][1][0][2])
                info[i].append(x[h][1][0][3])
                infoscore.append(str(x[h][1][0][4][0][i]) + '-' + str(x[h][1][0][4][0][1 - i]))
                infoscore.append(x[h][1][0][4][1][i])
                if x[h][1][0][4][2] is list:
                    infoscore.append(x[h][1][0][4][2][i])
                else:
                    infoscore.append(x[h][1][0][4][2])
                if x[h][1][0][4][3] is list:
                    infoscore.append(x[h][1][0][4][3][i])
                else:
                    infoscore.append(x[h][1][0][4][3])
                info[i].append(infoscore)

            # 增加比分和时间的索引
            score_time = []
            score_score = []
            score_timea = []
            for i in [0, 1]:
                for j in range(0, len(x[h][1][1][i])):
                    score_time.append(x[h][1][1][i][j][0])
                    score_score.append(x[h][1][1][i][j][3])
            for elem in score_time:
                if len(elem.split('+')) > 1:
                    score_time_trans = int(elem.split('\'')[0]) + (int(elem.split('\'')[-1].split('+')[-1]) / 100)
                else:
                    score_time_trans = int(elem.split('\'')[0])
                score_timea.append(score_time_trans)

            score_dic_l = []
            for elem in score_timea:
                score_dic_l.append([elem, score_score[score_timea.index(elem)]])
            score_dic_t = tuple(score_dic_l)
            score_dic = dict(score_dic_t)
            score_timea.sort()

            # 主客队比分
            score_all = [[], []]
            for i in [0, 1]:
                score = []
                score.append(x[h][1][0][4][0][i])
                # 标注
                score_c = []
                # 字体
                score_f = []
                # LG,OG,PG
                score_lop = []
                for j in range(0, len(x[h][1][1][i])):
                    # 进球之前的比分
                    score_b = []
                    score_b.insert(0, str(int(x[h][1][1][i][j][3].split('-')[i].strip()) - 1))
                    score_b.insert(1, x[h][1][1][i][j][3].split('-')[1 - i].strip())
                    score_b.insert(1, '-')
                    scorebefore = ''.join(score_b)
                    # 每个球判断是否LG,OG,PG
                    if int(x[h][1][1][i][j][0].split('\'')[0].split('+')[0]) < 85:
                        if x[h][1][1][i][j][2] == 'OG':
                            score_lop.append('OG @ ' + x[h][1][1][i][j][0] + ' when ' + scorebefore)
                        elif x[h][1][1][i][j][2] == 'PG':
                            score_lop.append('PG @ ' + x[h][1][1][i][j][0] + ' when ' + scorebefore)
                    elif int(x[h][1][1][i][j][0].split('\'')[0].split('+')[0]) <= 90:
                        if x[h][1][1][i][j][2] == 'OG':
                            score_lop.append('Late Goal & OG @ ' + x[h][1][1][i][j][0] + ' when ' + scorebefore)
                        elif x[h][1][1][i][j][2] == 'PG':
                            score_lop.append('Late Goal & PG @ ' + x[h][1][1][i][j][0] + ' when ' + scorebefore)
                        else:
                            score_lop.append('Late Goal @ ' + x[h][1][1][i][j][0] + ' when ' + scorebefore)

                # 红牌
                score_r = []
                if len(x[h][i + 2][1][4]) > 0:
                    score_f.append('Red')
                    for j in range(0, len(x[h][i + 2][1][4])):
                        if len(x[h][i + 2][1][4][j].split('+')) > 1:
                            red_time = int(x[h][i + 2][1][4][j].split('+')[0].split('\'')[0]) + (
                                        int(x[h][i + 2][1][4][j].split('+')[-1].split('\'')[0]) / 100)
                        else:
                            red_time = int(x[h][i + 2][1][4][j].split('\'')[0])
                        score_timea.append(red_time)
                        score_timea.sort()
                        if score_timea.index(red_time) > 0:
                            red_score = score_dic[score_timea[score_timea.index(red_time) - 1]]
                        else:
                            red_score = '0 - 0'
                        score_red = []
                        score_red.insert(0, red_score.split('-')[i].strip())
                        score_red.insert(1, red_score.split('-')[1 - i].strip())
                        score_red.insert(1, '-')
                        scorered = ''.join(score_red)
                        score_r.append(str(10 - j) + ' men @ ' + x[h][i + 2][1][4][j] + ' when ' + scorered)
                        score_timea.remove(red_time)

                else:
                    score_f.append('No Red')
                if len(x[h][i + 2][1][4]) > 0:
                    score_c.extend(score_r)

                # 失点
                score_m = []
                if len(x[h][i + 2][1][3]) > 0:
                    for j in range(0, len(x[h][i + 2][1][3])):
                        if len(x[h][i + 2][1][3][j].split('+')) > 1:
                            pm_time = int(x[h][i + 2][1][3][j].split('+')[0].split('\'')[0]) + (
                                        int(x[h][i + 2][1][3][j].split('+')[-1].split('\'')[0]) / 100)
                        else:
                            pm_time = int(x[h][i + 2][1][3][j].split('\'')[0])
                        score_timea.append(pm_time)
                        score_timea.sort()
                        if score_timea.index(pm_time) > 0:
                            pm_score = score_dic[score_timea[score_timea.index(pm_time) - 1]]
                        else:
                            pm_score = '0 - 0'
                        score_timea.remove(pm_time)
                        score_pm = []
                        score_pm.insert(0, pm_score.split('-')[i].strip())
                        score_pm.insert(1, pm_score.split('-')[1 - i].strip())
                        score_pm.insert(1, '-')
                        scorepm = ''.join(score_pm)
                        score_m.append('PK Missed @ ' + x[h][i + 2][1][3][j] + ' when ' + scorepm)
                    score_c.extend(score_m)

                score_c.extend(score_lop)

                # 加时、点球大战
                if type(x[h][1][0][4][3]) == list:
                    score_c.insert(0, 'Penalties ' + x[h][1][0][4][3][i])
                if type(x[h][1][0][4][2]) == list:
                    score_c.insert(0, 'Extra Time ' + x[h][1][0][4][2][i])

                score.append(score_f)
                score.append(score_c)

                score_all[i].append(score)

            # 主、客队球员
            player_all = [[], []]
            for i in [0, 1]:
                player_team = []
                # 换下球员
                subout = []
                earlysub = []
                for j in range(0, len(x[h][i + 2][3])):
                    if len(x[h][i + 2][3][j]) > 14:
                        if x[h][i + 2][3][j][14] != 'no early sub':
                            subout.append([x[h][i + 2][3][j][12], x[h][i + 2][3][j][13]])
                            earlysub.append([x[h][i + 2][3][j][12], x[h][i + 2][3][j][13], [x[h][i + 2][3][j][0]]])
                        else:
                            subout.append([x[h][i + 2][3][j][12], x[h][i + 2][3][j][13]])

                # 首发球员
                starter = []
                for j in range(0, 11):
                    player = []
                    player_x = ['X']
                    player_f = []
                    player_c = []
                    # 名字
                    player.append(x[h][i + 2][2][j][0])
                    # 换下
                    if len(x[h][i + 2][2][j][2]) > 0:
                        player_x.append('-')
                    # 红黄牌
                    if len(x[h][i + 2][2][j][5]) > 0:
                        player_f.append('R')
                        player_c.append('1R @ ' + x[h][i + 2][2][j][5][0])
                    elif len(x[h][i + 2][2][j][4]) > 0:
                        player_f.append('R')
                        player_c.append('2Y @ ' + x[h][i + 2][2][j][3][0] + ',' + x[h][i + 2][2][j][4][0])
                    elif len(x[h][i + 2][2][j][3]) > 0:
                        player_f.append('Y')
                    else:
                        player_f.append('')
                    # 进球、助攻
                    if x[h][i + 2][2][j][6] > 0:
                        player_f.append('B')
                        player_x.append(str(x[h][i + 2][2][j][6]))
                        if x[h][i + 2][2][j][8] > 0:
                            player_c.append(str(x[h][i + 2][2][j][8]) + 'A')
                    elif x[h][i + 2][2][j][8] > 0:
                        if x[h][i + 2][2][j][8] > 1:
                            player_x.append(str(x[h][i + 2][2][j][8]) + 'A')
                        else:
                            player_x.append('A')
                        player_f.append('')
                    else:
                        player_f.append('')

                    # 点球、乌龙、失点、加时
                    if len(x[h][i + 2][2][j][7]) > 0:
                        player_c.append('PG @ ' + ','.join(x[h][i + 2][2][j][7]))
                    if len(x[h][i + 2][2][j][10]) > 0:
                        player_c.append('OG @ ' + ','.join(x[h][i + 2][2][j][10]))
                    if len(x[h][i + 2][2][j][11]) > 0:
                        player_c.append('PK Missed @ ' + ','.join(x[h][i + 2][2][j][11]))
                    if len(x[h][i + 2][2][j][9]) > 0:
                        player_c.append('Extra time Goal @ ' + ','.join(x[h][i + 2][2][j][9]))
                    # earlysub
                    for elem in earlysub:
                        if x[h][i + 2][2][j][0] == elem[0][0]:
                            player_c.insert(0, 'Early Sub @ ' + elem[1])

                    player.append(player_x)
                    player.append(player_f)
                    player.append(player_c)
                    # 链接
                    player.append(x[h][i + 2][2][j][13])
                    # 号码
                    player.append(x[h][i + 2][2][j][1])
                    starter.append(player)
                player_team.append(starter)

                # 替补球员
                sub_bench = []

                for j in range(0, len(x[h][i + 2][3])):
                    player = []
                    player_x = []
                    player_f = []
                    player_c = []
                    # 名字
                    player.append(x[h][i + 2][3][j][0])
                    # 红黄牌
                    if len(x[h][i + 2][3][j][4]) > 0:
                        player_f.append('R')
                        player_c.append('1R @ ' + x[h][i + 2][3][j][4][0])
                    elif len(x[h][i + 2][3][j][3]) > 0:
                        player_f.append('R')
                        player_c.append('2Y @ ' + x[h][i + 2][3][j][2][0] + ',' + x[h][i + 2][3][j][3][0])
                    elif len(x[h][i + 2][3][j][2]) > 0:
                        player_f.append('Y')
                    else:
                        player_f.append('')
                    # 进球、助攻
                    if x[h][i + 2][3][j][5] > 0:
                        player_f.append('B')
                        player_x.append(str(x[h][i + 2][3][j][5]))
                        if x[h][i + 2][3][j][7] > 0:
                            player_c.append(str(x[h][i + 2][3][j][7]) + 'A')
                    elif x[h][i + 2][3][j][7] > 0:
                        if x[h][i + 2][3][j][7] > 1:
                            player_x.append(str(x[h][i + 2][3][j][7]) + 'A')
                        else:
                            player_x.append('A')
                        player_f.append('')
                    else:
                        player_f.append('')
                    # 点球、乌龙、失点、加时
                    if len(x[h][i + 2][3][j][6]) > 0:
                        player_c.append('PG @ ' + ','.join(x[h][i + 2][3][j][6]))
                    if len(x[h][i + 2][3][j][9]) > 0:
                        player_c.append('OG @ ' + ','.join(x[h][i + 2][3][j][9]))
                    if len(x[h][i + 2][3][j][10]) > 0:
                        player_c.append('PK Missed @ ' + ','.join(x[h][i + 2][3][j][10]))
                    if len(x[h][i + 2][3][j][8]) > 0:
                        player_c.append('Extra time Goal @ ' + ','.join(x[h][i + 2][3][j][8]))
                    # 换上、bench、sub out
                    if len(x[h][i + 2][3][j]) > 14:
                        playerlink = x[h][i + 2][3][j][15]
                        player_x.insert(0, 's')
                        player_f.append('')
                        # early sub on
                        for elem in earlysub:
                            if x[h][i + 2][3][j][0] == elem[2][0]:
                                player_x[0] = 'S'
                                player_f[2] = 'R'

                        # Sub Out
                        for elem in subout:
                            if x[h][i + 2][3][j][0] == elem[0][0]:
                                player_x[0] = 'S-'
                                player_f[2] = 'R'
                                player_c.insert(0, 'Sub out @ ' + elem[1])
                                for elem_2 in earlysub:
                                    if x[h][i + 2][3][j][0] == elem_2[0][0]:
                                        player_c[0] = 'Sub Out & Early Sub @ ' + elem_2[1]
                    else:
                        player_x.insert(0, 'b')
                        player_f.append('b')
                        playerlink = x[h][i + 2][3][j][13]

                    player.append(player_x)
                    player.append(player_f)
                    player.append(player_c)
                    player.append(playerlink)
                    player.append(x[h][i + 2][3][j][1])
                    sub_bench.append(player)
                player_team.append(sub_bench)

                # OG
                team_og = ['OG']
                og = 0
                if len(x[h][i + 2][1][2]) > 0:
                    for elem in x[h][i + 2][1][2]:
                        if elem != '':
                            og += 1
                team_og.append(og)
                player_team.append(team_og)

                player_all[i].append(player_team)

            # 数据分队
            for i in [0, 1]:
                info[i].append(['H', 'A'][i])
                match[i].append(info[i])
                mathchscore = []
                mathchscore.extend(score_all[i])
                mathchscore.extend(score_all[1 - i])
                match[i].append(mathchscore)
                match[i].extend(player_all[i])

            elist.extend(match)

        # 索引
        # elist
        # [i]match
        # i.0 info
        # i.0.0.队名'str'
        # i.0.1.队名[Self,Opponent]
        # i.0.2.日期[Date,Time]
        # i.0.3.赛事'str'
        # i.0.4.场地和观众[球场，城市，观众数]
        # i.0.5.比分：
        # i.0.5.0.Full time(X-X)没空格
        # i.0.5.1.Half time(X - X)有空格
        # i.0.5.2.Extra time(X - X)有空格/('No Extra Time')
        # i.0.5.3.Penalties(X - X)有空格/('No Penalties')
        # i.0.6. 主客场['H'/'A']
        # i.1.[0,1] score[Self,Opponent]
        # i.1.[0,1].0 score(number)
        # i.1.[0,1].1 Font(红底'Red'/'No Red')
        # i.1.[0,1].2 score comments[]
        # i.2 player
        # i.2.0 starter
        # i.2.0.j  (0-10)
        # i.2.0.j.0 name
        # i.2.0.j.1 首发标记['X'(,'-')]
        # --.2 font[]
        # --.2.0 红黄底('R'/'Y'/'')
        # --.2.1 粗体('B'/'')
        # --.3 comment[]
        # i.2.0.j.4 球员链接
        # i.2.1 sub-bench
        # i.2.1.k (0- )
        # i.2.1.k.0 name
        # i.2.1.k.1 替补标记[]
        # --.2 font
        # --.2.0 红黄底('R'/'Y'/'')
        # --.2.1 粗体('B'/'')
        # --.2.2 bench/early sub out('b'/'R')
        # --.3 comment[]
        # i.2.1.k.4 球员链接
        # i.2.2 OG
        # i.2.2.0 'OG'
        # i.2.2.1 OG number

        return elist

    def getsoccerway(self, url):
        # 网址
        headers = {
            'User-Agent': "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1"}
        all_url = url
        html = requests.get(all_url, headers=headers)
        # Soup = BeautifulSoup(html.text, 'lxml')
        tree = lxml.html.fromstring(html.text)
        # 写入excel
        # wb = xlwt.Workbook()
        # 建立多维列表
        Game = []
        TeamHome = []
        TeamAway = []
        GameEvent = []
        Title = []

        # 队名

        team = tree.xpath('//*[@id="subheading"]/h1/text()')[0]
        team_s = team.split('vs.')
        # sh = wb.add_sheet('H_' + team_s[0].strip()[0:4])
        # sa = wb.add_sheet('A_' + team_s[1].strip()[0:4])
        # se = wb.add_sheet('Event')
        # sh.write(1, 1, team_s[0].strip())
        # sh.write(1, 2, 'H')
        # sa.write(1, 1, team_s[1].strip())
        # sa.write(1, 2, 'A')
        TeamHome.insert(0, team_s[0].strip())
        TeamAway.insert(0, team_s[1].strip())
        Title.insert(0, [team_s[0].strip(), team_s[1].strip()])

        # 日期

        date_ta = tree.xpath('//*[@id="page_match_1_block_match_info_4"]/div[2]/div[2]/div[1]/dl/dt')
        date_da = tree.xpath('//*[@id="page_match_1_block_match_info_4"]/div[2]/div[2]/div[1]/dl/dd')
        date_tb = [elem.text_content() for elem in date_ta]
        date_db = [elem.text_content() for elem in date_da]

        date_s = date_db[date_tb.index('Date')]
        date_t = date_s.split(' ')
        Calendar = {"January": "01", "February": "02", "March": "03", "April": "04", "May": "05", "June": "06",
                    "July": "07", "August": "08", "September": "09", "October": "10", "November": "11",
                    "December": "12"}

        if 'Kick-off' in date_tb:
            datets = date_db[date_tb.index('Kick-off')]
            datet_s = datets.split('\n')[1].strip()
            dt_a = date_t[2] + '/' + Calendar[date_t[1]] + '/' + date_t[0] + '/' + datet_s.split(':')[0] + '/' + \
                   datet_s.split(':')[1]
            dt = datetime.datetime.strptime(dt_a, '%Y/%m/%d/%H/%M')
            dt_cn = dt + datetime.timedelta(hours=7)
            dt_b = dt_cn.strftime('%Y/%m/%d-%H:%M').split('-')
            date_a = dt_b[0]
            date_b = dt_b[1]
        else:
            date_a = date_t[2] + '/' + Calendar[date_t[1]] + '/' + date_t[0]
            date_b = "No Kick-off Time Information"

        Title.insert(1, [date_a, date_b])
        # sh.write(1, 0, date_a)
        # sh.write(2, 0, date_b)
        # sa.write(1, 0, date_a)
        # sa.write(2, 0, date_b)
        # 赛事，场地，观众
        # 赛事
        compe = ""
        if 'Competition' in date_tb:
            compe = date_db[date_tb.index('Competition')]
            # sh.write(3, 0, compe)
            # sa.write(3, 0, compe)

        Title.insert(2, compe)
        # 场地,观众

        VA_ta = tree.xpath('//*[@id="page_match_1_block_match_info_4"]/div[2]/div[2]/div[3]/dl/dt')
        VA_da = tree.xpath('//*[@id="page_match_1_block_match_info_4"]/div[2]/div[2]/div[3]/dl/dd')
        VA_tb = [elem.text_content() for elem in VA_ta]
        VA_db = [elem.text_content() for elem in VA_da]
        Venue_V = ''
        Venue_C = ''
        attd = ''
        if 'Venue' in VA_tb:
            Venue = VA_db[VA_tb.index('Venue')]
            Venue_VC = Venue.split('(')
            Venue_V = Venue_VC[0]
            Venue_VC.pop(0)
            Venue_C = ' '.join(' '.join(Venue_VC).split(")"))

            # sh.write(4, 0, Venue_V)
            # sh.write(5, 0, Venue_C)
            # sa.write(4, 0, Venue_V)
            # sa.write(5, 0, Venue_C)
        if 'Attendance' in VA_tb:
            attd = VA_db[VA_tb.index('Attendance')]
            # sh.write(6, 0, int(attd))
            # sa.write(6, 0, int(attd))

        Title.insert(3, [Venue_V, Venue_C, attd])

        # 比分
        GameScore = []
        score_ea = tree.xpath('//*[@id="page_match_1_block_match_info_4"]/div[2]/div[2]/div[2]/dl/dt')
        score_sa = tree.xpath('//*[@id="page_match_1_block_match_info_4"]/div[2]/div[2]/div[2]/dl/dd')
        score_e = [elem.text_content() for elem in score_ea]
        score_s = [elem.text_content() for elem in score_sa]
        if 'Full-time' in score_e:
            score = score_s[score_e.index('Full-time')]
            GameScore.append([int(score.split('-')[0].strip()), int(score.split('-')[1].strip())])
        else:
            GameScore.append("No Full Time")
        # sh.write(1, 3, int(score.split('-')[0].strip()))
        # sh.write(1, 4, int(score.split('-')[1].strip()))
        # sa.write(1, 3, int(score.split('-')[1].strip()))
        # sa.write(1, 4, int(score.split('-')[0].strip()))
        # sh.write(1, 6, score.strip())
        # sa.write(1, 6, score.split('-')[1].strip()+' - '+score.split('-')[0].strip())
        if 'Half-time' in score_e:
            scoreh = score_s[score_e.index('Half-time')]
            GameScore.append([scoreh.strip(), scoreh.split('-')[1].strip() + ' - ' + scoreh.split('-')[0].strip()])
        else:
            GameScore.append("No Half Time")
        # sh.write(1, 5, scoreh.strip())
        # sa.write(1, 5, scoreh.split('-')[1].strip()+' - '+scoreh.split('-')[0].strip())
        if 'Extra-time' in score_e:
            scoree = score_s[score_e.index('Extra-time')]
            GameScore.append([scoree.strip(), scoree.split('-')[1].strip() + ' - ' + scoree.split('-')[0].strip()])
        else:
            GameScore.append("No Extra Time")
        # sh.write(1, 7, scoree.strip())
        # sa.write(1, 7, scoree.split('-')[1].strip()+' - '+scoree.split('-')[0].strip())
        if 'Penalties' in score_e:
            scorep = score_s[score_e.index('Penalties')]
            GameScore.append([scorep.strip(), scorep.split('-')[1].strip() + ' - ' + scorep.split('-')[0].strip()])
        else:
            GameScore.append("No Penalties")
            # sh.write(1, 8, scorep.strip())
            # sa.write(1, 8, scorep.split('-')[1].strip()+' - '+scorep.split('-')[0].strip())

        Title.insert(4, GameScore)

        # 表头
        # sh.write(0,0,'Date')
        # sh.write(0,1,'Team')
        # sh.write(0,2,'H/A')
        # sh.write(0,3,'Score')
        # sh.write(0,4,'O-Score')
        # sh.write(0,5,'Half-time')
        # sh.write(0,6,'Full-time')
        # sh.write(0,7,'Extra-time')
        # sh.write(0,8,'Penalties')
        # sh.write(0,9,'RCs')

        # sh.write(3, 5, 'Goal')
        # sh.write(3, 6, 'PG')
        # sh.write(3, 7, 'Assist')
        # sh.write(3, 15, 'ET Goal')
        # sh.write(3, 8, 'YC')
        # sh.write(3, 9, 'Y2C')
        # sh.write(3, 10, 'RC')
        # sh.write(3, 11, 'OG')
        # sh.write(3, 12, 'PK Miss')
        # sh.write(3, 13, 'Early Sub')
        # sh.write(3, 14, 'Goal Time')

        # sa.write(0,0,'Date')
        # sa.write(0,1,'Team')
        # sa.write(0,2,'H/A')
        # sa.write(0,3,'Score')
        # sa.write(0,4,'O-Score')
        # sa.write(0,5,'Half-time')
        # sa.write(0,6,'Full-time')
        # sa.write(0,7,'Extra-time')
        # sa.write(0,8,'Penalties')
        # sa.write(0,9,'RCs')

        # sa.write(3, 5, 'Goal')
        # sa.write(3, 6, 'PG')
        # sa.write(3, 7, 'Assist')
        # sa.write(3, 15, 'ET Goal')
        # sa.write(3, 8, 'YC')
        # sa.write(3, 9, 'Y2C')
        # sa.write(3, 10, 'RC')
        # sa.write(3, 11, 'OG')
        # sa.write(3, 12, 'PK Miss')
        # sa.write(3, 13, 'Early Sub')
        # sa.write(3, 14, 'Goal Time')

        # se.write(2, 0, 'Time')
        # se.write(2, 1, 'Home')
        # se.write(2, 2, 'OG\PG')
        # se.write(2, 3, 'Assist')
        # se.write(2, 4, 'Score')
        # se.write(2, 5, 'Away')
        # se.write(2, 6, 'OG\PG')
        # se.write(2, 7, 'Assist')
        # se.write(0, 1, 'LG')
        # se.write(0, 0, 'PM')
        # se.write(0, 2, 'PG')
        # se.write(0, 3, 'OG')
        # se.write(0, 4, 'Events Num')
        # se.write(0, 5, 'LG')
        # se.write(0, 6, 'PG')
        # se.write(0, 7, 'OG')
        # se.write(0, 8, 'PM')
        PMH = []
        PMA = []
        # 事件
        GoalEvent = []
        HomeGoalAll = []
        AwayGoalAll = []
        eventid = '#page_match_1_block_match_goals_13'
        event = tree.cssselect('%s > div > table > tr' % eventid)
        if len(event) == 0:
            eventid = '#page_match_1_block_match_goals_12'
            event = tree.cssselect('%s > div > table > tr' % eventid)
        event_num = eventid[-2:]

        # se.write(1, 4, len(event))
        OGH = []
        PGH = []
        LGH = []
        OGA = []
        PGA = []
        LGA = []
        ASH = []
        ASA = []
        RCH = []
        RCA = []
        for eventn in range(1, len(event) + 1):
            HomeGoal = []
            AwayGoal = []
            # 时间
            event_time_a = tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[1]/div/span[1]' % (event_num, eventn))
            event_time_b = tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[3]/div/span[1]' % (event_num, eventn))
            event_time = [''.join([elem.text_content() for elem in event_time_a]),
                          ''.join([elem.text_content() for elem in event_time_b])]
            # se.write(eventn + 2, 0, ''.join(event_time))
            # 比分
            event_score = tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[2]/div' % (event_num, eventn))[
                0].text_content()
            # se.write(eventn + 2, 4, event_a)

            # 主队

            event_a = tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[1]/div/a' % (event_num, eventn))
            eventa = ''.join([elem.text_content() for elem in event_a])
            opg_a = ''.join(tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[1]/div/text()' % (
                event_num, eventn))).strip()
            opga = len(''.join(opg_a).split())
            assista = tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[1]/div/span[@class="assist"]/a/text()' % (
                event_num, eventn))
            # se.write(eventn + 2, 1, eventa)
            if eventa != '':
                HomeGoal.append(eventa)
                if opga > 0:
                    opga_a = opg_a.split('(')[-1].split(')')[0]
                    HomeGoal.append(opga_a)
                    # se.write(eventn + 2, 2, opga_a)
                    if opga_a == 'OG':
                        OGH.append(event_time)
                    elif opga_a == "PG":
                        PGH.append(event_time)
                elif len(assista) > 0:
                    assist_a = tree.xpath(
                        '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[1]/div/span[2]/a/text()[1]' % (
                        event_num, eventn))[
                        0]
                    HomeGoal.append(assist_a.strip())
                    # se.write(eventn + 2, 3, assist_a.strip())
                    ASH.append(assist_a.strip())
                else:
                    HomeGoal.append('')
                if len(''.join([elem.text_content() for elem in event_time_a]).split('\'')[0].split('+')[0]) > 0:
                    if int(''.join([elem.text_content() for elem in event_time_a]).split('\'')[0].split('+')[0]) >= 85:
                        if int(''.join([elem.text_content() for elem in event_time_a]).split('\'')[0].split('+')[0]) <= 90:
                            LGH.append(''.join([elem.text_content() for elem in event_time_a]))
                HomeGoal.insert(0, ''.join(event_time))
                HomeGoal.insert(3, event_score)
                HomeGoalAll.append(HomeGoal)

            # 客队

            event_a = tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[3]/div/a' % (event_num, eventn))
            eventa = ''.join([elem.text_content() for elem in event_a])
            opg_a = ''.join(tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[3]/div/text()' % (
                event_num, eventn))).strip()
            opga = len(''.join(opg_a).split())
            assista = tree.xpath(
                '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[3]/div/span[@class="assist"]/a/text()' % (
                event_num, eventn))
            if eventa != '':
                AwayGoal.insert(2, eventa)
                # se.write(eventn + 2, 5, eventa)
                if opga > 0:
                    opga_a = opg_a.split('(')[-1].split(')')[0]
                    AwayGoal.append(opga_a)
                    # se.write(eventn + 2, 6, opga_a)
                    if opga_a == 'OG':
                        OGA.append(event_time)
                    elif opga_a == "PG":
                        PGA.append(event_time)
                elif len(assista) > 0:
                    assist_a = tree.xpath(
                        '//*[@id="page_match_1_block_match_goals_%s"]/div/table/tr[%s]/td[3]/div/span[2]/a/text()[1]' % (
                        event_num, eventn))[
                        0]
                    AwayGoal.append(assist_a.strip())
                    # se.write(eventn + 2, 7, assist_a.strip())
                    ASA.append(assist_a.strip())
                else:
                    AwayGoal.append('')
                if len(''.join([elem.text_content() for elem in event_time_b]).split('\'')[0].split('+')[0]) > 0:
                    if int(''.join([elem.text_content() for elem in event_time_b]).split('\'')[0].split('+')[0]) >= 85:
                        if int(''.join([elem.text_content() for elem in event_time_b]).split('\'')[0].split('+')[0]) <= 90:
                            LGA.append(''.join([elem.text_content() for elem in event_time_b]))
                AwayGoal.insert(0, ''.join(event_time))
                AwayGoal.insert(3, event_score)
                AwayGoalAll.append(AwayGoal)
        GoalEvent.append(HomeGoalAll)
        GoalEvent.append(AwayGoalAll)
        # se.write(1, 1, ' '.join(LGH))
        # se.write(1, 2, ' '.join(PGH))
        # se.write(1, 3, ' '.join(OGH))
        # se.write(1, 5, ' '.join(LGA))
        # se.write(1, 6, ' '.join(PGA))
        # se.write(1, 7, ' '.join(OGA))

        # 主队首发球员
        PlayerHome = []
        for teamn in range(1, 12):
            PlayerHomei = []
            playerh_a = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="player large-link"]/a' % teamn)
            playerh = [elem.text_content().strip() for elem in playerh_a][0]
            numh_a = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="shirtnumber"]' % teamn)
            numh = ''.join([elem.text_content().strip() for elem in numh_a])
            subed = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="player large-link"]/img' % teamn)
            span_t = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="bookings"]/span/img' % teamn)
            span_time = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="bookings"]/span' % teamn)
            if len(numh_a) != 0:
                PlayerHomei.append(int(numh))
            else:
                PlayerHomei.append("no number")
            # sh.write(teamn + 3, 1, int(numh))
            PlayerHomei.insert(0, playerh.strip())
            # sh.write(teamn + 3, 2, playerh.strip())
            subed_a = [elem.get('title') for elem in subed]
            PlayerHomei.append(subed_a)
            # sh.write(teamn + 3, 3, subed_a)
            span_t_all = [elem.get('src').split('/')[-1].split('.')[0] for elem in span_t]
            span_time_a = [elem.text_content().strip() for elem in span_time]
            goal = 0
            goal_t = []
            ET_goal = []
            pg = []
            og = []
            pm = []
            # 判断事件
            YCi = []
            Y2Ci = []
            RCi = []
            for j in range(0, len(span_t_all)):
                if span_t_all[j] == 'G' or span_t_all[j] == 'PG':
                    goal += 1
                    goal_t.append(span_time_a[j])
                    if int(span_time_a[j].split('\'')[0].split('+')[0]) > 90:
                        ET_goal.append(span_time_a[j])
                        goal -= 1
                if span_t_all[j] == 'PG':
                    pg.append(span_time_a[j])
                if span_t_all[j] == 'YC':
                    YCi.append(span_time_a[j])
                    # sh.write(teamn + 3, 8, span_time_a[j])
                if span_t_all[j] == 'Y2C':
                    Y2Ci.append(span_time_a[j])
                    RCH.append(span_time_a[j])
                    # sh.write(teamn + 3, 9, span_time_a[j])
                if span_t_all[j] == 'RC':
                    RCH.append(span_time_a[j])
                    RCi.append(span_time_a[j])
                    # sh.write(teamn + 3, 10, span_time_a[j])
                if span_t_all[j] == 'OG':
                    og.append(span_time_a[j])
                if span_t_all[j] == 'PM':
                    pm.append(span_time_a[j])
                    PMH.append(span_time_a[j])
            PlayerHomei.append(YCi)
            PlayerHomei.append(Y2Ci)
            PlayerHomei.append(RCi)
            assis = 0
            for elem in ASH:
                if elem == playerh.strip():
                    assis += 1

            # 球员链接
            playerh_link = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="player large-link"]/a' % teamn)
            playerlink = [elem.get('href') for elem in playerh_link]
            # if goal > 0:
            # sh.write(teamn + 3, 5, goal)
            PlayerHomei.append(goal)
            PlayerHomei.append(pg)
            # sh.write(teamn + 3, 6, pg)
            # if assis > 0:
            # sh.write(teamn + 3, 7, assis)
            PlayerHomei.append(assis)
            PlayerHomei.append(ET_goal)
            PlayerHomei.append(og)
            PlayerHomei.append(pm)
            PlayerHomei.append(goal_t)
            PlayerHomei.append(playerlink)
            # sh.write(teamn + 3, 15, ET_goal)
            # sh.write(teamn + 3, 11, og)
            # sh.write(teamn + 3, 12, pm)
            # sh.write(teamn + 3, 14, ','.join(goal_t))
            PlayerHome.append(PlayerHomei)

        # 客队首发球员
        PlayerAway = []
        for teamn in range(1, 12):
            PlayerAwayi = []
            playerh_a = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="player large-link"]/a' % teamn)
            playerh = [elem.text_content().strip() for elem in playerh_a][0]
            numh_a = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="shirtnumber"]' % teamn)
            numh = ''.join([elem.text_content().strip() for elem in numh_a])
            subed = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="player large-link"]/img' % teamn)
            span_t = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="bookings"]/span/img' % teamn)
            span_time = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="bookings"]/span' % teamn)
            if len(numh_a) != 0:
                PlayerAwayi.append(int(numh))
            else:
                PlayerAwayi.append('no number')
            # sa.write(teamn + 3, 1, int(numh))
            # sa.write(teamn + 3, 2, playerh.strip())
            PlayerAwayi.insert(0, playerh.strip())
            subed_a = [elem.get('title') for elem in subed]
            PlayerAwayi.append(subed_a)
            # sa.write(teamn + 3, 3, subed_a)
            span_t_all = [elem.get('src').split('/')[-1].split('.')[0] for elem in span_t]
            span_time_a = [elem.text_content().strip() for elem in span_time]
            goal = 0
            goal_t = []
            ET_goal = []
            pg = []
            og = []
            pm = []
            # 判断事件
            YCi = []
            Y2Ci = []
            RCi = []
            for j in range(0, len(span_t_all)):
                if span_t_all[j] == 'G' or span_t_all[j] == 'PG':
                    goal += 1
                    goal_t.append(span_time_a[j])
                    if int(span_time_a[j].split('\'')[0].split('+')[0]) > 90:
                        ET_goal.append(span_time_a[j])
                        goal -= 1
                if span_t_all[j] == 'PG':
                    pg.append(span_time_a[j])
                if span_t_all[j] == 'YC':
                    YCi.append(span_time_a[j])
                    # sa.write(teamn + 3, 8, span_time_a[j])
                if span_t_all[j] == 'Y2C':
                    # sa.write(teamn + 3, 9, span_time_a[j])
                    RCA.append(span_time_a[j])
                    Y2Ci.append(span_time_a[j])
                if span_t_all[j] == 'RC':
                    # sa.write(teamn + 3, 10, span_time_a[j])
                    RCA.append(span_time_a[j])
                    RCi.append(span_time_a[j])
                if span_t_all[j] == 'OG':
                    og.append(span_time_a[j])
                if span_t_all[j] == 'PM':
                    pm.append(span_time_a[j])
                    PMA.append(span_time_a[j])
            PlayerAwayi.append(YCi)
            PlayerAwayi.append(Y2Ci)
            PlayerAwayi.append(RCi)
            assis = 0
            for elem in ASA:
                if elem == playerh.strip():
                    assis += 1

            # 球员链接
            playerh_link = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups table"]/tbody/tr[%s]/td[@class="player large-link"]/a' % teamn)
            playerlink = [elem.get('href') for elem in playerh_link]

            PlayerAwayi.append(goal)
            PlayerAwayi.append(pg)
            PlayerAwayi.append(assis)
            PlayerAwayi.append(ET_goal)
            PlayerAwayi.append(og)
            PlayerAwayi.append(pm)
            PlayerAwayi.append(goal_t)
            PlayerAwayi.append(playerlink)
            # if goal > 0:
            # sa.write(teamn + 3, 5, goal)
            # sa.write(teamn + 3, 6, pg)
            # if assis > 0:
            # sa.write(teamn + 3, 7, assis)
            # sa.write(teamn + 3, 15, ET_goal)
            # sa.write(teamn + 3, 11, og)
            # sa.write(teamn + 3, 12, pm)
            # sa.write(teamn + 3, 14, ','.join(goal_t))
            PlayerAway.append(PlayerAwayi)

        # 主队替补
        SubHome = []
        subh_l = tree.xpath(
            '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups substitutions table"]/tbody/tr/td[@class="player large-link"]')
        # sh.write(17, 0, 'Sub Num')
        # sh.write(18, 0, len(subh_l))
        for subnh in range(1, len(subh_l) + 1):
            SubHomei = []
            # 球员
            playerh_s_a = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="player large-link"]/p[1]/a' % subnh)
            playerh_s = [elem.text_content().strip() for elem in playerh_s_a][0]
            numh_s_a = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="shirtnumber"]' % subnh)
            numh_s = ''.join([elem.text_content().strip() for elem in numh_s_a])
            if len(numh_s_a) != 0:
                SubHomei.append(int(numh_s))
            else:
                SubHomei.append('no number')
            # sh.write(subnh + 17, 1, int(numh_s))
            # sh.write(subnh + 17, 2, playerh_s.strip())
            SubHomei.insert(0, playerh_s.strip())
            # 事件
            span_t = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="bookings"]/span/img' % subnh)
            span_time = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="bookings"]/span' % subnh)
            span_t_all = [elem.get('src').split('/')[-1].split('.')[0] for elem in span_t]
            span_time_a = [elem.text_content().strip() for elem in span_time]
            goal = 0
            goal_t = []
            ET_goal = []
            pg = []
            og = []
            pm = []
            # 判断事件
            YCi = []
            Y2Ci = []
            RCi = []
            for j in range(0, len(span_t_all)):
                if span_t_all[j] == 'G' or span_t_all[j] == 'PG':
                    goal += 1
                    goal_t.append(span_time_a[j])
                    if int(span_time_a[j].split('\'')[0].split('+')[0]) > 90:
                        ET_goal.append(span_time_a[j])
                        goal -= 1
                if span_t_all[j] == 'PG':
                    pg.append(span_time_a[j])
                if span_t_all[j] == 'YC':
                    YCi.append(span_time_a[j])
                    # sh.write(subnh + 17, 8, span_time_a[j])
                if span_t_all[j] == 'Y2C':
                    # sh.write(subnh + 17, 9, span_time_a[j])
                    RCH.append(span_time_a[j])
                    Y2Ci.append(span_time_a[j])
                if span_t_all[j] == 'RC':
                    # sh.write(subnh + 17, 10, span_time_a[j])
                    RCH.append(span_time_a[j])
                    RCi.append(span_time_a[j])
                if span_t_all[j] == 'OG':
                    og.append(span_time_a[j])
                if span_t_all[j] == 'PM':
                    pm.append(span_time_a[j])
                    PMH.append(span_time_a[j])
            SubHomei.append(YCi)
            SubHomei.append(Y2Ci)
            SubHomei.append(RCi)
            assis = 0
            for elem in ASH:
                if elem == playerh_s.strip():
                    assis += 1

            # if goal > 0:
            # sh.write(subnh + 17, 5, goal)
            # sh.write(subnh + 17, 6, pg)
            # if assis > 0:
            # sh.write(subnh + 17, 7, assis)
            # sh.write(subnh + 17, 15, ET_goal)
            # sh.write(subnh + 17, 11, og)
            # sh.write(subnh + 17, 12, pm)
            # sh.write(subnh + 17, 14, ','.join(goal_t))
            SubHomei.append(goal)
            SubHomei.append(pg)
            SubHomei.append(assis)
            SubHomei.append(ET_goal)
            SubHomei.append(og)
            SubHomei.append(pm)
            SubHomei.append(goal_t)

            # 换人
            subout = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="player large-link"]/p[2]/a' % subnh)
            subout_a = [elem.text_content() for elem in subout]
            # sh.write(subnh + 17, 3, subout_a)
            SubHomei.append(subout_a)
            sub_t = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="player large-link"]/p[2]' % subnh)
            for elem in sub_t:
                sub_t_a = elem.text_content().split(' ')
                # sh.write(subnh + 17, 4, sub_t_a[-1])
                SubHomei.append(sub_t_a[-1])
                if int(sub_t_a[-1].split('\'')[0].split('+')[0]) <= 45:
                    SubHomei.append(sub_t_a[-1])
                else:
                    SubHomei.append('no early sub')
                # sh.write(subnh + 17, 13, sub_t_a[-1])

            # 球员链接
            playerh_link = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container left"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="player large-link"]/p[1]/a' % subnh)
            playerlink = [elem.get('href') for elem in playerh_link]
            SubHomei.append(playerlink)

            SubHome.append(SubHomei)

        # 客队替补
        SubAway = []
        subh_l = tree.xpath(
            '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups substitutions table"]/tbody/tr/td[@class="player large-link"]')
        # sa.write(17, 0, 'Sub Num')
        # sa.write(18, 0, len(subh_l))
        for subnh in range(1, len(subh_l) + 1):
            SubAwayi = []
            # 球员
            playerh_s_a = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="player large-link"]/p[1]/a' % subnh)
            playerh_s = [elem.text_content().strip() for elem in playerh_s_a][0]
            numh_s_a = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="shirtnumber"]' % subnh)
            numh_s = ''.join([elem.text_content().strip() for elem in numh_s_a])
            if len(numh_s_a) != 0:
                SubAwayi.append(int(numh_s))
            else:
                SubAwayi.append('no number')
            # sa.write(subnh + 17, 1, int(numh_s))
            # sa.write(subnh + 17, 2, playerh_s.strip())
            SubAwayi.insert(0, playerh_s.strip())
            # 事件
            span_t = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="bookings"]/span/img' % subnh)
            span_time = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="bookings"]/span' % subnh)
            span_t_all = [elem.get('src').split('/')[-1].split('.')[0] for elem in span_t]
            span_time_a = [elem.text_content().strip() for elem in span_time]
            goal = 0
            goal_t = []
            ET_goal = []
            pg = []
            og = []
            pm = []
            # 判断事件
            YCi = []
            Y2Ci = []
            RCi = []
            for j in range(0, len(span_t_all)):
                if span_t_all[j] == 'G' or span_t_all[j] == 'PG':
                    goal += 1
                    goal_t.append(span_time_a[j])
                    if int(span_time_a[j].split('\'')[0][:1]) > 90:
                        ET_goal.append(span_time_a[j])
                        goal -= 1
                if span_t_all[j] == 'PG':
                    pg.append(span_time_a[j])
                if span_t_all[j] == 'YC':
                    # sa.write(subnh + 17, 8, span_time_a[j])
                    YCi.append(span_time_a[j])
                if span_t_all[j] == 'Y2C':
                    # sa.write(subnh + 17, 9, span_time_a[j])
                    RCA.append(span_time_a[j])
                    Y2Ci.append(span_time_a[j])
                if span_t_all[j] == 'RC':
                    # sa.write(subnh + 17, 10, span_time_a[j])
                    RCA.append(span_time_a[j])
                    RCi.append(span_time_a[j])
                if span_t_all[j] == 'OG':
                    og.append(span_time_a[j])
                if span_t_all[j] == 'PM':
                    pm.append(span_time_a[j])
                    PMA.append(span_time_a[j])
            SubAwayi.append(YCi)
            SubAwayi.append(Y2Ci)
            SubAwayi.append(RCi)
            assis = 0
            for elem in ASA:
                if elem == playerh_s.strip():
                    assis += 1
            # if goal > 0:
            # sa.write(subnh + 17, 5, goal)
            # sa.write(subnh + 17, 6, pg)
            # if assis > 0:
            # sa.write(subnh + 17, 7, assis)
            # sa.write(subnh + 17, 15, ET_goal)
            # sa.write(subnh + 17, 11, og)
            # sa.write(subnh + 17, 12, pm)
            # sa.write(subnh + 17, 14, ','.join(goal_t))
            SubAwayi.append(goal)
            SubAwayi.append(pg)
            SubAwayi.append(assis)
            SubAwayi.append(ET_goal)
            SubAwayi.append(og)
            SubAwayi.append(pm)
            SubAwayi.append(goal_t)

            # 换人
            subout = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="player large-link"]/p[2]/a' % subnh)
            subout_a = [elem.text_content() for elem in subout]
            # sa.write(subnh + 17, 3, subout_a)
            SubAwayi.append(subout_a)
            sub_t = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="player large-link"]/p[2]' % subnh)
            for elem in sub_t:
                sub_t_a = elem.text_content().split(' ')
                SubAwayi.append(sub_t_a[-1])
                # sa.write(subnh + 17, 4, sub_t_a[-1])
                if int(sub_t_a[-1].split('\'')[0].split('+')[0]) <= 45:
                    SubAwayi.append(sub_t_a[-1])
                else:
                    SubAwayi.append('no early sub')
                # sa.write(subnh + 17, 13, sub_t_a[-1])

            # 球员链接
            playerh_link = tree.xpath(
                '//*[@id="yui-main"]/div/div/div[2]/div[@class="combined-lineups-container"]/div[@class="container right"]/table[@class="playerstats lineups substitutions table"]/tbody/tr[%s]/td[@class="player large-link"]/p[1]/a' % subnh)
            playerlink = [elem.get('href') for elem in playerh_link]
            SubAwayi.append(playerlink)

            SubAway.append(SubAwayi)

        # se.write(1, 0, ' '.join(PMH))
        # se.write(1, 8, ' '.join(PMA))
        RCH.sort()
        RCA.sort()
        TeamHome.append([LGH, PGH, OGH, PMH, RCH])
        TeamAway.append([LGA, PGA, OGA, PMA, RCA])
        # sh.write(1,9,len(RCH))
        # sa.write(1,9,len(RCA))
        # for i in range(0,len(RCH)):
        # sh.write(2,9+i,RCH[i])
        # for i in range(0,len(RCA)):
        # sa.write(2,9+i,RCA[i])

        # 存储
        # wb.save(''.join(date_a.split('/'))+(' ')+team_s[0].strip()+' VS '+team_s[1].strip()+'.xls',)
        # print('Game:', team_s[0].strip(), ' VS ', team_s[1].strip(), ' @', date_a, ' saved')
        Game.append(str('Game:' + team_s[0].strip() + ' VS ' + team_s[1].strip() + ' @' + date_a + ' saved'))
        GameEvent.append(Title)
        GameEvent.append(GoalEvent)
        TeamHome.append(PlayerHome)
        TeamHome.append(SubHome)
        TeamAway.append(PlayerAway)
        TeamAway.append(SubAway)
        Game.append(GameEvent)
        Game.append(TeamHome)
        Game.append(TeamAway)

        # 索引
        # Game：
        # 0.'Game...saverd'
        # 1.GameEvent:
        # 1.0.Title:
        # 1.0.0.队名[H,A]
        # 1.0.1.日期[Date,Time]
        # 1.0.2.赛事
        # 1.0.3.场地和观众
        # 1.0.4.比分：
        # 1.0.4.0.Full time[H,A]
        # 1.0.4.1.Half time[H,A]
        # 1.0.4.2.Extra time[H,A]
        # 1.0.4.3.Penalties[H,A]
        # 1.1.GoalEvent[H,A],
        # 1.1.(0,1).i.Each Goal of Each Team:
        # 1.1.(0,1).i.0.Goal time
        # 1.1.(0,1).i.1.Scorer
        # 1.1.(0,1).i.2. OG/PG/Assist player/''
        # 1.1.(0,1).i.3. score after goal
        # 2.Team Home:
        # 3.Team Away:
        # (2,3).0. 队名
        # (2,3).1. [LG,PG,OG,PM,RC]
        # (2,3).2.首发
        # (2,3).2.i.0.名字
        # (2,3).2.i.1.号码
        # .2.换下[]
        # .3.黄牌[]
        # .4.两黄变一红[]
        # .5.直接红牌[]
        # .6.Goal
        # .7.PG[]
        # .8.助攻
        # .9.加时进球[]
        # .10.OG[]
        # .11.踢丢点球[]
        # .12.Goal time[]
        # .13 球员链接
        # (2,3).3.替补
        # (2,3).3.i.0.名字
        # (2,3).3.i.1.号码
        # .2.黄牌[]
        # .3.两黄变一红[]
        # .4.直接红牌[]
        # .5.Goal
        # .6.PG[]
        # .7.助攻
        # .8.加时进球[]
        # .9.OG[]
        # .10.踢丢点球[]
        # .11.Goal time[]
        # .12 换下的球员（‘’）
        # .13 换人时间()
        # .14 Early Sub()
        # .15球员链接(b = #13)

        return Game

    def elist_to_epl(self, elist):
        soc_ps_dic = {'AFC Bournemouth': 'Bournemouth',
                      'Arsenal': 'Arsenal',
                      'Brighton & Hove Albion': 'Brighton',
                      'Burnley': 'Burnley',
                      'Cardiff City': 'Cardiff',
                      'Chelsea': 'Chelsea',
                      'Crystal Palace': 'Crystal_Palace',
                      'Everton': 'Everton',
                      'Fulham': 'Fulham',
                      'Huddersfield Town': 'Huddersfield',
                      'Leicester City': 'Leicester',
                      'Liverpool': 'Liverpool',
                      'Manchester City': 'Man_City',
                      'Manchester United': 'Man_Utd',
                      'Newcastle United': 'Newcastle',
                      'Southampton': 'Southampton',
                      'Tottenham Hotspur': 'Tottenham',
                      'Watford': 'Watford',
                      'West Ham United': 'West_Ham',
                      'Wolverhampton Wanderers': 'Wolves'}
        soc_game_dic = {'AFC Bournemouth': 'Bournemouth',
                        'Arsenal': 'Arsenal',
                        'Brighton & Hove Albion': 'Brighton',
                        'Burnley': 'Burnley',
                        'Cardiff City': 'Cardiff',
                        'Chelsea': 'Chelsea',
                        'Crystal Palace': 'Crystal Palace',
                        'Everton': 'Everton',
                        'Fulham': 'Fulham',
                        'Huddersfield Town': 'Huddersfield',
                        'Leicester City': 'Leicester',
                        'Liverpool': 'Liverpool',
                        'Manchester City': 'Manchester City',
                        'Manchester United': 'Manchester Utd',
                        'Newcastle United': 'Newcastle',
                        'Southampton': 'Southampton',
                        'Tottenham Hotspur': 'Tottenham',
                        'Watford': 'Watford',
                        'West Ham United': 'West Ham',
                        'Wolverhampton Wanderers': 'Wolves'}
        wb = xw.Book('EPL_Playersheet_2018-2019.xlsm')
        for t in elist:
            info = t[0]
            self = info[0]
            if self in soc_ps_dic:
                st = wb.sheets[soc_ps_dic[self]]
                ha = ['H', 'A'][elist.index(t) % 2]
                ang = info[1][1]
                if ang in soc_game_dic:
                    ang = soc_game_dic[ang]

                game = st.range((10, 1), (80, 4)).value
                row = -1
                for g in game:
                    if g[2] == ha:
                        if g[1] == ang:
                            row = game.index(g) + 10
                if row == -1:
                    row = int(input('%s row:' % self))

                plist = st.range((8, 29), (8, 80)).value

                def pp2(x):
                    x2 = []
                    for y in x:
                        if y is not None:
                            if len(y.split(' ')) > 1:
                                z = ' '.join([y.split(' ')[0][0] + '.'] + y.split(' ')[1:])
                            else:
                                z = y
                        else:
                            z = y
                        x2.append(z)
                    return x2

                plist2 = pp2(plist)
                nlist = st.range((6, 29), (6, 80)).value
                result = st.range((row, 29), (row, 80)).value
                st.range((row, 29), (row, 80)).api.Font.Bold = False
                st.range((row, 29), (row, 80)).api.Font.ColorIndex = 1

                home = t[1][0]
                away = t[1][1]
                st.range(row, 5).value = home[0]
                st.range(row, 6).value = away[0]
                if len(home[2]) > 0:
                    try:
                        st.range(row, 5).api.AddComment('\n'.join(home[2]))
                    except:
                        st.range(row, 5).api.Comment.Text('\n'.join(home[2]))
                    st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, 5).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                if len(away[2]) > 0:
                    try:
                        st.range(row, 6).api.AddComment('\n'.join(away[2]))
                    except:
                        st.range(row, 6).api.Comment.Text('\n'.join(away[2]))
                    st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                    st.range(row, 6).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                if home[1][0] == ['Red']:
                    st.range(row, 5).api.Interior.Color = 255
                if away[1][0] == ['Red']:
                    st.range(row, 6).api.Interior.Color = 255

                hal = t[2]

                for player in hal[0]:
                    ind = 0
                    name = player[0]
                    number = player[5]
                    if name in plist:
                        ind = plist.index(name)
                    elif name in plist2:
                        ind = plist2.index(name)
                    elif number in nlist:
                        ind = nlist.index(number)
                    elif str(number) in nlist:
                        ind = nlist.index(str(number))

                    if ind == 0:
                        print(ha, '-', name, 'error')
                    else:
                        result[ind] = ''.join(player[1])
                        col = ind + 29
                        if player[2][1] == 'B':
                            st.range(row, col).api.Font.Bold = True
                        else:
                            st.range(row, col).api.Font.Bold = False
                        if player[2][0] == 'Y':
                            st.range(row, col).api.Interior.Color = 65535
                        elif player[2][0] == 'R':
                            st.range(row, col).api.Interior.Color = 255
                        else:
                            st.range(row, col).api.Interior.ColorIndex = 0
                        if len(player[3]) > 0:
                            try:
                                st.range(row, col).api.AddComment('\n'.join(player[3]))
                            except:
                                st.range(row, col).api.Comment.Text('\n'.join(player[3]))
                            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'

                for sub in hal[1]:
                    ind = 0
                    name = sub[0]
                    number = sub[5]
                    if name in plist:
                        ind = plist.index(name)
                    elif name in plist2:
                        ind = plist2.index(name)
                    elif number in nlist:
                        ind = nlist.index(number)
                    elif str(number) in nlist:
                        ind = nlist.index(str(number))

                    if ind == 0:
                        print(ha, '-', name, 'error')
                    else:
                        result[ind] = ''.join(sub[1])
                        col = ind + 29
                        if sub[2][1] == 'B':
                            st.range(row, col).api.Font.Bold = True
                        else:
                            st.range(row, col).api.Font.Bold = False
                        if sub[2][0] == 'Y':
                            st.range(row, col).api.Interior.Color = 65535
                        elif sub[2][0] == 'R':
                            st.range(row, col).api.Interior.Color = 255
                        else:
                            st.range(row, col).api.Interior.ColorIndex = 0

                        comment = sub[3]
                        if sub[2][2] == 'R':
                            if sub[2][0] == 'R':
                                comment.append('Early sub on')
                            else:
                                st.range(row, col).api.Font.Color = 255
                        if len(comment) > 0:
                            try:
                                st.range(row, col).api.AddComment('\n'.join(comment))
                            except:
                                st.range(row, col).api.Comment.Text('\n'.join(comment))
                            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
                            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
                st.range((row, 29), (row, 80)).value = result

                og = hal[2][1]
                if og > 0:
                    st.range(row, plist.index('OG') + 29).value = og
            else:
                print('%s not in ps' % self)

    def soccerway_download(self):
        print('***Soccerway Update***')
        soccerwayurl = input('Input Url:')
        if soccerwayurl == 'end':
            return 'end'
        x = self.sow_matchs(soccerwayurl)
        Game = self.list_to_elist(x)
        return Game


