# coding=utf-8 #
# Author GJN #

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import json
import time
import pandas as pd
import xlwings as xw
import lxml.html
import datetime
import requests
import unicodedata


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
                        '|' + y[0] + '|' + '/'.join([y[3], y[4], y[5]]) + '| |' + y[7] + '| |' + y[8] + '| |' + y[
                            9] + '|')
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
                        if int(''.join([elem.text_content() for elem in event_time_a]).split('\'')[0].split('+')[
                                   0]) <= 90:
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
                        if int(''.join([elem.text_content() for elem in event_time_b]).split('\'')[0].split('+')[
                                   0]) <= 90:
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
            self.selfname = info[0]
            if self.selfname in soc_ps_dic:
                st = wb.sheets[soc_ps_dic[self.selfname]]
                ha = ['H', 'A'][elist.index(t) % 2]
                ang = info[1][1]
                if ang in soc_game_dic:
                    ang = soc_game_dic[ang]
                self.soccerway_input(st, ha, ang, t)

            else:
                print('%s not in ps' % self.selfname)

    def elist_to_pfl(self, elist):
        # path = r'E:\Company\League\PFL\\'
        # df = pd.read_excel(path + 'team name.xlsx')
        # soc_ps_dic = df[['PS', 'Soccerway']].set_index('Soccerway').to_dict()['PS']
        # soc_game_dic = df[['Whoscored', 'PS']].set_index('Whoscored').to_dict()['PS']

        soc_ps_dic = {'Belenenses': 'Belenenses',
 'Benfica': 'Benfica',
 'Boavista': 'Boavista',
 'Chaves': 'Chaves',
 'Desportivo Aves': 'Desportivo_Aves',
 'Feirense': 'Feirense',
 'Marítimo': 'Marítimo',
 'Moreirense': 'Moreirense',
 'Nacional': 'Nacional',
 'Portimonense': 'Portimonense',
 'Porto': 'Porto',
 'Rio Ave': 'Rio_Ave',
 'Santa Clara': 'Santa_Clara',
 'Sporting Braga': 'Braga',
 'Sporting CP': 'Sporting_CP',
 'Tondela': 'Tondela',
 'Vitória Guimarães': 'Vitória_Guimarães',
 'Vitória Setúbal': 'Vitória_Setúbal'}
        wb = xw.Book('PFL_Player_Sheet_18-19.xlsm')
        for t in elist:
            info = t[0]
            self.selfname = info[0]
            if self.selfname in soc_ps_dic:
                st = wb.sheets[soc_ps_dic[self.selfname]]
                ha = ['H', 'A'][elist.index(t) % 2]
                ang = info[1][1]
                self.soccerway_input(st, ha, ang, t)

            else:
                print('%s not in ps' % self.selfname)

    def soccerway_download(self):
        print('***Soccerway Update***')
        soccerwayurl = input('Input Url:')
        if soccerwayurl == 'end':
            return 'end'
        x = self.sow_matchs(soccerwayurl)
        Game = self.list_to_elist(x)
        return Game

    # 取消input， 直接输入地址
    def soccerway_download2(self, url):
        print('***Soccerway Update***')
        x = self.sow_matchs(url)
        Game = self.list_to_elist(x)
        return Game

    def socway_player(self, url):
        # 网址
        headers = {
            'User-Agent': "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1"}
        all_url = 'https://int.soccerway.com/' + url
        html = requests.get(all_url, headers=headers)
        tree = lxml.html.fromstring(html.text)
        player_dt = tree.xpath(
            '//*[@id="page_player_1_block_player_passport_3"]/div/div/div[@class="yui-u first"]/div/dl/dt')
        player_dd = tree.xpath(
            '//*[@id="page_player_1_block_player_passport_3"]/div/div/div[@class="yui-u first"]/div/dl/dd')
        player_dt_t = [elem.text_content() for elem in player_dt]
        player_dd_t = [elem.text_content() for elem in player_dd]

        if 'Age' in player_dt_t:
            player_age = player_dd_t[player_dt_t.index('Age')]
        else:
            player_age = ''
        pos_dic = {'Goalkeeper': 'GK', 'Defender': 'D', 'Midfielder': 'M', 'Attacker': 'A'}
        if 'Position' in player_dt_t:
            player_pos = pos_dic[player_dd_t[player_dt_t.index('Position')]]
        else:
            player_pos = ''
        return [player_pos, player_age]

    def soccerway_input(self, st, ha, ang, t):
        game = st.range((10, 1), (80, 4)).value
        row = -1
        for g in game:
            if g[3] != 'PRL':
                if g[2] == ha:
                    if g[1] == ang:
                        row = game.index(g) + 10
        if row == -1:
            row = int(input('%s row:' % self.selfname))

        plist = st.range((8, 29), (8, 120)).value

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

        blank = len(plist) - 1
        while plist[blank] is None:
            blank -= 1
        blank += 2
        nlist = st.range((6, 29), (6, 120)).value
        result = st.range((row, 29), (row, 120)).value
        st.range((row, 29), (row, 120)).api.Font.Bold = False
        st.range((row, 29), (row, 120)).api.Font.ColorIndex = 1

        home = t[1][0]
        away = t[1][1]
        st.range(row, 5).value = home[0]
        st.range(row, 6).value = away[0]
        if len(home[2]) > 0:
            self.write_comment(st, row, 5, home[2])
        if len(away[2]) > 0:
            self.write_comment(st, row, 6, away[2])
        if home[1][0] == ['Red']:
            st.range(row, 5).api.Interior.Color = 255
        if away[1][0] == ['Red']:
            st.range(row, 6).api.Interior.Color = 255

        hal = t[2]

        for player in hal[0]:
            ind = 0
            name = player[0]
            name2 = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('UTF-8')
            number = player[5]
            if name in plist:
                ind = plist.index(name)
            elif name in plist2:
                ind = plist2.index(name)
            elif name2 in plist2:
                ind = plist2.index(name2)
                print(st.name, '-', name, ' Changed to ', name2)
            # elif number in nlist:
            #     ind = nlist.index(number)
            # elif str(number) in nlist:
            #     ind = nlist.index(str(number))

            if ind == 0:
                url = player[4][0]
                pos_age = self.socway_player(url)
                ind = blank
                blank += 1
                col = ind + 29
                st.range(7, col).value = pos_age[0]
                st.range(8, col).value = name
                st.range(8, col).api.Interior.ColorIndex = 48
                st.range(9, col).value = pos_age[1]
                print(st.name, '-', name, ' or ', name2, ' not Found, Added')

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
                self.write_comment(st, row, col, player[3])

        for sub in hal[1]:
            ind = 0
            name = sub[0]
            name2 = unicodedata.normalize('NFD', name).encode('ascii', 'ignore').decode('UTF-8')
            number = sub[5]
            if name in plist:
                ind = plist.index(name)
            elif name in plist2:
                ind = plist2.index(name)
            elif name2 in plist2:
                ind = plist2.index(name2)
                print(st.name, '-', name, ' Changed to ', name2)
            # elif number in nlist:
            #     ind = nlist.index(number)
            # elif str(number) in nlist:
            #     ind = nlist.index(str(number))

            if ind == 0:
                url = sub[4][0]
                pos_age = self.socway_player(url)
                ind = blank
                blank += 1
                col = ind + 29
                st.range(7, col).value = pos_age[0]
                st.range(8, col).value = name
                st.range(8, col).api.Interior.ColorIndex = 48
                st.range(9, col).value = pos_age[1]
                print(st.name, '-', name, ' or ', name2,  ' not Found, Added')

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
                self.write_comment(st, row, col, comment)
        st.range((row, 29), (row, 80)).value = result

        og = hal[2][1]
        if og > 0:
            if 'OG' not in plist:
                plist[blank] = 'OG'
                ind = blank
                blank += 1
                col = ind + 29
                st.range(8, col).value = 'OG'
                st.range(8, col).api.Font.Bold = True

            st.range(row, plist.index('OG') + 29).value = og

    def elist_to_unl(self, elist):
        for t in elist:
            psd = {'Albania': 'UNL_A-D.xlsm',
                     'Andorra': 'UNL_A-D.xlsm',
                     'Armenia': 'UNL_A-D.xlsm',
                     'Austria': 'UNL_A-D.xlsm',
                     'Azerbaijan': 'UNL_A-D.xlsm',
                     'Belarus': 'UNL_A-D.xlsm',
                     'Belgium': 'UNL_A-D.xlsm',
                     'Bosnia-Herzegovina': 'UNL_A-D.xlsm',
                     'Bulgaria': 'UNL_A-D.xlsm',
                     'Croatia': 'UNL_A-D.xlsm',
                     'Cyprus': 'UNL_A-D.xlsm',
                     'Czech Republic': 'UNL_A-D.xlsm',
                     'Denmark': 'UNL_A-D.xlsm',
                     'England': 'UNL_E-I.xlsm',
                     'Estonia': 'UNL_E-I.xlsm',
                     'Faroe Islands': 'UNL_E-I.xlsm',
                     'Finland': 'UNL_E-I.xlsm',
                     'France': 'UNL_E-I.xlsm',
                     'Georgia': 'UNL_E-I.xlsm',
                     'Germany': 'UNL_E-I.xlsm',
                     'Gibraltar': 'UNL_E-I.xlsm',
                     'Greece': 'UNL_E-I.xlsm',
                     'Hungary': 'UNL_E-I.xlsm',
                     'Iceland': 'UNL_E-I.xlsm',
                     'Israel': 'UNL_E-I.xlsm',
                     'Italy': 'UNL_E-I.xlsm',
                     'Kazakhstan': 'UNL_K-P.xlsm',
                     'Kosovo': 'UNL_K-P.xlsm',
                     'Latvia': 'UNL_K-P.xlsm',
                     'Liechtenstein': 'UNL_K-P.xlsm',
                     'Lithuania': 'UNL_K-P.xlsm',
                     'Luxembourg': 'UNL_K-P.xlsm',
                     'Macedonia': 'UNL_K-P.xlsm',
                     'Malta': 'UNL_K-P.xlsm',
                     'Moldova': 'UNL_K-P.xlsm',
                     'Montenegro': 'UNL_K-P.xlsm',
                     'Netherlands': 'UNL_K-P.xlsm',
                     'Northern Ireland': 'UNL_K-P.xlsm',
                     'Norway': 'UNL_K-P.xlsm',
                     'Poland': 'UNL_K-P.xlsm',
                     'Portugal': 'UNL_K-P.xlsm',
                     'Republic of Ireland': 'UNL_R-W.xlsm',
                     'Romania': 'UNL_R-W.xlsm',
                     'Russia': 'UNL_R-W.xlsm',
                     'San Marino': 'UNL_R-W.xlsm',
                     'Scotland': 'UNL_R-W.xlsm',
                     'Serbia': 'UNL_R-W.xlsm',
                     'Slovakia': 'UNL_R-W.xlsm',
                     'Slovenia': 'UNL_R-W.xlsm',
                     'Spain': 'UNL_R-W.xlsm',
                     'Sweden': 'UNL_R-W.xlsm',
                     'Switzerland': 'UNL_R-W.xlsm',
                     'Turkey': 'UNL_R-W.xlsm',
                     'Ukraine': 'UNL_R-W.xlsm',
                     'Wales': 'UNL_R-W.xlsm'}
            info = t[0]
            self.selfname = info[0]
            if self.selfname == 'FYR Macedonia':
                self.selfname = 'Macedonia'
            if self.selfname in psd:
                wb = xw.Book(psd[self.selfname])
                st = wb.sheets[self.selfname]
                ha = ['H', 'A'][elist.index(t) % 2]
                ang = info[1][1]
                self.soccerway_input(st, ha, ang, t)

            else:
                print('%s not in ps' % self.selfname)

    def write_comment(self, st, row, col, com):
        try:
            st.range(row, col).api.AddComment('\n'.join(com))
        except:
            st.range(row, col).api.Comment.Text('\n'.join(com))
        try:
            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Bold = True
            st.range(row, col).api.Comment.Shape.TextFrame.Characters().Font.Name = 'Tahoma'
        except:
            pass











