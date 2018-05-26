from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import selenium
import datetime
import pprint


def whoscored_url(url):


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


def whoscored_board_timeline(driver):
    dataall = []
    # 射门数和控球率
    for i in [3, 5]:
        datas = []
        for j in [1, 3]:
            data = driver.find_element_by_xpath('//*[@id="match-centre-stats"]/ul/li[%s]/div[1]/span[%s]' % (i, j)).text
            datas.append(data)
        title = driver.find_element_by_xpath('//*[@id="match-centre-stats"]/ul/li[%s]/h4' % i).text

        datas.insert(0, title)
        dataall.append(datas)
    # Timeline
    i = 1
    timelineall = []
    while True:
        try:
            time = driver.find_element_by_xpath('//*[@id="live-incidents"]/div/table/tbody/tr[%s]/td[2]' % i).text
            # 获得事件
            for eventindex in [[1, 'home'], [3, 'away']]:
                j = eventindex[0]
                eventdiv1 = driver.find_element_by_xpath(
                    '//*[@id="live-incidents"]/div/table/tbody/tr[%s]/td[%s]/div[1]' % (i, j))
                eventclass = eventdiv1.get_attribute('class')
                if eventclass == 'clear':
                    pass
                else:
                    # 每一条timeline的每个div
                    x = 1
                    eventdiv = []
                    while True:
                        # 每个div里又分div、a、（span）
                        try:
                            datatitle = driver.find_element_by_xpath(
                                '//*[@id="live-incidents"]/div/table/tbody/tr[%s]/td[%s]/div[%s]' % (i, j, x)).get_attribute('title')
                            player = driver.find_element_by_xpath(
                                '//*[@id="live-incidents"]/div/table/tbody/tr[%s]/td[%s]/div[%s]/a' % (i, j, x)).text
                            div = driver.find_element_by_xpath(
                                '//*[@id="live-incidents"]/div/table/tbody/tr[%s]/td[%s]/div[%s]/div' % (i, j, x))
                            minute = div.get_attribute('data-minute')
                            second = div.get_attribute('data-second')
                            playerid = div.get_attribute('data-player-id')
                            datatype = div.get_attribute('data-type')
                            # 尝试current score
                            try:
                                currentscore = driver.find_element_by_xpath(
                                    '//*[@id="live-incidents"]/div/table/tbody/tr[%s]/td[%s]/div[%s]/span[@class="current-score"]' % (
                                    i, j, x)).text
                            except selenium.common.exceptions.NoSuchElementException:
                                currentscore = 'No Current Score'
                                pass
                            except Exception as inst:
                                print('Current score Span Error')
                                print(type(inst))
                                print(inst.args)
                                currentscore = 'Current Score Error'
                                pass
                            eventdiv.append([datatitle, player, playerid, datatype, minute, second, currentscore])
                            print(time, datatitle)
                            x += 1
                        except selenium.common.exceptions.NoSuchElementException:
                            break
                        except Exception as inst:
                            print('Timeline Div Error')
                            print(type(inst))
                            print(inst.args)
                            x += 1
                    h_a = eventindex[1]
                    timelineall.append([time, h_a, eventdiv])
            i += 1
        except selenium.common.exceptions.NoSuchElementException:
            break
        except Exception as inst:
            print('Timeline Error')
            print(type(inst))
            print(inst.args)
            i += 1
    dataall.append(timelineall)
    return dataall


def whoscored_chalkboard(driver):
    #选取元素
    #射门数据
    shotall = []
    for h in range(4):
        indexh = [6, 3, 5, 4]
        results = []
        for i in range(2, indexh[h]+2):
            shot = []
            for j in range(2):
                sot = driver.find_element_by_xpath(
                    '//*[@id="chalkboard"]/div[2]/div[1]/div[%s]/div[%s]/span[%s]' % (h + 2, i, j+1)).text
                shot.append(sot)
            label = driver.find_element_by_xpath(
                '//*[@id="chalkboard"]/div[2]/div[1]/div[%s]/div[%s]/label' % (h + 2, i)).text
            shot.insert(0, label)
            results.append(shot)
        shotall.append(results)
    return shotall


def whoscored_player_stat(driver):

    alldata = []
    # 射门信息
    allshotdata = []
    try:
        for i in range(1, 7):
            shotdata = []
            shota = driver.find_element_by_xpath(
                '//*[@id="match-report-team-statistics"]/div[1]/div[%s]/span[1]/span' % i).text
            shotb = driver.find_element_by_xpath(
                '//*[@id="match-report-team-statistics"]/div[1]/div[%s]/span[2]' % i).text
            shotc = driver.find_element_by_xpath(
                '//*[@id="match-report-team-statistics"]/div[1]/div[%s]/span[3]/span' % i).text
            shotdata.append(shotb)
            shotdata.append(shota)
            shotdata.append(shotc)
            allshotdata.append(shotdata)
        print('Shot Data Downloaded')
    except Exception as inst:
        print('Shot Data Error')
        print(type(inst))
        print(inst.args)
    alldata.append(allshotdata)
    # 控球率
    Possessiondata = []
    try:
        shotdata = []
        possa = driver.find_element_by_xpath(
            '//*[@id="match-report-team-statistics"]/div[2]/div[2]/span/span[2]/span').text
        possb = driver.find_element_by_xpath(
            '//*[@id="match-report-team-statistics"]/div[2]/div[2]/span/span[3]/span').text
        Possessiondata.append(possa)
        Possessiondata.append(possb)
    except Exception as inst:
        print('Possession Error')
        print(type(inst))
        print(inst.args)
    alldata.append(Possessiondata)
    # 球员信息
    xpathlist = ['*[@id="statistics-table-home-summary"]/table/tbody',
                 '*[@id="statistics-table-away-summary"]/table/tbody']
    for xp in xpathlist:
        homeaway = ['Home', 'Away']
        i = 1
        while i < 30:
            playeralldata = []
            try:
                playerdata = []
                # 名字
                player = driver.find_element_by_xpath(
                    '//%s/tr[%s]/td[@class="pn"]/a' % (xp, i)).text
                playerid = driver.find_element_by_xpath(
                    '//%s/tr[%s]/td[@class="pn"]/a' % (xp, i)).get_attribute('href')
                # 换上换下
                try:
                    inouttime = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="pn"]/a/span' % (xp, i)).text
                    if len(inouttime) > 0:
                        inouttype = driver.find_element_by_xpath(
                            '//%s/tr[%s]/td[@class="pn"]/a/span/span' % (xp, i)).get_attribute(
                            'data-type')
                        playerid = driver.find_element_by_xpath(
                            '//%s/tr[%s]/td[@class="pn"]/a/span/span' % (xp, i)).get_attribute(
                            'data-player-id')
                    else:
                        inouttype = ''
                        inouttime = ''
                    if inouttime != '':
                        player = player.replace(inouttime, '')
                    playerdata.append(player)
                    playerdata.append(playerid)
                    playerdata.append(inouttype)
                    playerdata.append(inouttime)
                except selenium.common.exceptions.NoSuchElementException:
                    playerdata.append(player)
                    playerdata.append(playerid)
                    playerdata.append(0)
                    playerdata.append(0)
                except Exception as inst:
                    print('Error in ' + player)
                    print(type(inst))
                    print(inst.args)
                playeralldata.append(playerdata)
                playerdata = []
                # 年龄和位置
                try:
                    # 年龄
                    age = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="pn"]/span[1]' % (xp, i)).text
                    playerdata.append(age)
                    # 位置
                    pos = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="pn"]/span[2]' % (xp, i)).text
                    playerdata.append(pos)
                except Exception as inst:
                    print(player + ' Age or Position Error')
                    print(type(inst))
                    print(inst.args)
                playeralldata.append(playerdata)
                playerdata = []
                # 关键数据
                try:
                    # Rating
                    rating = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="rating "]' % (xp, i)).text
                    playerdata.append(rating)
                    # Key Events
                    j = 1
                    allevent = []
                    while j < 10:
                        events = []
                        try:
                            event = driver.find_element_by_xpath(
                                '//%s/tr[%s]/td[@style="text-align: left"]/span[@class="incident-wrapper"]/span[%s]' % (
                                    xp, i, j)).get_attribute('data-type')
                            eventminute = driver.find_element_by_xpath(
                                '//%s/tr[%s]/td[@style="text-align: left"]/span[@class="incident-wrapper"]/span[%s]' % (
                                    xp, i, j)).get_attribute('data-minute')
                            eventsecond = driver.find_element_by_xpath(
                                '//%s/tr[%s]/td[@style="text-align: left"]/span[@class="incident-wrapper"]/span[%s]' % (
                                    xp, i, j)).get_attribute('data-second')
                            events.append(event)
                            events.append(eventminute)
                            events.append(eventsecond)
                            allevent.append(events)
                            j += 1
                        except selenium.common.exceptions.NoSuchElementException:
                            break
                        except Exception as inst:
                            print(player + ' Events Error')
                            print(type(inst))
                            print(inst.args)
                            i += 1
                    playerdata.append(allevent)
                except Exception as inst:
                    print(player + ' Key Data Error')
                    print(type(inst))
                    print(inst.args)
                playeralldata.append(playerdata)
                playerdata = []
                # 其他数据
                try:
                    # 射门数
                    shottotal = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="ShotsTotal "]' % (xp, i)).text
                    playerdata.append(shottotal)
                    # 射正数
                    shotontarget = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="ShotOnTarget "]' % (xp, i)).text
                    playerdata.append(shotontarget)
                    # 关键传球数
                    keypass = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="KeyPassTotal "]' % (xp, i)).text
                    playerdata.append(keypass)
                    # 传球成功率
                    passsuccess = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="PassSuccessInMatch "]' % (xp, i)).text
                    playerdata.append(passsuccess)
                    # 空中争抢
                    DuelAerialWon = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="DuelAerialWon "]' % (xp, i)).text
                    playerdata.append(DuelAerialWon)
                    # 触球数
                    Touches = driver.find_element_by_xpath(
                        '//%s/tr[%s]/td[@class="Touches "]' % (xp, i)).text
                    playerdata.append(Touches)
                except Exception as inst:
                    print(player + ' Other Data Error')
                    print(type(inst))
                    print(inst.args)
                i += 1
                playeralldata.append(playerdata)
                playerdata = []
                alldata.append(playeralldata)
            except selenium.common.exceptions.NoSuchElementException:
                print(homeaway[xpathlist.index(xp)] + ' Team Total ' + str(i - 1) + ' Player Downloaded')
                break
            except Exception as inst:
                print(type(inst))
                print(inst.args)
                i += 1

    return alldata


def whoscored_board_top5_shot(driver):
    #射门数和控球率
    dataall = []
    for i in [3, 5]:
        datas = []
        for j in [1, 3]:
            data = driver.find_element_by_xpath('//*[@id="match-centre-stats"]/ul/li[%s]/div[1]/span[%s]' % (i, j)).text
            datas.append(data)
        title = driver.find_element_by_xpath('//*[@id="match-centre-stats"]/ul/li[%s]/h4' % i).text

        datas.insert(0, title)
        dataall.append(datas)
    # 打开Rating tab
    elem = driver.find_element_by_xpath('//*[@id="match-centre-stats"]/ul/li[1]/div[2]/span')
    elem.click()
    driver.implicitly_wait(3)
    # 读取top5
    top5 = []
    for i in range(5):
        j = i + 1
        top1 = []
        for k in range(4):
            m = k + 2
            top = driver.find_element_by_xpath(
                '//*[@id="match-centre-stats"]/ul/li[2]/div[1]/div/div[%s]/div/span[%s]' % (j, m)).text
            top1.append(top)
        top5.append(top1)
    dataall.append(top5)
    # 打开Shots tab
    elem = driver.find_element_by_xpath('//*[@id="match-centre-stats"]/ul/li[1]/div[2]/span')
    elem.click()
    driver.implicitly_wait(3)
    elem = driver.find_element_by_xpath('//*[@id="match-centre-stats"]/ul/li[3]/div[2]/span')
    elem.click()
    driver.implicitly_wait(3)
    # 读取射门数据
    shots = []
    for i in range(5):
        j = i + 1
        shota = driver.find_element_by_xpath(
            '//*[@id="match-centre-stats"]/ul/li[4]/div[1]/ul/li[%s]/div/span[3]' % j).text
        shoth = driver.find_element_by_xpath(
            '//*[@id="match-centre-stats"]/ul/li[4]/div[1]/ul/li[%s]/div/span[1]' % j).text
        shott = driver.find_element_by_xpath(
            '//*[@id="match-centre-stats"]/ul/li[4]/div[1]/ul/li[%s]/h4' % j).text
        shots.append([shott, shoth, shota])
    dataall.append(shots)
    return dataall


def get_whoscored_top5_shot(parameter):
    driver = parameter[0]
    starttime = parameter[1]
    # 读取Top5\shot
    whoscored1 = whoscored_board_top5_shot(driver)
    pprint.pprint(whoscored1)
    endtime = datetime.datetime.now()
    print('Top5\shot Run Time ' + str((endtime - starttime).seconds) + 's')

    driver.quit()


def get_whoscored_timeline_player(parameter):
    driver = parameter[0]
    starttime = parameter[1]
    # 读取Board页面
    whoscored1 = whoscored_board_timeline(driver)
    pprint.pprint(whoscored1)
    endtime = datetime.datetime.now()
    print('Board Run Time ' + str((endtime - starttime).seconds) + 's')

    # 打开Player Stats页面
    elem = driver.find_element_by_xpath('//*[@id="sub-sub-navigation"]/ul/li[2]/a')
    elem.click()
    driver.implicitly_wait(3)
    endtime = datetime.datetime.now()
    print('Player Stats Urltime = ' + str((endtime - starttime).seconds) + 's')

    # 读取Player Stats页面
    whoscored2 = whoscored_player_stat(driver)
    pprint.pprint(whoscored2)
    endtime = datetime.datetime.now()
    print('Player Stats Run Time ' + str((endtime - starttime).seconds) + 's')

    driver.quit()


def get_whoscored_full_stat(parameter):
    driver = parameter[0]
    starttime = parameter[1]
    # 读取Timeline
    whoscored1 = whoscored_board_timeline(driver)
    pprint.pprint(whoscored1)
    endtime = datetime.datetime.now()
    print('Board Run Time ' + str((endtime - starttime).seconds) + 's')

    # 读取Top5\shot
    whoscored2 = whoscored_board_top5_shot(driver)
    pprint.pprint(whoscored2)
    endtime = datetime.datetime.now()
    print('Board Run Time ' + str((endtime - starttime).seconds) + 's')

    # 打开chalkboard页面
    elem = driver.find_element_by_xpath('//*[@id="live-match-options"]/li[3]/a/span[2]')
    elem.click()
    driver.implicitly_wait(3)
    endtime = datetime.datetime.now()
    print('Chalkboard UrlTime = ' + str((endtime - starttime).seconds) + 's')

    # 读取chalkboard页面
    whoscored3 = whoscored_chalkboard(driver)
    pprint.pprint(whoscored3)
    endtime = datetime.datetime.now()
    print('Chalkboard Run Time ' + str((endtime - starttime).seconds) + 's')

    # 打开Player Stats页面
    elem = driver.find_element_by_xpath('//*[@id="sub-sub-navigation"]/ul/li[2]/a')
    elem.click()
    driver.implicitly_wait(3)
    endtime = datetime.datetime.now()
    print('Player Stats Urltime = ' + str((endtime - starttime).seconds) + 's')

    # 读取Player Stats页面
    whoscored4 = whoscored_player_stat(driver)
    pprint.pprint(whoscored4)
    endtime = datetime.datetime.now()
    print('Player Stats Run Time ' + str((endtime - starttime).seconds) + 's')

    driver.quit()


def init_whoscored_matchcenter(matchcenterurl):
    starttime = datetime.datetime.now()
    # Chrome Headless
    chrome_options = Options()
    #chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    #chrome_options.add_extension('Adblock-Plus_v1.8.12.crx')
    driver = webdriver.Chrome(chrome_options=chrome_options)

    # 打开网址
    driver.get(matchcenterurl)
    driver.implicitly_wait(3)
    endtime = datetime.datetime.now()
    print('Board Urltime = ' + str((endtime - starttime).seconds) + 's')

    return [driver, starttime]


def main_whoscored(matchcenterurl):
    while True:
        print('-----------------------------------------------')
        print('|   Get Top 5 & Shot From Whoscored, Ipput 1  |')
        print('-----------------------------------------------')
        print('|Get Timeline & Player From Whoscored, Input 2|')
        print('-----------------------------------------------')
        print('|    Get Full Data From Whoscored, Input 3    |')
        print('-----------------------------------------------')
        ip = input('             Choose a Function:')
        print('--------------------Loading--------------------')
        urls = whoscored_url(matchcenterurl)
        pprint.pprint(urls)
        for url in urls:
            parameter = init_whoscored_matchcenter(url)
            if ip == '1':
                print('***Top 5 & Shot***')
                print('------------------------------')
                get_whoscored_top5_shot(parameter)
            elif ip == '2':
                print('***Timeline & Player***')
                print('------------------------------')
                get_whoscored_timeline_player(parameter)
            elif ip == '3':
                print('***Full Data***')
                print('------------------------------')
                get_whoscored_full_stat(parameter)
            else:
                print('Wrong Input, Try Again')
                print('------------------------------')


#'https://www.whoscored.com/Matches/1201922/Live/Germany-Bundesliga-2017-2018-Borussia-Dortmund-Schalke-04'
url = 'https://www.whoscored.com/Regions/81/Tournaments/3/Seasons/6902/Stages/15243/Fixtures/Germany-Bundesliga-2017-2018'
main_whoscored(url)


