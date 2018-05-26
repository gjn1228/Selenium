from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import selenium
import datetime
import pprint



url = 'https://www.whoscored.com/Teams/60/Show/Spain-Alaves'



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
        # chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_extension('Adblock-Plus_v1.8.12.crx')
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



print(whoscored_url(url))



