from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import selenium
import datetime
import pprint

starttime = datetime.datetime.now()
# Chrome Headless
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
driver = webdriver.Chrome()

# 打开网址
driver.get('https://www.whoscored.com/Regions/81/Tournaments/3/Germany-Bundesliga')
driver.implicitly_wait(3)
endtime = datetime.datetime.now()
print('Urltime = ' + str((endtime - starttime).seconds) + 's')
# 射门数和控球率
# 选取元素
xpathid = 'tournament-fixture'


def whoscored_url_tournaments(driver, xpathid):


    def whoscored_url_tournaments_main(driver, xpathid):
        i = 1
        Matches = []
        date = ''
        while i < 100:
            try:
                tr = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]' % (xpathid, i)).get_attribute('class')
                if tr == 'rowgroupheader':
                    date = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]/th' % (xpathid, i)).text
                else:
                    result = driver.find_element_by_xpath(
                        '//*[@id="%s"]/tbody/tr[%s]/td[@class="result"]/a' % (xpathid, i)).text
                    if result != 'vs':
                        hteam = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]/td[4]/a' % (xpathid, i)).text
                        ateam = driver.find_element_by_xpath('//*[@id="%s"]/tbody/tr[%s]/td[6]/a' % (xpathid, i)).text
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
            print(Matches)

    return Matches


pprint.pprint(whoscored_url_tournaments(driver, xpathid))
endtime = datetime.datetime.now()
print('Alltime = ' + str((endtime - starttime).seconds) + 's')

