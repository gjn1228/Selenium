from selenium import webdriver

from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.common.keys import Keys


#Chrome Headless
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
#Chrome 打开soccerway联赛网址
driver= webdriver.Chrome(chrome_options=chrome_options)
#driver= webdriver.Chrome()
driver.get('https://int.soccerway.com/national/germany/bundesliga/20172018/regular-season/r41485/')
driver.implicitly_wait(3)
#选择gameweek选项
elem = driver.find_element_by_xpath('//*[@id="page_competition_1_block_competition_matches_summary_5_1_2"]')
elem.click()
driver.implicitly_wait(3)
#选择下拉菜单，和目标轮次
elem = driver.find_element_by_xpath('//*[@id="page_competition_1_block_competition_matches_summary_5_page_dropdown"]')
elem.click()
round = str(int(input('Round:')) - 1)
elem.find_element_by_xpath('//option[@value="%s"]'%round).click()
driver.implicitly_wait(3)


results = driver.find_element_by_xpath('//*[@id="page_competition_1_block_competition_matches_summary_5_match-2480339"]/td[2]')
print(results.text)


games_a = driver.find_elements_by_xpath(
    '//td[@class="score-time score"]/a ')
games = [('https://int.soccerway.com/' + elem.get_attribute('href')) for elem in games_a]
gamesss = [('No.' + str(games.index(elem)) + elem[25:]) for elem in games]
print('\n'.join(gamesss))


driver.quit()

#elem.send_keys('pycon')
#elem.send_keys(Keys.RETURN)
#print(driver.page_source)



