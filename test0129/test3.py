from selenium import webdriver
from selenium.webdriver.chrome.options import Options


#driver = webdriver.Chrome(executable_path=r"C:\Users\嘉楠\PycharmProjects\chromedriver\chromedriver.exe")
driver = webdriver.Chrome(executable_path=r"D:\chromedriver.exe")

driver.get('https://www.whoscored.com/Matches/1222117/Live/Spain-La-Liga-2017-2018-Eibar-Malaga')
driver.implicitly_wait(3)

elem = driver.find_element_by_xpath('//*[@id="live-match-options"]/li[3]/a/span[2]')
elem.click()
driver.implicitly_wait(3)

elem = driver.find_element_by_xpath('//*[@id="chalkboard"]/div[2]/div[1]/div[2]/div[4]/span[1]')
print(elem.text)



end = input('Input Any Key To End:')



