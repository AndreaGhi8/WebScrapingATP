from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd

url = "https://www.atptour.com/en/stats/leaderboard"

fix1 = "?boardType=serve&timeFrame="
fix2 = "&surface="
fix3 = "&versusRank=all&formerNo1=false"
x1 = 1991

PATH = "C:\Program Files (x86)\chromedriver.exe"

tab1 = pd.DataFrame(columns= ['Year', 'Avg. SR', 'Avg. SR CLAY', 'Avg. SR GRASS', 'Avg. SR HARD', 'Avg. SR TOP 10', 'Avg. SR CLAY TOP 10', 'Avg. SR GRASS TOP 10', 'Avg. SR HARD TOP 10'])
tab2 = pd.DataFrame(columns= ['Year', 'Avg. Aces', 'Avg. Aces CLAY', 'Avg. Aces GRASS', 'Avg. Aces HARD', 'Avg. Aces TOP 10', 'Avg. Aces CLAY TOP 10', 'Avg. Aces GRASS TOP 10', 'Avg. Aces HARD TOP 10'])

for years in range(32):
    x11 = str(x1)
    count = 1

    for surfaces in range(4):

        if(count == 1):
            x2 = "all"

            tab = pd.DataFrame(columns= ['Number', 'Player Name', 'Serve Rating', 'Avg. Aces/Match'])
            
            driver = webdriver.Chrome(PATH)
            url1 = url + fix1 + x11 + fix2 + x2 + fix3
            driver.get(url1)

            p = driver.find_elements(By.CLASS_NAME, "stats-listing-row")
            c1 = 0
            total_sr = 0
            total_sr_top10 = 0
            total_aces = 0
            total_aces_top10 = 0
            for m in p:
                c1 += 1
                c11 = str(c1)

                number = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[1]")
                name = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[2]")
                serve_rating = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[3]")
                aces = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[8]")

                temp = pd.DataFrame([[number.text, name.text, serve_rating.text, aces.text]], columns = ['Number', 'Player Name', 'Serve Rating', 'Avg. Aces/Match'])
                tab = pd.concat([tab, temp], ignore_index = True)

                total_sr += float(serve_rating.text)
                total_aces += float(aces.text)

                if(c1 <= 10):
                    total_sr_top10 = total_sr
                    total_aces_top10 = total_aces

            avg_sr_all = total_sr/c1
            avg_sr_all_top10 = total_sr_top10/10
            avg_aces_all = total_aces/c1
            avg_aces_all_top10 = total_aces_top10/10

            avg_sr_all = ("%.2f" % avg_sr_all)
            avg_sr_all_top10 = ("%.2f" % avg_sr_all_top10)
            avg_aces_all = ("%.2f" % avg_aces_all)
            avg_aces_all_top10 = ("%.2f" % avg_aces_all_top10)

            tab.to_excel('' + x11 + "_" + x2 + '.xlsx', index = False)
            driver.quit()
            count += 1

        elif(count == 2):
            x2 = "clay"

            tab = pd.DataFrame(columns= ['Number', 'Player Name', 'Serve Rating', 'Avg. Aces/Match'])

            driver = webdriver.Chrome(PATH)
            url1 = url + fix1 + x11 + fix2 + x2 + fix3
            driver.get(url1)

            p = driver.find_elements(By.CLASS_NAME, "stats-listing-row")
            c1 = 0
            total_sr = 0
            total_sr_top10 = 0
            total_aces = 0
            total_aces_top10 = 0
            for m in p:
                c1 += 1
                c11 = str(c1)

                number = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[1]")
                name = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[2]")
                serve_rating = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[3]")
                aces = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[8]")

                temp = pd.DataFrame([[number.text, name.text, serve_rating.text, aces.text]], columns = ['Number', 'Player Name', 'Serve Rating', 'Avg. Aces/Match'])
                tab = pd.concat([tab, temp], ignore_index = True)

                total_sr += float(serve_rating.text)
                total_aces += float(aces.text)

                if(c1 <= 10):
                    total_sr_top10 = total_sr
                    total_aces_top10 = total_aces
                
            avg_sr_clay = total_sr/c1
            avg_sr_clay_top10 = total_sr_top10/10
            avg_aces_clay = total_aces/c1
            avg_aces_clay_top10 = total_aces_top10/10

            avg_sr_clay = ("%.2f" % avg_sr_clay)
            avg_sr_clay_top10 = ("%.2f" % avg_sr_clay_top10)
            avg_aces_clay = ("%.2f" % avg_aces_clay)
            avg_aces_clay_top10 = ("%.2f" % avg_aces_clay_top10)

            tab.to_excel('' + x11 + "_" + x2 + '.xlsx', index = False)
            driver.quit()
            count += 1
        
        elif(count == 3):
            x2 = "grass"

            tab = pd.DataFrame(columns= ['Number', 'Player Name', 'Serve Rating', 'Avg. Aces/Match'])

            driver = webdriver.Chrome(PATH)
            url1 = url + fix1 + x11 + fix2 + x2 + fix3
            driver.get(url1)

            p = driver.find_elements(By.CLASS_NAME, "stats-listing-row")
            c1 = 0
            total_sr = 0
            total_sr_top10 = 0
            total_aces = 0
            total_aces_top10 = 0
            for m in p:
                c1 += 1
                c11 = str(c1)

                number = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[1]")
                name = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[2]")
                serve_rating = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[3]")
                aces = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[8]")

                temp = pd.DataFrame([[number.text, name.text, serve_rating.text, aces.text]], columns = ['Number', 'Player Name', 'Serve Rating', 'Avg. Aces/Match'])
                tab = pd.concat([tab, temp], ignore_index = True)

                total_sr += float(serve_rating.text)
                total_aces += float(aces.text)

                if(c1 <= 10):
                    total_sr_top10 = total_sr
                    total_aces_top10 = total_aces

            if(x1 == 2020):             #nel 2020 non si è disputata la stagione sull'erba a causa della pandemina di covid-19
                avg_sr_grass = 0
                avg_sr_grass_top10 = 0
                avg_aces_grass = 0
                avg_aces_grass_top10 = 0
            elif(x1 != 2020):
                avg_sr_grass = total_sr/c1
                avg_sr_grass_top10 = total_sr_top10/10
                avg_aces_grass = total_aces/c1
                avg_aces_grass_top10 = total_aces_top10/10

                avg_sr_grass = ("%.2f" % avg_sr_grass)
                avg_sr_grass_top10 = ("%.2f" % avg_sr_grass_top10)
                avg_aces_grass = ("%.2f" % avg_aces_grass)
                avg_aces_grass_top10 = ("%.2f" % avg_aces_grass_top10)

            tab.to_excel('' + x11 + "_" + x2 + '.xlsx', index = False)
            driver.quit()
            count += 1
        
        elif(count == 4):
            x2 = "hard"

            tab = pd.DataFrame(columns= ['Number', 'Player Name', 'Serve Rating', 'Avg. Aces/Match'])

            driver = webdriver.Chrome(PATH)
            url1 = url + fix1 + x11 + fix2 + x2 + fix3
            driver.get(url1)

            p = driver.find_elements(By.CLASS_NAME, "stats-listing-row")
            c1 = 0
            total_sr = 0
            total_sr_top10 = 0
            total_aces = 0
            total_aces_top10 = 0
            for m in p:
                c1 += 1
                c11 = str(c1)

                number = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[1]")
                name = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[2]")
                serve_rating = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[3]")
                aces = m.find_element(By.XPATH, "//tr[" + c11 +"]/td[8]")

                temp = pd.DataFrame([[number.text, name.text, serve_rating.text, aces.text]], columns = ['Number', 'Player Name', 'Serve Rating', 'Avg. Aces/Match'])
                tab = pd.concat([tab, temp], ignore_index = True)

                total_sr += float(serve_rating.text)
                total_aces += float(aces.text)

                if(c1 <= 10):
                    total_sr_top10 = total_sr
                    total_aces_top10 = total_aces
                
            avg_sr_hard = total_sr/c1
            avg_sr_hard_top10 = total_sr_top10/10
            avg_aces_hard = total_aces/c1
            avg_aces_hard_top10 = total_aces_top10/10

            avg_sr_hard = ("%.2f" % avg_sr_hard)
            avg_sr_hard_top10 = ("%.2f" % avg_sr_hard_top10)
            avg_aces_hard = ("%.2f" % avg_aces_hard)
            avg_aces_hard_top10 = ("%.2f" % avg_aces_hard_top10)

            temp1 = pd.DataFrame([[x1, avg_sr_all, avg_sr_clay, avg_sr_grass, avg_sr_hard, avg_sr_all_top10, avg_sr_clay_top10, avg_sr_grass_top10, avg_sr_hard_top10]], columns = ['Year', 'Avg. SR', 'Avg. SR CLAY', 'Avg. SR GRASS', 'Avg. SR HARD', 'Avg. SR TOP 10', 'Avg. SR CLAY TOP 10', 'Avg. SR GRASS TOP 10', 'Avg. SR HARD TOP 10'])
            temp2 = pd.DataFrame([[x1, avg_aces_all, avg_aces_clay, avg_aces_grass, avg_aces_hard, avg_aces_all_top10, avg_aces_clay_top10, avg_aces_grass_top10, avg_aces_hard_top10]], columns = ['Year', 'Avg. Aces', 'Avg. Aces CLAY', 'Avg. Aces GRASS', 'Avg. Aces HARD', 'Avg. Aces TOP 10', 'Avg. Aces CLAY TOP 10', 'Avg. Aces GRASS TOP 10', 'Avg. Aces HARD TOP 10'])
            tab1 = pd.concat([tab1, temp1], ignore_index = True)
            tab2 = pd.concat([tab2, temp2], ignore_index = True)

            tab.to_excel('' + x11 + "_" + x2 + '.xlsx', index = False)
            driver.quit()
            count += 1

    x1 += 1

tab1.to_excel('Avg_Serve_Rating.xlsx', index = False)
tab2.to_excel('Avg_Aces.xlsx', index = False)