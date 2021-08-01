import pandas as pd
import requests
from bs4 import BeautifulSoup


def get_url(mode='系所', dpt='9010', sem='110-1', startrec=0):
    if type(dpt) == int:
        dpt = str(dpt)
    if len(dpt) == 3:
        dpt += '0'
    if mode == '系所':
        return "https://nol.ntu.edu.tw/nol/coursesearch/search_for_02_dpt.php?alltime=yes&allproced=yes&selcode=-1&dptname=" + dpt + \
               "&coursename&teachername&current_sem=" + sem + \
               "&yearcode=0&op&startrec=0&week1&week2&week3&week4&week5&week6&proced0&proced1&proced2&proced3&proced4&procedE&proced5&proced6&proced7&proced8&proced9&procedA&procedB&procedC&procedD&allsel=yes&selCode1&selCode2&selCode3&page_cnt=20000&fbclid=IwAR2rwV0EENp6Yf5XmXrl1CGotWEXHtw2eBb-7gGFO-PbhjJCfYN3v1q_9sc"
    elif mode == '通識':
        return "https://nol.ntu.edu.tw/nol/coursesearch/search_for_03_co.php?alltime=yes&allproced=yes&selcode=-1&coursename=&teachername=&current_sem=" + sem + "&yearcode=0&op=&startrec=0&week1=&week2=&week3=&week4=&week5=&week6=&proced0=&proced1=&proced2=&proced3=&proced4=&procedE=&proced5=&proced6=&proced7=&proced8=&proced9=&procedA=&procedB=&procedC=&procedD=&allsel=yes&selCode1=&selCode2=&selCode3=&page_cnt=20000"
    elif mode == '體育':
        return "https://nol.ntu.edu.tw/nol/coursesearch/search_for_09_gym.php?current_sem=110-1&op=S&startrec=" + \
               str(startrec) + "&cou_cname=&tea_cname=&year_code=2&checkbox=&checkbox2=&alltime=yes&allproced=yes&week1=&week2=&week3=&week4=&week5=&week6=&proced0=&proced1=&proced2=&proced3=&proced4=&procedE=&proced5=&proced6=&proced7=&proced8=&proced9=&procedA=&procedB=&procedC=&procedD="
    elif mode == '外文':
        return "https://nol.ntu.edu.tw/nol/coursesearch/search_for_01_major.php?alltime=yes&allproced=yes＆selcode=-1&coursename&teachername&couarea=7&current_sem=110-1&yearcode=0&op&startrec=0&week1&week2&week3&week4&week5&week6&proced0&proced1&proced2&proced3&proced4&procedE&proced5&proced6&proced7&proced8&proced9&procedA&procedB&procedC&procedD&allsel=yes&selCode1&selCode2&selCode3&page_cnt=20000&fbclid=IwAR1rbn6rZfQM5ETlZRqXNvMN-Wg6ChCBduyccCxOPdDfzgOMIPclNcSJmnw"
    elif mode == '共同':
        return "https://nol.ntu.edu.tw/nol/coursesearch/search_for_01_major.php?alltime=yes&allproced=yes＆selcode=-1&coursename&teachername&current_sem=110-1&yearcode=0&op&startrec=0&week1&week2&week3&week4&week5&week6&proced0&proced1&proced2&proced3&proced4&procedE&proced5&proced6&proced7&proced8&proced9&procedA&procedB&procedC&procedD&allsel=yes&selCode1&selCode2&selCode3&page_cnt=20000&fbclid=IwAR1rbn6rZfQM5ETlZRqXNvMN-Wg6ChCBduyccCxOPdDfzgOMIPclNcSJmnw"


def get_soup(url, code='big5'):
    response = requests.get(url)
    response.encoding = code
    soup = BeautifulSoup(response.text, "html.parser")
    return soup


def soup_to_df(soup):
    table_raw = soup.find('table', border="1", cellspacing="1", cellpadding="1", bordercolorlight="#CCCCCC",
                          bordercolordark="#CCCCCC")
    table_row = table_raw.select('tr')
    data = []
    for tr in table_row[1:]:
        td = tr.find_all('td')
        data.append([tr.text.strip() for tr in td])

    try:
        df = pd.DataFrame(data,
                          columns=['流水號', '授課對象', '課號', '班次', '課程名稱', '簡介影片', '學分', '課程識別碼', '全/半年', '必/選修', '授課教師',
                                   '加選方式', '時間教室', '總人數', '選課限制條件', '備註', '課程網頁', '本學期我預計要選的課程'])
    except ValueError:
        df = pd.DataFrame(data,
                          columns=['流水號', '授課對象', '課號', '班次', '課程名稱', '簡介影片', '學分', '課程識別碼', '全/半年', '授課教師', '加選方式',
                                   '時間教室', '總人數', '選課限制條件', '備註', '課程網頁', '本學期我預計要選的課程'])
    df['classroom'] = ''
    for i in range(7):
        for j in range(15):
            df['time_' + str(i) + '_' + str(j)] = 0
    for i in range(1, 9):
        df['A' + str(i)] = 0
    return df


def url_to_df(mode='系所', dpt='9010', sem='110-1', startrec=0):
    url = get_url(mode, dpt, sem, startrec)
    soup = get_soup(url)
    df = soup_to_df(soup)
    return df


if __name__ == "__main__":
    # print(get_url(mode='通識'))
    # print(url_to_df(mode='通識'))
    # print(url_to_df(mode='系所', dpt='9010'))
    print(url_to_df(mode='體育'))
    # print(url_to_df(mode='所有'))
