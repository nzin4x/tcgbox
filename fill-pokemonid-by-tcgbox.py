import pandas
import requests
from bs4 import BeautifulSoup
import time

df = pandas.read_excel('c:\Temp\포켓몬 판매 정리 2022.xlsx', sheet_name='Sheet1')

# print whole sheet data
idRows =  df['시리즈'].astype(str) + '+' + df['카드 번호'].map('{0:03d}'.format).values

print(idRows)

for idx, id in enumerate(idRows):

    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
    res = requests.get('https://tcgbox.co.kr/product/search.html?banner_action=&keyword=' + id, headers=headers)

    list = res.text
    soup = BeautifulSoup(list, 'html.parser')

    try:
        검색카드목록 = soup.select('#contents > div.xans-element-.xans-search.xans-search-result > ul')

        founds = BeautifulSoup(str(검색카드목록[0]), 'html.parser')

        cards = founds.select('li.item')

        cardid = id.replace("+", " ")

        for card in cards: 
            if cardid.upper() in str(card).upper():
                nameEle = BeautifulSoup(str(card), 'html.parser').select('div > p.name > strong > a > span:last-child')
                linkEle = BeautifulSoup(str(card), 'html.parser').select('div > p.name > strong > a ')
                priceEle = BeautifulSoup(str(card), 'html.parser').select('div > ul> li:first-child > span')

                name = (nameEle[0].next)
                link = 'https://tcgbox.co.kr' + (linkEle[0]['href'])
                price = (priceEle[0].next)

                # 검색 결과를 dataframe 에 저장
                df.at[idx, '카드 이름'] = name
                df.at[idx, '링크'] = link
                df.at[idx, '가격'] = price
            else:
                card
    except:
        print('순서' + str(idx) + '이상 ' )

    time.sleep(0.5)

# dataframe 을 엑셀에 저장
print(df)
df.to_excel('c:\Temp\포켓몬 판매 정리 2022 결과.xlsx', sheet_name='Sheet1')
