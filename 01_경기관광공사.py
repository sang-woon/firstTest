import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


wb = Workbook() # 새 워크북 생성
ws = wb.active # 현재 활성화된 sheet 가져옴
ws.title = "NewSheet" # sheet 의 이름을 변경

new_ws = wb["NewSheet"] # Dict 형태로 Sheet 에 접근

print(wb.sheetnames) # 모든 Sheet 이름 확인

# https://dojang.io/mod/page/view.php?id=2241
for i in range(1, 50):
    print(i, "******************************************************************")
    # 메인화면 이동하기
    url = f'https://www.ggtour.or.kr/tourdb/goosuk.php?tmenu=&smenu=&stitle=&page={i}&board=71&tbl=content&tsort=2&msort=15&s_plus_code1=&s_category_code1=&s_category_code1_1=&key=#a'

    # https://www.whatismybrowser.com/detect/what-is-my-user-agent/
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36'}

    res = requests.get(url, headers=headers)
    res.raise_for_status()

    soup = BeautifulSoup(res.text, "lxml")
    with open('test.html', 'w', encoding='utf8') as f:
        f.write(soup.prettify())

    # 이미지 별 url 갖고 오기 위해서 div 찾음
    table_list = soup.select('#sub_form > div.list_wrap.info > div.gall_board > ul > li')

    # print(len(table_list))
    # #sub_form > div.list_wrap.info > div.gall_board > ul

    # print(table_list[0])


    for item in table_list:
        # print("****************************** ", item)
        detail_page_url = "https://www.ggtour.or.kr/"+item.find('a').get('href')

        res = requests.get(detail_page_url, headers=headers)
        res.raise_for_status()

        res.raise_for_status()
        soup = BeautifulSoup(res.text, "lxml")

        # 시군명 / 관광지명
        tourist_name = soup.find('h5', attrs={'class': 'data_tit'}).get_text().replace(" ", "").split('\n')

        # 관광지 구분
        tourist_category = soup.find('h3', attrs={'class': 'data_info'}).get_text().split(' ')

        # 이미지 갖고 오기
        tourist_img = soup.find('img', attrs={'class': 'image-slide'})




        try:

            # 문의전화
            tourist_tel = soup.find('li', attrs={'class': 'il-tel'}).get_text().replace(" ", ",").replace("\n", ",").split(',')

            # 이용요금/가격 il-fee
            tourist_fee = soup.find('li', attrs={'class': 'il-fee'}).get_text().replace(" ", ",").replace("\n", ",").split(',')

            # 홈페이지주소
            tourist_link = soup.find_all('li', attrs={'class': 'il-link'})[0].get_text().replace(" ", ",").replace("\n",",").split(',')

            # 소개글
            tourist_introduce = soup.find_all('li', attrs={'class': 'il-link'})[1].get_text().strip().split("\n\r\n")
            tourist_introduce = tourist_introduce[1].strip()

            ws.append([tourist_name[1].strip(), tourist_name[2], tourist_category[0], tourist_category[2], tourist_tel[2],
                       tourist_fee[2], tourist_link[2], tourist_introduce])

        except:
            tourist_tel = ""
            tourist_fee = ""
            tourist_link = ""
            tourist_introduce = ""

            ws.append(
                [tourist_name[1].strip(), tourist_name[2], tourist_category[0], tourist_category[2], tourist_tel,
                 tourist_fee, tourist_link, tourist_introduce])






# 엑셀에 저장하기

wb.save("sample.xlsx")
wb.close()

# 엑셀에 저장하기