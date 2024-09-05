import requests
from bs4 import BeautifulSoup
import re

class NaverApi:
    def __init__(self):
        self.url = 'https://finance.naver.com/sise/theme.naver'
        self.column_texts = []
        self.themedict = dict()


        self.page = self.extract_page_number()
        print(f"page 번호 : {self.page}")

    def extract_page_number(self):

        response = requests.get(self.url)
        soup = BeautifulSoup(response.content, 'html.parser')

        link_td = soup.find('td', class_='pgRR')
        if link_td:
            a_tag = link_td.find('a')
            if a_tag and 'href' in a_tag.attrs:
                href = a_tag['href']
                # 페이지 번호 추출 (정규 표현식 사용)
                match = re.search(r'page=(\d+)', href)
                if match:
                    page_number = match.group(1)
                    return page_number
        return None

    def allhavetheme(self):

        self.column_texts = []
        for i in range(1, int(self.page) +1):
        #for i in range(1, 2):
            link = f"{self.url}?&page={i}"
            self.column_texts.append(self.theme_one(link,"type_1","col_type1"))



    def theme_one(self, link, tableName, className):
        # 테이블 찾기
        response = requests.get(link)
        soup = BeautifulSoup(response.content, 'html.parser')

        table = soup.find('table', class_=tableName)

        all_theme = []
        # 테이블의 모든 행 추출
        for row in table.find_all('tr'):
            if className:
                columns = row.find_all('td', class_=className)
            else:
                columns = row.find_all('td')

            coltext = []
            for col in columns:
                # 셀 텍스트 추출

                text = col.text.strip()
                if text:  # 빈 텍스트가 아닌 경우만 추가
                    coltext.append(text)

                # 셀 내의 링크 추출
                link = col.find('a')
                if link:
                    href = link.get('href')
                    link_text = link.text.strip()
                    if link_text:
                        coltext.append(href)

            if coltext:  # 빈 리스트가 아닌 경우만 추가
                all_theme.append(coltext)
        return all_theme

    def jongmok_print(self):

        # 각 행의 데이터를 출력
        for index, data in enumerate(self.column_texts):
            print(f"no{index + 1}: {data}")
        else:
            print("No table found.")

    # 테마별 분류에서 딕트로 종목별 테마를 만들자

    def makedict(self):

        data = []
        theme_sort = self.column_texts
        basicurl = "https://finance.naver.com"


        for i in theme_sort:
            print(f"{i[0]} 처리중")
            for j in i:
                data.append([j[0],self.theme_one(basicurl + j[1],'type_5','name')])

        stock_theme_dict = {}



        # Populate the dictionary
        for theme, stocks in data:

            for stock in stocks:
                stock_name, stock_code = stock[0], stock[1]

                match = re.search(r'code=(\d+)', stock_code)
                if match:
                    stock_code = match.group(1)

                # If stock_code exists, append the theme; if not, create a new list
                if stock_code in stock_theme_dict:
                    stock_theme_dict[stock_code].append(theme)
                else:
                    stock_theme_dict[stock_code] = [theme]


        # Save stock_theme_dict to a Python file (e.g., stock_data.py)
        with open("stock_data.py", "w", encoding="utf-8") as file:
            file.write("def stockcode(code):")
            file.write("    stock = {\n")
            for stock_code, themes in stock_theme_dict.items():
                file.write(f'    "{stock_code}": {themes},\n')
            file.write("        }\n")
            file.write("    return stock.get(str(code), ['Unknown Code'])")

        print("Data saved to stock_data.py")

api = NaverApi()
api.allhavetheme()

#api.theme_one("https://finance.naver.com/sise/sise_group_detail.naver?type=theme&no=126", 'type_5',"name")
#pi.jongmok_print()
api.makedict()