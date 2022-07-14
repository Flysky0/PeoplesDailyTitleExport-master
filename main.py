import datetime
import json
import os
import re
import warnings

import openpyxl
import pyodbc
import requests
from bs4 import BeautifulSoup
from PyPDF2 import PdfFileMerger

from config import DatabaseHost, DatabasePassword, smartPassword, smartUserName
from smartLogin import SmartLogin


class Database:
    # 初始化数据库连接
    DatabaseName = 'PeopleDaily'
    DateTableName = "DateIndex"
    contentTableName = "ArticleIndex"

    def __init__(self):
        self.cursor = self.connetMsSqlServer()
        self.CreateDatabase()
        self.CreateTable()

    def CreateDatabase(self):
        self.cursor.execute(
            f"select * From master.dbo.sysdatabases where name= '{self.DatabaseName}'")
        if self.cursor.fetchone() == None:
            self.cursor.execute(
                f"Create Database [{self.DatabaseName}]")
        self.cursor.execute(f"use [{self.DatabaseName}]")

    def CreateTable(self):
        DateTableName = self.DateTableName
        contentTableName = self.contentTableName
        SQL = f"""IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[{DateTableName}]') AND type in (N'U'))
                    CREATE TABLE [dbo].[{DateTableName}] (
                    [Date]           DATE           NOT NULL,
                    [ArticalNumbers] INT            NULL,
                    [DailyURL]       NVARCHAR (MAX) NULL,
                    PRIMARY KEY CLUSTERED ([Date] ASC)
                )""".replace("\n", " ")
        self.cursor.execute(SQL)
        SQL = f"""IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[{contentTableName}]') AND type in (N'U'))
                    CREATE TABLE [dbo].[{contentTableName}] (
                    [ArticalIndex] INT            NOT NULL,
                    [Date]         DATE           NOT NULL,
                    [ArticalTitle] NVARCHAR (MAX) NOT NULL,
                    [Author]       NVARCHAR (MAX) NULL,
                    [Keyword]      NVARCHAR (MAX) NULL,
                    [Page]         NVARCHAR (MAX) NULL,
                    [ArticalURL]   NVARCHAR (MAX) NOT NULL,
                    PRIMARY KEY CLUSTERED ([ArticalIndex] ASC),
                    FOREIGN KEY ([Date]) REFERENCES [dbo].[{DateTableName}] ([Date])
                );""".replace("\n", " ")
        self.cursor.execute(SQL)

    def connetMsSqlServer(self):
        server = DatabaseHost
        username = 'sa'
        password = DatabasePassword
        DB = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +
                            server + ';UID=' + username + ';PWD=' + password, autocommit=True)
        cursor = DB.cursor()
        return cursor

    def InsertData_Date(self, Date, ArticalNumbers, DailyURL):
        if DailyURL != None:
            DailyURL = f"'{DailyURL}'"
        InsertSQL = f"INSERT INTO [{self.DateTableName}] VALUES ('{Date}',{ArticalNumbers},{DailyURL})"
        try:
            self.cursor.execute(InsertSQL)
        except Exception as err:
            # 判断错误是否因主键冲突
            PRIMARYKEY_ERROR_CHECK = re.search('PRIMARY KEY', str(err))
            if PRIMARYKEY_ERROR_CHECK == None:
                raise
            else:
                print(f"主键错误：{Date}已存在")
        # self.cursor.commit()

    def InsertData_Artical(self, Index, Date, Title, Author, keyword, Page, URL):
        InsertSQL = f"INSERT INTO [{self.DateTableName}] VALUES ({Index},'{Date}','{Title}','{Author}','{keyword}','{Page}','{URL}')"
        try:
            self.cursor.execute(InsertSQL)
        except Exception as err:
            # 判断错误是否因主键冲突
            PRIMARYKEY_ERROR_CHECK = re.search('PRIMARY KEY', str(err))
            if PRIMARYKEY_ERROR_CHECK == None:
                raise
            else:
                print(f"主键错误：{Date}已存在")


class Day:
    PeopleDailyURL = "https://ss.zhizhen.com/s?sw=NewspaperTitle%28%E4%BA%BA%E6%B0%91%E6%97%A5%E6%8A%A5%29&nps=a5074152ece3d23811d0255ed743ac43"
    # PeopleDailyURL = "https://vpncas.ahut.edu.cn/https/77726476706e69737468656265737421e7e056d23d38614a760d87e29b5a2e/s?sw=NewspaperTitle%28%E4%BA%BA%E6%B0%91%E6%97%A5%E6%8A%A5%29&nps=a5074152ece3d23811d0255ed743ac43"
    TargetTitle = "NewspaperTitle(人民日报) _超星发现系统"
    Database = Database()

    def GetDateIndex(self):
        session = requests.Session()
        # session.headers = headers
        response = session.get(self.PeopleDailyURL)

        # 解析首页
        soup = BeautifulSoup(response.text, "html.parser")
        # 获取日期列表
        DateList = soup.find_all("input", attrs={'id': 'guidedata1'})
        for subDateList in DateList:
            DateString = subDateList.attrs["value"]
            pattern = re.compile(r'[1-2]\d{3}\.[0-1]\d\.[0-3]\d')
            DateListProcessed = pattern.findall(DateString)
            for date in DateListProcessed:
                Database.InsertData_Date(self.Database, date, "NULL", "NULL")
        # 获取日期列表链接
        DateListURL = [self.PeopleDailyURL + i["href"] for i in DateList]
        # 获取日期列表标题
        DateListTitle = [i.text for i in DateList]
        # 获取日期列表数量
        DateListNumbers = [i.text.split(" ")[0] for i in DateList]
        # 获取日期列表链接
        DateListURL = [self.PeopleDailyURL + i["href"] for i in DateList]
        return DateListTitle, DateListNumbers, DateListURL

    def GetDateIndex_VPN(self):
        cookies = SmartLogin(self.PeopleDailyURL, smartUserName,
                             smartPassword, self.TargetTitle)
        # 获取首页
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
            'Connection': 'keep-alive',
            # 'Cookie': cookies,
            'Host': 'vpncas.ahut.edu.cn',
            'Referer': 'https://vpncas.ahut.edu.cn/https/77726476706e69737468656265737421f3f652d226387d44300d8db9d6562d/cas/login?service=https://vpncas.ahut.edu.cn/login?cas_login=true',
            'sec-ch-ua': '" Not;A Brand";v="99", "Microsoft Edge";v="103", "Chromium";v="103"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-Fetch-Dest': 'document',
            'sec-Fetch-Mode': 'navigate',
            'sec-Fetch-Site': 'same-origin',
            'sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.114 Safari/537.36 Edg/103.0.1264.49'
        }
        session = requests.Session()
        # session.headers = headers
        response = session.get(self.PeopleDailyURL,
                               headers=headers, cookies=cookies)

        # 解析首页
        soup = BeautifulSoup(response.text, "html.parser")
        # 获取日期列表
        DateList = soup.find_all("a", class_="news_title")
        # 获取日期列表链接
        DateListURL = [self.PeopleDailyURL + i["href"] for i in DateList]
        # 获取日期列表标题
        DateListTitle = [i.text for i in DateList]
        # 获取日期列表数量
        DateListNumbers = [i.text.split(" ")[0] for i in DateList]
        # 获取日期列表链接
        DateListURL = [self.PeopleDailyURL + i["href"] for i in DateList]
        return DateListTitle, DateListNumbers, DateListURL

    NOW = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
    YEAR = str(NOW.year).zfill(4)
    MONTH = str(NOW.month).zfill(2)
    DAY = str(NOW.day).zfill(2)
    DATE = ''.join([YEAR, MONTH, DAY])

    DIR = os.path.join('Download', DATE)
    MERGED_DIR = os.path.join('Download', 'MERGED')
    # PAGES_FILE_PATH = os.path.join(DIR, f'{DATE}.zip')
    MERGED_FILE_PATH = os.path.join(MERGED_DIR, f'{DATE}.pdf')

    HOME_URL = f'http://paper.people.com.cn/rmrb/html/{YEAR}-{MONTH}/{DAY}/nbs.D110000renmrb_01.htm'
    PAGE_COUNT = requests.get(HOME_URL).text.count('pageLink')

    if not os.path.isdir(DIR):
        os.makedirs(DIR)
    if not os.path.isdir(MERGED_DIR):
        os.makedirs(MERGED_DIR)


class Page:
    def __init__(self, page: str):
        self.page = page
        self.html_url = f'http://paper.people.com.cn/rmrb/html/{Day.YEAR}-{Day.MONTH}/{Day.DAY}/nbs.D110000renmrb_{self.page}.htm'
        self.html = requests.get(self.html_url).text
        self.pdf = requests.get(
            (
                'http://paper.people.com.cn/rmrb/images/{0}-{1}/{2}/{3}/rmrb{0}{1}{2}{3}.pdf'
                .format(Day.YEAR, Day.MONTH, Day.DAY, page)
            )
        ).content
        self.path = os.path.join(Day.DIR, f'{self.page}.pdf')

    def __str__(self) -> str:
        return (
            f'{self.__class__.__name__}'
            f'[date={Day.DATE}, page={self.page}, title={self.title}]'
        )

    def __repr__(self) -> str:
        return self.__str__()

    @property
    def title(self):
        return re.findall('<p class="left ban">(.*?)</p>', self.html)[0]

    @property
    def articles(self):
        return [
            (
                (
                    'http://paper.people.com.cn/rmrb/html/{}-{}/{}/{}'
                    .format(Day.YEAR, Day.MONTH, Day.DAY, i[0])
                ),
                i[1].strip()
            ) for i in
            re.findall('<a href=(nw.*?)>(.*?)</a>', self.html)
        ]

    def save_pdf(self):
        with open(self.path, 'wb') as f:
            f.write(self.pdf)


def main():
    # warnings.filterwarnings('ignore')
    Day().GetDateIndex()
    pages = [Page(str(i + 1).zfill(2)) for i in range(Day.PAGE_COUNT)]
    # pages_file = zipfile.ZipFile(Day.PAGES_FILE_PATH, 'w') #建立压缩包
    merged_file = PdfFileMerger(False)
    data = {
        'date': Day.DATE,
        'page_count': str(Day.PAGE_COUNT),
        # 'pages_file_path': Day.PAGES_FILE_PATH,
        'merged_file_path': Day.MERGED_FILE_PATH,
        'release_body': (
            f'# [{Day.DATE}]({Day.HOME_URL})'
            f'\n\n今日 {Day.PAGE_COUNT} 版'
        )
    }

    # Process
    for page in pages:
        # Save pdf
        page.save_pdf()

        # Pages file
        # pages_file.write(page.path, os.path.basename(page.path))

        # Merged file
        merged_file.append(page.path)

        # Data，版面、版面URL
        data['release_body'] += f'\n\n## [{page.title}]({page.html_url})\n'
        for article in page.articles:
            # URL、标题
            data['release_body'] += f'\n- [{article[1]}]({article[0]})'

        # Info
        print(f'Processed {page}')

    # Save
    # pages_file.close()
    merged_file.write(Day.MERGED_FILE_PATH)
    merged_file.close()
    # for page in pages:
    #     os.remove(page.path)
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


if __name__ == '__main__':
    main()
