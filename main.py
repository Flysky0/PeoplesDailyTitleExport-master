import datetime
import json
import os
import re
import warnings
from DatabaseSetting import DatabasePassword, DatabaseHost
import openpyxl
import pyodbc
import requests
from PyPDF2 import PdfFileMerger


class Database:
    # 初始化数据库连接
    DatabaseName = 'PeopleDaily'

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
        TableName = "DateIndex"
        SQL = f"""IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[{TableName}]') AND type in (N'U'))
                    CREATE TABLE [dbo].[{TableName}] (
                    [Date]           DATE           NOT NULL,
                    [ArticalNumbers] INT            NULL,
                    [DailyURL]       NVARCHAR (MAX) NULL,
                    PRIMARY KEY CLUSTERED ([Date] ASC)
                )""".replace("\n", " ")
        self.cursor.execute(SQL)
        contentTableName = "ArticleIndex"
        SQL = f"""IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[{contentTableName}]') AND type in (N'U'))
                    CREATE TABLE [dbo].[{contentTableName}] (
                    [ArticalIndex] INT            NOT NULL,
                    [Date]         DATE           NOT NULL,
                    [ArticalTitle] NVARCHAR (MAX) NOT NULL,
                    [Author]       NVARCHAR (MAX) NULL,
                    [Page]         NVARCHAR (MAX) NULL,
                    [ArticalURL]   NVARCHAR (MAX) NOT NULL,
                    [Keyword]      NVARCHAR (MAX) NULL,
                    PRIMARY KEY CLUSTERED ([ArticalIndex] ASC),
                    FOREIGN KEY ([Date]) REFERENCES [dbo].[{TableName}] ([Date])
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

    def AddData(self, Date, page, title, a):
        # Create Table
        for TableName in ["DateIndex", "DailyContents"]:
            SQL = f"""IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[{TableName}]') AND type in (N'U'))
                        CREATE TABLE [dbo].[{TableName}]
                        ([IssueID] [nvarchar](50) not null,[historyTime] datetime not null,
                        [userID] [nvarchar](50) not null,
                        [orderID] [nvarchar](50) not null , [price] float not null
                        Primary Key([IssueID],[historyTime],[userID],[orderID],[price]))""".replace("\n", " ")
            self.cursor.execute(SQL)
        # insertSQL = f"INSERT INTO [{TableName}] VALUES ('{IssueID}','{historyTime}','{userID}','{orderID}','{price}')"
        try:
            pass
            # self.cursor.execute(insertSQL)
        except Exception as err:
            # 判断错误是否因主键冲突
            PRIMARYKEY_ERROR_CHECK = re.search('PRIMARY KEY', str(err))
            if PRIMARYKEY_ERROR_CHECK == None:
                raise
        self.cursor.commit()


class Day:
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
    warnings.filterwarnings('ignore')
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
