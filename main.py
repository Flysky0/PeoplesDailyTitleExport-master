import time
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill
import pyodbc
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm
from config import DatabaseHost, DatabasePassword, smartPassword, smartUserName
from smartLogin import SmartLogin


class Database:
    # 初始化数据库连接
    DatabaseName = 'PeopleDaily'
    DateTableName = "DateIndex"
    contentTableName = "ArticleIndex"

    def __init__(self):
        self.cursor = self.connetMsSqlServer()
        # self.cursor.execute("USE MASTER")   #Debug
        # self.cursor.execute("DROP DATABASE PeopleDaily")  # 删除旧数据库 #Debug
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
                    [ArticalIndex] NVARCHAR (MAX) NOT NULL,
                    [Date]         DATE           NOT NULL,
                    [ArticalTitle] NVARCHAR (MAX) NOT NULL,
                    [Author]       NVARCHAR (MAX) NULL,
                    [Keyword]      NVARCHAR (MAX) NULL,
                    [Page]         NVARCHAR (MAX) NULL,
                    [ArticalURL]   NVARCHAR (MAX) NOT NULL,
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

    def InsertData_Artical(self, Index, Date, Title, Author, keyword, Page, URL):
        InsertSQL = f"INSERT INTO [{self.contentTableName}] ([ArticalIndex],[Date],[ArticalTitle],[Author],[Keyword],[Page],[ArticalURL]) VALUES ('{Index}','{Date}','{Title}','{Author}','{keyword}','{Page}','{URL}')"
        try:
            self.cursor.execute(InsertSQL)
        except Exception as err:
            # 判断错误是否因主键冲突
            PRIMARYKEY_ERROR_CHECK = re.search('PRIMARY KEY', str(err))
            if PRIMARYKEY_ERROR_CHECK == None:
                raise
            else:
                print(f"主键错误：{Index}已存在")

    def SelectDate(self, date):
        DateSelectSQL = f"SELECT Date FROM [{self.DateTableName}] Where Date = '{date.replace('.','-')}'"
        DateList = self.cursor.execute(DateSelectSQL).fetchall()
        return bool(DateList)

    def ExportToXlsx():
        pass


class Xlsx():
    def __init__(self, filename):
        self.filename = filename
        self.WorkBook = openpyxl.Workbook()
        self.WorkSheet = self.WorkBook.active

    def CellStyle(self, xlsx):
        titlestyle = NamedStyle(name="TitleStyle",
                                font=Font(name='微软雅黑', size=24, bold=True, italic=False,
                                          vertAlign=None, underline='none', strike=False, color="00FFFFFF"),
                                fill=PatternFill(
                                    fill_type="solid", fgColor="000066CC"),
                                border=Border(),
                                alignment=Alignment(horizontal='center', vertical='center',
                                                    text_rotation=0, wrap_text=False, shrink_to_fit=True, indent=0),
                                number_format='General',
                                )
        oddstyle = NamedStyle(name="OddStyle",
                              font=Font(name='微软雅黑', size=12, bold=False, italic=False,
                                        vertAlign=None, underline='none', strike=False, color="00000000"),
                              fill=PatternFill(
                                  fill_type="solid", fgColor="00FFFFFF"),
                              border=Border(),
                              alignment=Alignment(horizontal='center', vertical='center',
                                                  text_rotation=0, wrap_text=False, shrink_to_fit=True, indent=0),
                              number_format='General',
                              )
        evenstyle = NamedStyle(name="EvenStyle",
                               font=Font(name='微软雅黑', size=12, bold=False, italic=False,
                                         vertAlign=None, underline='none', strike=False, color="00000000"),
                               fill=PatternFill(
                                   fill_type="solid", fgColor="00FFFFCC"),
                               border=Border(),
                               alignment=Alignment(horizontal='center', vertical='center',
                                                   text_rotation=0, wrap_text=False, shrink_to_fit=True, indent=0),
                               number_format='General',
                               )
        if not ('EvenStyle' in list(xlsx.style_names)):
            xlsx.add_named_style(evenstyle)
        if not ('TitleStyle' in list(xlsx.style_names)):
            xlsx.add_named_style(titlestyle)
        if not ('OddStyle' in list(xlsx.style_names)):
            xlsx.add_named_style(oddstyle)

    def SetCellStyle(self, Row):
        for Column in range(1, self.WorkSheet.max_column + 1):
            if Row % 2 == 0:
                self.WorkSheet[openpyxl.get_column_letter(
                    Column)+str(Row)].style = "OddStyle"
            else:
                self.WorkSheet[openpyxl.get_column_letter(
                    Column)+str(Row)].style = "EvenStyle"

    def SetColumnWidth(self, Column, Width):
        self.WorkSheet.column_dimensions[openpyxl.get_column_letter(
            Column)].width = Width


class DateList:
    PeopleDailyURL = "https://ss.zhizhen.com/s?sw=NewspaperTitle%28%E4%BA%BA%E6%B0%91%E6%97%A5%E6%8A%A5%29&nps=a5074152ece3d23811d0255ed743ac43"
    # PeopleDailyURL = "https://vpncas.ahut.edu.cn/https/77726476706e69737468656265737421e7e056d23d38614a760d87e29b5a2e/s?sw=NewspaperTitle%28%E4%BA%BA%E6%B0%91%E6%97%A5%E6%8A%A5%29&nps=a5074152ece3d23811d0255ed743ac43"
    TargetTitle = "NewspaperTitle(人民日报) _超星发现系统"
    Database = Database()

    def __init__(self) -> None:
        self.GetDateIndex()
        self.GetDay()

    def GetDateIndex(self):
        session = requests.Session()
        # session.headers = headers
        response = session.get(self.PeopleDailyURL)
        # 解析首页
        soup = BeautifulSoup(response.text, "html.parser")
        # 获取日期列表
        DateListSource = soup.find_all("input", attrs={'id': 'guidedata1'})
        DateList = []
        for subDateList in DateListSource:
            DateString = subDateList.attrs["value"]
            pattern = re.compile(r'[1-2]\d{3}\.[0-1]\d\.[0-3]\d')
            DateList.append(pattern.findall(DateString))
        self.DateList = DateList

    def GetDay(self):
        for dates_Year in tqdm(self.DateList):
            for date in tqdm(dates_Year):
                if not(self.Database.SelectDate(date)):
                    WebPage(self.Database, date)

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


class WebPage:
    PeopleDailyBaseDateURL = "https://ss.zhizhen.com/s?sw=NewspaperTitle(%E4%BA%BA%E6%B0%91%E6%97%A5%E6%8A%A5)&nps=a5074152ece3d23811d0255ed743ac43&npdate="
    ArticalNumberPattern = re.compile(r'返回<span>(.+?)</span>结果')
    TimeDelay = 2

    def __init__(self, Database, Date):
        self.Database = Database
        self.date = Date
        self.LoadFristPage()
        self.MaxPage = int(self.ArticalNumber / 15 + 0.9999)

    def LoadFristPage(self):
        # 获取日期链接
        self.FirstDateURL = self.PeopleDailyBaseDateURL + self.date
        # 获取日期页面
        dateSoup = self.requestPage(self.FirstDateURL)
        ArticalNumberContent = dateSoup.find_all(
            "div", attrs={'class': 'left'})[0]
        ArticalNumber = GetRegular(
            self.ArticalNumberPattern, ArticalNumberContent).replace(",", "")
        self.ArticalNumber = int(ArticalNumber)
        self.Database.InsertData_Date(
            self.date, self.ArticalNumber, self.FirstDateURL)
        self.GetArticalList(dateSoup)

    def GetArticalList(self, dateSoup):
        ArticalList = dateSoup.find_all(
            "div", attrs={'class': 'savelist clearfix'})
        for index, subArticalList in enumerate(ArticalList):
            Article(self.Database, index, self.date,
                    subArticalList, len(str(self.ArticalNumber)))

    def GetNextPage(self):
        if self.MaxPage > 1:
            for PageIndex in range(2, self.MaxPage + 1):
                DateURL = self.PeopleDailyBaseDateURL + self.date + \
                    f"&size=15&isort=0&x=0_476&pages={PageIndex}"
                dateSoup = self.requestPage(DateURL)
                self.GetArticalList(dateSoup)

    def requestPage(self, URL):
        while True:
            try:
                response = requests.get(URL)
                break
            except:
                pass
        time.sleep(self.TimeDelay)
        return BeautifulSoup(response.text, "html.parser")


class Article:
    TitlePattern = re.compile(
        r'<input id="favtitle\d+" type="hidden" value="(.+?)"[>|/>]')
    URLPattern = re.compile(
        r'<input id="favurl\d+" type="hidden" value="(http://ss\.zhizhen\.com/.+?)"[>|/>]')
    AuthorPattern = re.compile(
        r'<input id="favauthor\d+" type="hidden" value="(.+?)第3版:要闻"[>|/>]')
    AuthorRemainPattern = re.compile(
        r'"/>, <input id="(.+)?" type="hidden" value="')
    KeywordPattern = re.compile(
        r'<li>关键词：(.+?)</li>')
    PagePattern = re.compile(
        r'<li>出处：[\d\D]+人民日报[\d\D]+(第\d+版.*|\d+版：.*|\d+版:.*).*?</li>')

    def __init__(self, Database, index, date, subArticalList, IndexWeith):
        self.Database = Database
        self.index = index
        self.date = date
        self.subArticalList = subArticalList
        self.IndexWeith = IndexWeith
        self.InsertToDatebase()

    def InsertToDatebase(self):
        ArticalForm = self.subArticalList.find_all(
            "form")[0].find_all(
            "input")
        ArticalTitle = GetRegular(
            self.TitlePattern, ArticalForm)
        ArticalAuthor = GetRegular(
            self.AuthorPattern, ArticalForm)
        if GetRegular(self.AuthorRemainPattern, ArticalAuthor) != "NULL":
            ArticalAuthor = "NULL"
        ArticalURL = GetRegular(self.URLPattern, ArticalForm)
        Articalul = self.subArticalList.find_all("ul")[0]
        ArticalKeyword = GetRegular(
            self.KeywordPattern, Articalul).replace('<font color="Red">', "").replace('</font>', "")
        ArticalPage = GetRegular(
            self.PagePattern, Articalul).replace("\xa0", "")
        if ArticalPage[0] != "第":
            ArticalPage = "第" + ArticalPage
        ArticalIndex = self.date.replace(
            ".", "") + "_" + str((self.index + 1)).zfill(self.IndexWeith)
        self.Database.InsertData_Artical(
            ArticalIndex, self.date, ArticalTitle, ArticalAuthor, ArticalKeyword, ArticalPage, ArticalURL)


def GetRegular(pattern, text):
    result = pattern.search(str(text))
    if result:
        return result.group(1)
    else:
        return "NULL"


def main():
    DateList()


if __name__ == '__main__':
    main()
