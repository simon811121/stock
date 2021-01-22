from bs4 import BeautifulSoup
import requests
from requests.exceptions import ConnectionError
from requests.exceptions import ReadTimeout
import time
import datetime
from datetime import date, timedelta
import pandas as pd
from openpyxl import load_workbook
import xlsxwriter
import numpy as np

# ------------------------------
#    assert funtion
# ------------------------------
def assertFunc(condition, txt, place):
    if not condition:
        print(txt + str(place))
        while(1):
            time.sleep(1)
#----------------------------------------------------------------------
#    parsing data from web
#----------------------------------------------------------------------
# ------------------------------
# <INPUT>
#        url:        for parsing
#        chkRange:  total rank need to get
# <OUTPUT>
#        DataFrame
# ------------------------------
def parseLeaderBoard(url, chkRange):
    # delay
    time.sleep(5)
    
    # web header
    headers = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}

    # get request
    i = 3
    while i > 0:
        try:
            resp = requests.get(url, headers=headers)
            break
        except (ConnectionError, ReadTimeout) as error:
            print(error)
            print('retry one more time after 60s', i, 'times left')
            time.sleep(60)
        i -= 1
    
    if i <= 0:
        return None

    # change html encoding
    resp.encoding = 'utf8'

    # get html data
    soup = BeautifulSoup(resp.text, 'html5lib')
    # --website layout--
    # txtStockListData -> divStockList -> tblStockList
    table = soup.find("table", {"id":"tblStockList"})

    # parsing titel from leaderborad table
    rows = table.find("tr")
    titles = []
    for row in rows:
        title = []
        title = row.find("div")
        col = title.get_text()
        titles.append(col)

    # parsing row data from leaderboard table
    rows = table.find_all("tr", id=lambda x: x and x.startswith('row'))
    parCols = []
    for row in rows:
        stockInfo = []
        stockInfo = row.find_all("td")
        cols = []
        for i in range(0, len(stockInfo)):
            col = stockInfo[i].get_text()
            cols.append(col)
        parCols.append(cols)

    # form DataFrame data from parsing result data
    df = pd.DataFrame.from_dict(parCols)
    df.columns = titles
    df.index = df['排名']
    del df['排名']
    
    return df[0:chkRange]  # 排名 1 ~ chkRange

#----------------------------------------------------------------------
#    save data
#----------------------------------------------------------------------
# ------------------------------
# <INPUT>
#        saveDf:    dataframe to save
#        startCol:  excel column offset to save
#        startRow:  excel row offset to save
#        txt:       title txt to save
# <OUTPUT>
#        saved excel in directory
# ------------------------------
def saveToExcel(saveDf, startCol, startRow, txt):
    if len(saveDf) == 0:
        print(txt + ' has no data')
        return None
    try:
        dateString = saveDf['法人買賣日期']
        dateString1 = dateString.values[0].split("/")
    except Exception as error:
        print(error)
        assertFunc(0, 'error code logic', 1)

    year = str(date.today().year)
    if date.today().month < 10:
        month = ' 0' + str(date.today().month)
    else:
        month = ' ' + str(date.today().month)

    outputFileName = year + month + '月 三大法人買賣超 紀錄.xlsx'
    sheetName = dateString1[0] + '-' + dateString1[1]
    titleRow = 1 if ((startRow - 1) <= 0) else (startRow - 1)
    titleCol = 1 if ((startCol - 1) <= 0) else (startCol - 1)

    try:
        outputFile = load_workbook(outputFileName)
        writer = pd.ExcelWriter(outputFileName, engine='openpyxl')  # pylint: disable=abstract-class-instantiated
        writer.book = outputFile
        writer.sheets = dict((ws.title, ws) for ws in outputFile.worksheets)
        saveDf.to_excel(writer, sheet_name=sheetName, startcol=startCol, startrow=startRow)
        worksheet = writer.sheets[sheetName]
        worksheet.cell(row=titleRow, column=titleCol).value = txt
        writer.save()
    except Exception as error: 
        print(error)
        saveDf.to_excel(outputFileName, sheet_name=sheetName, startcol=startCol, startrow=startRow)
        outputFile = load_workbook(outputFileName)
        worksheet = outputFile.get_sheet_by_name(sheetName)
        worksheet.cell(row=titleRow, column=titleCol).value = txt
        outputFile.save(outputFileName)

#----------------------------------------------------------------------
#    compare data
#----------------------------------------------------------------------
# ------------------------------
# <INPUT>
#        df1:    dataframe 1st to compare
#        df2:    dataframe 2nd to compare
#        startCol:  excel column offset to save
#        startRow:  excel row offset to save
#        txt:       title txt to save
# <OUTPUT>
#        result df
# --------------------------------------
# 這個 function 是用來取得剛爬完蟲的資料
# --------------------------------------
def getRepeatStockIdDf(df1, df2, startCol, startRow, txt):
    rsltDf = pd.merge(df1, df2, on=['代號'], how='inner')
    if len(rsltDf) == 0: # no repeate stock id between two rank list
        return rsltDf
    rsltDf = rsltDf[['代號', '名稱_x', '法人買賣日期_x']]
    rsltDf.rename(columns={'名稱_x' : '名稱', '法人買賣日期_x':'法人買賣日期'}, inplace=True)
    rsltDf.index = np.arange(1, len(rsltDf) + 1)
    if len(rsltDf) != 0:
        saveToExcel(rsltDf, startCol, startRow, txt)
    else:
        try:
            dateString = df1['法人買賣日期'][1]
            dateString1 = dateString.split("/")
        except Exception as error:
            print(error)
            assertFunc(0, 'error code logic', 2)

        year = str(date.today().year)
        if date.today().month < 10:
            month = ' 0' + str(date.today().month)
        else:
            month = ' ' + str(date.today().month)

        outputFileName = year + month + '月 三大法人買賣超 紀錄.xlsx'
        sheetName = dateString1[0] + '-' + dateString1[1]
        titleRow = 1 if ((startRow - 1) <= 0) else (startRow - 1)
        try:
            outputFile = load_workbook(outputFileName)
            worksheet = outputFile.get_sheet_by_name(sheetName)
            worksheet.cell(row=titleRow, column=startCol).value = txt
            outputFile.save(outputFileName)
        except Exception as error:
            print(error)
            assertFunc(0, 'error code logic', 3)
    return rsltDf
# ------------------------------
# <INPUT>
#        df1:    dataframe 1st to compare
#        df2:    dataframe 2nd to compare
# <OUTPUT>
#        result df
# ---------------------------------------------------
# 這個 function 是用來比較今日爬完結果，及昨日爬完結果
# ---------------------------------------------------
def getRepeatStockRankDf(df1, df2):
    if len(df1) == 0 or len(df2) == 0:
        return None
    try:        
        rsltDf = pd.merge(df1, df2, on=['代號', '名稱'], how='inner')
    except:
        pass

    return rsltDf

#----------------------------------------------------------------------
#    get yesterday's rank overlap detail
#----------------------------------------------------------------------
# ------------------------------
# <INPUT>
#        srcDf:    dataframe from yesterday's data
# <OUTPUT>
#        success:  get '代號' or not
#        index:    that contain '代號'
# ------------------------------
def splitDfByDiffOverlap(srcDf, startOfst):
    chkIdx = 0
    for index, row in srcDf.iloc[startOfst:].iterrows():
        if row.values[0] == '代號':
            return True, index
        chkIdx = index
    if chkIdx == len(srcDf):
        return False, index

#----------------------------------------------------------------------
#    code main
#----------------------------------------------------------------------
print(datetime.datetime.now())
LEADERBOARD_MAX_RANK_CHK = 30

# f = foreign = 外資
# i = invset trust = 投信
# d = dealer = 自營商
# 外資單日
url = 'https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E5%A4%96%E8%B3%87%E7%B4%AF%E8%A8%88%E8%B2%B7%E8%B6%85%E5%BC%B5%E6%95%B8+%E2%80%93+%E7%95%B6%E6%97%A5%40%40%E5%A4%96%E8%B3%87%E7%B4%AF%E8%A8%88%E8%B2%B7%E8%B6%85%40%40%E5%A4%96%E8%B3%87%E8%B2%B7%E8%B6%85%E5%BC%B5%E6%95%B8+%E2%80%93+%E7%95%B6%E6%97%A5'
fRsltDf = parseLeaderBoard(url, LEADERBOARD_MAX_RANK_CHK)
saveToExcel(fRsltDf, 0, 1, '外資 單日')

# 投信單日
url = 'https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E6%8A%95%E4%BF%A1%E7%B4%AF%E8%A8%88%E8%B2%B7%E8%B6%85%E5%BC%B5%E6%95%B8+%E2%80%93+%E7%95%B6%E6%97%A5%40%40%E6%8A%95%E4%BF%A1%E7%B4%AF%E8%A8%88%E8%B2%B7%E8%B6%85%40%40%E6%8A%95%E4%BF%A1%E8%B2%B7%E8%B6%85%E5%BC%B5%E6%95%B8+%E2%80%93+%E7%95%B6%E6%97%A5'
iRsltDf = parseLeaderBoard(url, LEADERBOARD_MAX_RANK_CHK)
saveToExcel(iRsltDf, 0, (LEADERBOARD_MAX_RANK_CHK + 5), '投信 單日')

# 自營商單日
url = 'https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E8%87%AA%E7%87%9F%E5%95%86%E7%B4%AF%E8%A8%88%E8%B2%B7%E8%B6%85%E5%BC%B5%E6%95%B8+%E2%80%93+%E7%95%B6%E6%97%A5%40%40%E8%87%AA%E7%87%9F%E5%95%86%E7%B4%AF%E8%A8%88%E8%B2%B7%E8%B6%85%40%40%E8%87%AA%E7%87%9F%E5%95%86%E8%B2%B7%E8%B6%85%E5%BC%B5%E6%95%B8+%E2%80%93+%E7%95%B6%E6%97%A5'
dRsltDf = parseLeaderBoard(url, LEADERBOARD_MAX_RANK_CHK)
saveToExcel(dRsltDf, 0, ((LEADERBOARD_MAX_RANK_CHK + 5) * 2), '自營商 單日')

LEADERBOARD_OVERLAP_DATA_COL_OFST = len(fRsltDf.columns) + 4

# 外資 & 投信 單日
fileRowOfst = 1
f_i_RsltDf = getRepeatStockIdDf(fRsltDf, iRsltDf, LEADERBOARD_OVERLAP_DATA_COL_OFST, fileRowOfst, '外資 & 投信 單日')

# 投信 & 自營商 單日
fileRowOfst = fileRowOfst + len(f_i_RsltDf) + 4
i_d_RsltDf = getRepeatStockIdDf(iRsltDf, dRsltDf, LEADERBOARD_OVERLAP_DATA_COL_OFST, fileRowOfst, '投信 & 自營商 單日')

# 外資 & 自營商 單日
fileRowOfst = fileRowOfst + len(i_d_RsltDf) + 4
f_d_RsltDf = getRepeatStockIdDf(fRsltDf, dRsltDf, LEADERBOARD_OVERLAP_DATA_COL_OFST, fileRowOfst, '外資 & 自營商 單日')

# 外資 & 投信 & 自營商 單日
fileRowOfst = fileRowOfst + len(f_d_RsltDf) + 4
f_i_d_RsltDf = getRepeatStockIdDf(f_i_RsltDf, i_d_RsltDf, LEADERBOARD_OVERLAP_DATA_COL_OFST, fileRowOfst, '外資 & 投信 & 自營商 單日')

# 設定國定假日
holidays_array_valid = [1, # 2020
                        1, # 2021
                        0, # 2022
                        0, # 2023
                        0, # 2024
                        0] # 2025
holidays_in_2020_month = [10, 10, 10, 1] # month
holidays_in_2020_day   = [ 9,  2,  1, 1] # day, # 必須從後面的往回填
holidays_in_2021_month = [12, 10,  9,  9,  6,  4, 4, 4,  2,  2,  2,  2,  2, 1] # month
holidays_in_2021_day   = [31, 11, 21, 20, 14, 30, 5, 2, 16, 15, 12, 11, 10, 1] # day, # 必須從後面的往回填
holidays_len = [4, # 2020
                14, # 2021
                0, # 2022
                0, # 2023
                0, # 2024
                0] # 2025

yesterday = date.today() - timedelta(days=1)
if not holidays_array_valid[yesterday.year - 2020]:
    print('please fill holidays in ' + str(yesterday.year))
    assertFunc(0, 'error code logic', 4)
if not holidays_len[yesterday.year - 2020]:
    print('please fill holidays length in ' + str(yesterday.year))
    assertFunc(0, 'error code logic', 5)

# 取得昨天的排行，並得到任兩大法人連續
while (yesterday.weekday() >= 5):  # 只挑 1 ~ 5
    yesterday = yesterday - timedelta(days=1)
    i = 0
    while i < holidays_len[yesterday.year - 2020]:
        holidays = date(yesterday.year, holidays_in_2021_month[i], holidays_in_2021_day[i])
        if yesterday == holidays:
            yesterday = yesterday - timedelta(days=1)
        if yesterday.month > holidays_in_2021_month[i]:
            break
        i += 1
        

if yesterday.month < 10:
    excel_name = str(yesterday.year) + ' 0' + str(yesterday.month) + '月 三大法人買賣超 紀錄.xlsx'
else:
    excel_name = str(yesterday.year) + ' ' + str(yesterday.month) + '月 三大法人買賣超 紀錄.xlsx'

month_str = '0' if yesterday.month < 10 else ''
if yesterday.day < 10:
    sheet_name = month_str + str(yesterday.month) + '-0' + str(yesterday.day)
else:
    sheet_name = month_str + str(yesterday.month) + '-' + str(yesterday.day)

fileColOfst = LEADERBOARD_OVERLAP_DATA_COL_OFST + 1
try:
    y_rank_df = pd.read_excel(excel_name, sheet_name=sheet_name, header=1, usecols=[fileColOfst, fileColOfst + 1])  # y = yesterday; [fileColOfst, fileColOfst + 1] = 代號, 名稱
    y_rank_df = y_rank_df.dropna()
    y_rank_df.index = np.arange(1, len(y_rank_df) + 1)
    
    index = 0
    rsltIndex = []
    parseResult, index = splitDfByDiffOverlap(y_rank_df, index)
    while parseResult:
        rsltIndex.append(index)
        parseResult, index = splitDfByDiffOverlap(y_rank_df, index)    

    y_rank_df = y_rank_df.drop(rsltIndex)
    y_rank_df = y_rank_df.drop_duplicates()
    y_rank_df = y_rank_df.rename(columns={'代號.1' : '代號', '名稱.1' : '名稱'})

    # 用昨天的排行榜，計算連續多日排行榜
    df1 = getRepeatStockRankDf(y_rank_df, f_i_RsltDf)
    df2 = getRepeatStockRankDf(y_rank_df, i_d_RsltDf)
    df3 = getRepeatStockRankDf(y_rank_df, f_d_RsltDf)
    df4 = getRepeatStockRankDf(y_rank_df, f_i_d_RsltDf)
    tmpDf = [df1, df2, df3, df4]
    c_rank_df = pd.concat(tmpDf)
    c_rank_df = c_rank_df.drop_duplicates()
    c_rank_df.index = np.arange(1, len(c_rank_df) + 1)

    fileColOfst = LEADERBOARD_OVERLAP_DATA_COL_OFST + len(f_i_RsltDf.columns) + 4
    fileRowOfst = 1
    saveToExcel(c_rank_df, fileColOfst, fileRowOfst, '任兩大法人連續')
except Exception as error:
    print(error)
    assertFunc(0, 'error code logic', 6)

# 取得要輸入的股票
tmpDf = [f_i_RsltDf, i_d_RsltDf, f_d_RsltDf]
s_rank_df = pd.concat(tmpDf)
tmpDf = [s_rank_df, f_i_d_RsltDf, c_rank_df]
s_rank_df = pd.concat(tmpDf)
s_rank_df = s_rank_df.drop_duplicates(keep=False)
s_rank_df.index = np.arange(1, len(s_rank_df) + 1)
fileRowOfst = fileRowOfst + len(c_rank_df) + 4
saveToExcel(s_rank_df, fileColOfst, fileRowOfst, '單日該輸入的股票')

print("code end")