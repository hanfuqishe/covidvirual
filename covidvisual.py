#coding:utf-8

import requests
import json
import tempfile
import os, sys, getopt
import shutil
import xlsxwriter
import webbrowser
import time
import codecs
from datetime import date

SmoothDays      = 1
AutoOpenChina   = False
AutoOpenGlobal  = False
Debugging       = False
ToSaveJson      = False

TempDir=os.path.join(tempfile.gettempdir(), 'covid')

proxy=""  # this can override environ settings.
#proxy="http://localhost:8081"
os.environ["http_proxy"]  = proxy
os.environ["https_proxy"] = proxy

def ClearTempDir():
    print("Clean previous temp files ... ", end='')
    if os.path.exists(TempDir):
        shutil.rmtree(TempDir)

    os.makedirs(TempDir)

    print('OK.')


def SaveJson(JsonDict, asFile):
    if ToSaveJson: 
        print('saving json ... ', end='')
        f = codecs.open(os.path.join(TempDir, asFile + '.json'), "w", 'utf-8')
        json.dump(JsonDict, f, ensure_ascii=False)
        f.close()

def FetchCountryData(ID, Name, WorkBook):

    QueryCountryFmtStr='https://i.snssdk.com/forum/ncov_data/?country_id=["%s"]&country_name=%s&click_from=overseas_epidemic_tab_list&data_type=[5,4]&policy_scene=%s&src_type=country'
    Url = QueryCountryFmtStr%(ID, Name, ID)

    print('Fetching ... ', end='', flush=True)

    r = requests.get(Url)
    if r.status_code == 200:
        text = r.json()
        ParseJSONrecursively(text)
    
        Countries = text["country_data"] # countries
        for _, Country in Countries.items(): # Normally there is only 1 item in this dict
            # Country: dict_keys(['id', 'code', 'name', 'nationalFlag', 'countryTotal', 'countryIncr', 'series', 'provinces', 'isTreatingNumClear', 'confirmedPerMil', 'continent', 'updateTime'])
            # print('found [%s] ... '%(Country['name']), end='')
            SaveJson(text, Country['name'])  # for research purpose
            AddToSheet(Series=Country['series'], WorkBook=WorkBook, SheetName=Country['name'])
            break  # Only dump first item. Should has only 1 item :-)

    else: # if r.status_code != 200
        print("Fetch data error for country:", ID, Name)


def AddSeries(chart, name, col, LineCount, Y2Axis=False):
    if Y2Axis :
        dash_type = 'round_dot'
    else:
        dash_type = 'solid'

    chart.add_series({
        'name'      : '=%s!$%s$1'       %(name, col),
        'categories': '=%s!$A$2:$A$%d'  %(name, LineCount),
        'values'    : '=%s!$%s$2:$%s$%d'%(name, col, col, LineCount),
        'line'      : {
            'dash_type': dash_type, 
            'width': 2
            },
        'marker'    : {
            #'type':'automatic', 
            'type':'none', 
            'size': 3, 
            },
        'y2_axis':    Y2Axis,
    })

def CreateWorkbook(XlsxName):
    WorkBook = xlsxwriter.Workbook(XlsxName , {
        'strings_to_numbers':  True,
        'strings_to_formulas': True,
        'default_date_format': 'yyyy-mm-dd',
        'in_memory': True
    })

    return WorkBook

def AddToSheet(Series, WorkBook, SheetName):

    # SortedSeries = sorted(Series, key = lambda e:(e.__getitem__('date'))) # list of dict. Here are the detailed data we wanted

    print('add Sheet:[%s]!%s ... '%(WorkBook.filename, SheetName), end='', flush=True)

    AsDate        = WorkBook.add_format({'font_name': 'calibri', 'num_format': 'yyyy-mm-dd'})
    AsNumber      = WorkBook.add_format({'font_name': 'calibri', 'num_format': '#,##0_ '})
    AsBoldNumber  = WorkBook.add_format({'font_name': 'calibri', 'font_color': '#D00000', 'bold': False, 'num_format': '#,##0_ '})
    AsPercent     = WorkBook.add_format({'font_name': 'calibri', 'num_format': '0.0%'})

    WorkSheet = WorkBook.add_worksheet(SheetName)

    WorkSheet.set_column('A:H', 10)
    
    DestRow = 0;  Yestoday = ()

    for Item in reversed(Series):
        Today = (Item['confirmedNum'], Item['deathsNum'], Item['curesNum'], Item['treatingNum'])
        #if (Today != Yestoday) | (Today[3] != 0):
        if  True :

            Yestoday = Today

            DestRow += 1
            XlsCurrentRow = DestRow + 1  # xlsxwriter count from 0, while Excel count from 1. Add 1 to represent current line while writing as content
            XlsAboveRow   = XlsCurrentRow - SmoothDays

            Col  = 0; WorkSheet.write_datetime(DestRow, Col, date.fromisoformat(Item['date']), AsDate)                                          # A - Date
            Col += 1; WorkSheet.write_row(DestRow, Col, Yestoday, AsNumber)                                                                     # B - Confirmed Number,  C - Death Number, D - Cured Number, E - Treating Number
            Col += 4; 
            if XlsAboveRow > 0:  
                WorkSheet.write(DestRow, Col, '=IFERROR((B%d-B%d)/%d, "")'   % (XlsCurrentRow, XlsAboveRow, SmoothDays), AsBoldNumber)          # F - Daily Confirmed
            Col += 1; 
            if XlsAboveRow > 0:  
                WorkSheet.write(DestRow, Col, '=IFERROR((C%d-C%d)/%d, "")'   % (XlsCurrentRow, XlsAboveRow, SmoothDays), AsNumber)              # G - Daily Death
            Col += 1; 
            if XlsAboveRow > 0:  
                WorkSheet.write(DestRow, Col, '=IFERROR((D%d-D%d)/%d, "")'   % (XlsCurrentRow, XlsAboveRow, SmoothDays), AsNumber)              # H - Daily Cured
            Col += 1; WorkSheet.write(DestRow, Col, '=IFERROR(G%d/(G%d+H%d), "")'% (XlsCurrentRow, XlsCurrentRow, XlsCurrentRow), AsPercent)    # I - Daily mortality
            Col += 1; WorkSheet.write(DestRow, Col, '=IFERROR(C%d/(C%d+D%d), "")'% (XlsCurrentRow, XlsCurrentRow, XlsCurrentRow), AsPercent)    # J - Overall mortality
            Col += 1; WorkSheet.write(DestRow, Col, '=IFERROR(C%d/B%d, "")'      % (XlsCurrentRow, XlsCurrentRow), AsPercent)                   # K - Mortality by Media

    XlsMaxRow = DestRow + 1

    WorkSheet.add_table(
        "A1:K%d"%(XlsMaxRow),  {
            # 'header_row': True, 'autofilter': False, 'name':country + '表', 'style': 'TableStyleMedium5',
            'header_row': True, 'autofilter': False,  'style': 'TableStyleMedium5',
            'columns': [
                    {'header': '日期'},                                 # A
                    {'header': '累计确诊'},                             # B
                    {'header': '死亡'},                                 # C
                    {'header': '治愈'},                                 # D
                    {'header': '现有确诊\nTreating'},                   # E
                    {'header': '%d日平均新增确诊\nNewly diagnosed'%(SmoothDays)},   # F
                    {'header': '%d日平均新增死亡\nNewly Death'%(SmoothDays)},       # G
                    {'header': '%d日平均新增治愈\nNewly Cured'%(SmoothDays)},       # H
                    {'header': '%d日平均死亡率\nDaily mortality'%(SmoothDays)},     # I
                    {'header': '总体死亡率\nOverall mortality'},        # J 
                    {'header': '死亡率by媒体\nMortality by Media'},     # k 
                ]
        }
    )

    Chart = WorkBook.add_chart({'type': 'line'})
    Chart.set_style(2)
    Chart.set_title  ({'name':SheetName + '的COVID-19疫情'})
    Chart.set_legend ({'position': 'bottom', 'font': {'size': 8}})
    Chart.set_y_axis ({'log_base':10})
    Chart.set_y2_axis({'num_format': '0%', 'min':0, 'max':1})

    print('add chart ... ', end='', flush=True)
    for ColIndex in [ord('B')] + list(range(ord('E'), ord('K'))): 
        # print(chr(ColIndex), end=".")
        AddSeries(Chart, SheetName, chr(ColIndex), XlsMaxRow, Y2Axis = (ColIndex >= ord('I')))

    ChartSheet = WorkBook.add_chartsheet('%s▲'%(SheetName))
    ChartSheet.set_chart(Chart)

    print('OK.')


def CloseAndBrowse(WorkBook: xlsxwriter.Workbook, AutoOpen):
    try:
        print("\nSaving file [%s] ... "%(WorkBook.filename), end='')

        ticks = time.time()
        WorkBook.close()
        ticks = time.time() - ticks 

        # print("OK. Time elapsed: ", time.strftime('%M:%S', time.localtime(ticks)), flush=True)
        print("OK. Time elapsed: %0.3f seconds"% (ticks), flush=True)

        if AutoOpen: 
            print("Opening Excel file ... ", end='', flush=True)
            webbrowser.open(WorkBook.filename)
            print('OK.')

        print("\n")
    except:
        # raise
        print("failed. File already opened?\n")


def ParseJSONrecursively(Dict):
    for key, value in Dict.items():
        if type(value) is str:
            try:
                Dict[key] = json.loads(value)
            except:
                pass
            
        if type(Dict[key]) is dict:
            ParseJSONrecursively(Dict[key])
 
def ProcessOverallToXlsx(WorkBook, CountriesData):
    ''' CountriesData # Array of all countries '''

    AsDate     = WorkBook.add_format({'font_name': 'calibri', 'num_format': 'yyyy-mm-dd'})
    AsString   = WorkBook.add_format({'font_name': 'calibri', })
    AsPercent1 = WorkBook.add_format({'font_name': 'calibri', 'num_format': '0.0%'})
    AsPercent2 = WorkBook.add_format({'font_name': 'calibri', 'num_format': '0.00%'})
    AsPercent3 = WorkBook.add_format({'font_name': 'calibri', 'num_format': '0.0000%'})
    AsNumber   = WorkBook.add_format({'font_name': 'calibri', 'num_format': '#,##0_ '})

    WorkSheet = WorkBook.add_worksheet('总览')
    WorkSheet.set_column('B:B', 12)
    WorkSheet.set_column('U:U', 12)

    DestRow = 1 

    for Country in CountriesData:
        Col = -1
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['continent'], AsString)
        Col += 1; WorkSheet.write_url(DestRow, Col,       'internal:' + Country['name'] + '!A1', string=Country['name'])
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['id'], AsString)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryTotal']['confirmedTotal'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryTotal']['suspectedTotal'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryTotal']['curesTotal'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryTotal']['deathsTotal'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryTotal']['treatingTotal'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryTotal']['inboundTotal'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryTotal']['asymptomaticTotal'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col, float(Country['countryTotal']['deathRatio'].strip('%'))/100, AsPercent1)
        Col += 1; WorkSheet.write    (DestRow, Col, float(Country['countryTotal']['curesRatio'].strip('%'))/100, AsPercent1)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryIncr']['confirmedIncr'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryIncr']['suspectedIncr'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryIncr']['curesIncr'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryIncr']['deathsIncr'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col,       Country['countryIncr']['treatingIncr'], AsNumber)
        Col += 1; WorkSheet.write    (DestRow, Col, float(Country['confirmedPerMil'])/1000000, AsPercent1)
        Col += 1; WorkSheet.write    (DestRow, Col, '=F%d/(F%d+G%d)'       % (DestRow + 1, DestRow + 1, DestRow + 1), AsPercent1)   # 总治愈率
        Col += 1; WorkSheet.write    (DestRow, Col, '=1-S%d'               % (DestRow + 1), AsPercent2)                             # 总死亡率
        Col += 1; WorkSheet.write    (DestRow, Col, '=IFERROR(D%d/R%d,"")' % (DestRow + 1, DestRow + 1), AsNumber)                  # 国民总人口
        Col += 1; WorkSheet.write    (DestRow, Col, '=K%d*R%d'             % (DestRow + 1, DestRow + 1), AsPercent2)                # 国民死亡率
        Col += 1; WorkSheet.write    (DestRow, Col, '=IFERROR(M%d/U%d,"")' % (DestRow + 1, DestRow + 1), AsPercent3)                # 新增率

        Col += 1; WorkSheet.write(DestRow, Col, time.strftime('%Y-%m-%d', time.localtime(Country['updateTime'])), AsDate)

        DestRow += 1 

    WorkSheet.add_table(
            "A1:%c%d"%(chr(ord('A')+ Col), DestRow),  {
                'header_row': True, 'autofilter': True,  'style': 'TableStyleMedium3',
                'columns': [
                        {'header': '大洲'},                 # t 
                        {'header': '国家'},                 # B
                        {'header': 'ID'},                  # A
                        {'header': '总确诊'},               # C
                        {'header': '总疑似'},               # D
                        {'header': '总治愈'},               # E
                        {'header': '总死亡'},               # F
                        {'header': '治疗中'},               # G
                        {'header': '总输入'},               # H
                        {'header': '无症状总数'},           # I
                        {'header': '死亡率'},               # J 
                        {'header': '治愈率'},               # k 
                        {'header': '新增确诊'},             # l 
                        {'header': '新增疑似'},             # m 
                        {'header': '新增治愈'},             # n 
                        {'header': '新增死亡'},             # o 
                        {'header': '新增治疗'},             # p 
                        {'header': '感染率'},               # q 
                        {'header': '患者治愈率'},           # r 
                        {'header': '患者死亡率'},           # s 
                        {'header': '国民总人口'},           # t 
                        {'header': '国民死亡率'},           # u 
                        {'header': '新增率'},           # u 
                        {'header': '更新时间'},             # v
                    ]
            }
        )
    

def print_usage():
    print('covidspider.py [-c] [-d] [-g] [-j]')
    print('    -a days\tAverage days. Smooth the comfirmed number by several days.')
    print('    -c\t\tOpen epidemic xlsx file of China automaticlly after finished.')
    print('    -d\t\tDebug mode. Only fetch data of the first country.')
    print('    -g\t\tOpen epidemic xlsx file of Global automaticlly after finished.')
    print('    -j\t\tSave data file in JSON file.')

# Entry starts here

try:
    argv = sys.argv[1:]
    opts, args = getopt.getopt(argv,"a:hcdgj")
except getopt.GetoptError:
    print_usage()
    sys.exit(2)

for opt, arg in opts:
    if opt == '-h':
        print_usage()
        sys.exit(0)
    elif opt == '-c':
        AutoOpenChina = True
    elif opt == '-g':
        AutoOpenGlobal = True
    elif opt == '-d':
        Debugging = True
    elif opt == '-j':
        ToSaveJson = True
    elif opt == '-a':
        SmoothDays = int(arg)
        if SmoothDays < 1:
            print("Error: Days must > 0")
            sys.exit(2)



if SmoothDays > 1:
    filename_china  = 'epidemic-%s-%dDsmoothy-China.xlsx'%(date.today().strftime('%Y%m%d'), SmoothDays)
    filename_global = 'epidemic-%s-%dDsmoothy-Global.xlsx'%(date.today().strftime('%Y%m%d'), SmoothDays)
else:
    filename_china  = 'epidemic-%s-China.xlsx'%(date.today().strftime('%Y%m%d'))
    filename_global = 'epidemic-%s-Global.xlsx'%(date.today().strftime('%Y%m%d'))


QueryWorldDataUrl ='https://i.snssdk.com/forum/ncov_data/?data_type=[2,4,8]'
print('Fetching Global and China data ... ', end='', flush=True)

r = requests.get(url=QueryWorldDataUrl)
if r.status_code == 200:
    # Yes, the website has returned newest data. Let's clear previous data, and save the newest data
    print('OK.')

    ClearTempDir()

else:
    print('Get data error. Check network.')
    exit(-1)

try:
    WorldDict = r.json()

    if type(WorldDict) is dict:
        ParseJSONrecursively(WorldDict)

    SaveJson(WorldDict, 'World')


    ##############################################################################################
    # Save data of Chinese Provinces
    ##############################################################################################

    WorkBookChina = CreateWorkbook(filename_china)
    Provinces = WorldDict['ncov_nation_data']['provinces']
    AddToSheet(WorkBook=WorkBookChina, SheetName='全国',  Series = WorldDict['ncov_nation_data']['nationwide'])
    for Province in Provinces:
        AddToSheet(WorkBook=WorkBookChina, SheetName=Province['name'], Series= Province['series'])

    CloseAndBrowse(WorkBookChina, AutoOpen = AutoOpenChina)
    ##############################################################################################


    ##############################################################################################
    # Save data of World and China
    ##############################################################################################
    WorkBookWorld = CreateWorkbook(filename_global)
    ProcessOverallToXlsx(WorkBookWorld, WorldDict['overseas_data']['country'])

    AddToSheet(SheetName='全球', WorkBook=WorkBookWorld,  Series = WorldDict['overseas_data']['series'])
    AddToSheet(SheetName='中国', WorkBook=WorkBookWorld,  Series = WorldDict['ncov_nation_data']['nationwide'])

    # Fetch and save data of all foreign countries in the world
    Countries = WorldDict["ncov_nation_data"]["world"]
    i = 1;  CountriesCount = len(Countries)
    ticks0 = time.time()

    for Country in Countries:  # array of dicts of countries
        print("%0.2fs\t%0.1f%%"%(time.time()-ticks0, i*100/CountriesCount), end='\t'); i += 1
        FetchCountryData(WorkBook= WorkBookWorld,ID = Country['id'], Name = Country['name'])
        if Debugging:
            break

    CloseAndBrowse(WorkBookWorld, AutoOpen = AutoOpenGlobal)
    ##############################################################################################
    
except:
    print("The returned data is not as expected. ")
    raise