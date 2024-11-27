import xlwt
import os
import tkinter
from tkinter import *
from tkinter import ttk
import xlrd
import tkinter.filedialog as ttfiledialog
import re
import json
from xlutils.copy import copy as xlscopy
import pyperclip as MyClipboard
import pyautogui
import time
import requests

def gooletranslate(key, source_language='CN-简体', target_language='EN-英语'):
    target_language_dict={
    "CN-简体":"ZhCN",
    "EN-英语":"En",
    "TW-繁体":"ZhTw",
    "ES-西班牙":"Es",
    "PL-波兰":"Pl",
    "PT-葡萄牙":"Pt",
    "RU-俄语":"Ru",
    "FR-法语":"Fr",
    "DE-德语":"De",
    "IT-意大利":"It",
    "JA-日语":"Ja",
    "FI-芬兰":"Fi",
    "VN-越南":"Vi",
    "KR-韩语":"Ko",
    "AR-阿拉伯":"Ar",
    "TR-土耳其":"Tr",
    "TH-泰语":"Th",
    "HU-匈牙利":"Hu",
    "EL-希腊":"El",
    "NL-荷兰":"Nl",
    "NO-挪威":"No",
    "MA-马来西亚":"My",
    "FA-波斯":"Fa",
    "DA-丹麦":"Da",
    "RO-罗马尼亚":"Ro",
    "BR-葡萄牙(巴西)":"Br",
    "ID-印尼":"Id"
    }
    source_language_dict={
    "CN-简体":"ZhCN",
    "EN-英语":"En",
    "TW-繁体":"ZhTw",
    "ES-西班牙":"Es",
    "PL-波兰":"Pl",
    "PT-葡萄牙":"Pt",
    "RU-俄语":"Ru",
    "FR-法语":"Fr",
    "DE-德语":"De",
    "IT-意大利":"It",
    "JA-日语":"Ja",
    "FI-芬兰":"Fi",
    "VN-越南":"Vi",
    "KR-韩语":"Ko",
    "AR-阿拉伯":"Ar",
    "TR-土耳其":"Tr",
    "TH-泰语":"Th",
    "HU-匈牙利":"Hu",
    "EL-希腊":"El",
    "NL-荷兰":"Nl",
    "NO-挪威":"No",
    "MA-马来西亚":"My",
    "FA-波斯":"Fa",
    "DA-丹麦":"Da",
    "RO-罗马尼亚":"Ro",
    "BR-葡萄牙(巴西)":"Br",
    "ID-印尼":"Id"
    }
    if target_language not in target_language_dict:
        return "不支持该种语言翻译"
    target_language = target_language_dict[target_language]
    translate_url = "http://19.87.8.22:36604/wordsTranslationDict/translate"
    translate_userinfo = "eyJhY2Nlc3NfdG9rZW4iOiI1MjIzZjZiNTgxMmY0NmMyYTJhMjgyM2Q4MGJhNWI1NiIsInJlZnJlc2hfdG9rZW4iOiI4NWM4MDcwOWI2NzA0YTk5OWJmYzE5ZGRiODM5MWE5NyIsImRvbWFpbiI6Ilh0b29sdGVjaCIsImlkIjoiMjMxMDE2MzQiLCJuaWNrX25hbWUiOiLogpblmInkvJ8iLCJhdmF0YXJfdXJsIjpudWxsLCJwaG9uZSI6bnVsbCwiaXNCaW5kRGluZ3RhbGsiOmZhbHNlLCJleHBpcmUiOjE3MjA1MjI4NzMyNDEsInJlcyI6W119"

    headers = {'userinfo': translate_userinfo,
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                             'Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0'}
    payload = {
        'CC': 'PC',
        'SourceLanguage': "ZhCN",
        'TargetLanguage': target_language,
        'words[0]': key  
    }
    response = requests.post(translate_url, headers=headers, data=payload)
    message = ""
    if response.status_code == 200:
        response_data = response.json()
        try:
            translation_result = response_data['data']['WordsResults'][0]['Result']
            message = translation_result
        except Exception as e:
            message = "--Undefined--" + str(e)
    return message
    
class myTranslator:
    def __init__(self):
        self.NoTranFlah = "--Undefined--"
        self.NOTranDataDict = {}
        self.logInfo=[]
        self.UNIT_dict = {}
        self.encodingDict={
        "HU-匈牙利":"1250",
        "RO-罗马尼亚":"1250",
        "CN-简体":"936",
        "TW-繁体":"950",
        "CS":"1250",
        "DA-丹麦":"1252",
        "DE-德语":"1252",
        "EL-希腊":"1253",
        "EN-英语":"utf-8",
        "ES-西班牙":"1252",
        "AR-阿拉伯":"1256",
        "FI-芬兰":"1252",
        "FR-法语":"1252",
        "IT-意大利":"1252",
        "JA-日语":"932",
        "KR-韩语":"949",
        "NL-荷兰":"1252",
        "NO-挪威":"1142(ISO-8859-10)",
        "PL-波兰":"1250",
        "PT-葡萄牙":"1252",
        "RU-俄语":"1251",
        "SV":"1252",
        "TR-土耳其":"1254",
        "FA-波斯":"1256",
        "MA-马来西亚":"utf-8",#
        "SR":"1252",
        "VN-越南":"1258",
        "ID-印尼":"Unicode",
        "BR-葡萄牙(巴西)":"1252"}
    
    def on_closing(self):
        self.MyGUI.after_cancel(self.updateId)
        self.MyGUI.destroy()

    def AddValue2Dict(self,k,v):
        if k not in self.ExistDataDict:
            if len(v)!=0:
                self.ExistDataDict[k] = v

    def creatFile(self):
        self.strVehicelxls = self.strVehicel+".xls"
        folder = os.path.exists(self.strVehicelxls)
        if not folder:
            xls = xlwt.Workbook()
            sht1 = xls.add_sheet("已经校验翻译")
            sht2 = xls.add_sheet("现有翻译")
            sht3 = xls.add_sheet("有道翻译")
            sht4 = xls.add_sheet("未翻译")
            # 设置字体格式
            Font0=xlwt.Font()
            Font0.name="Times New Roman"
            Font0.colour_index = 2
            Font0.bold = True # 加粗
            style0=xlwt.XFStyle()
            listLangue = ["CN-简体","EN-英语","TW-繁体","ES-西班牙","PL-波兰","PT-葡萄牙","RU-俄语","FR-法语","DE-德语","IT-意大利","JA-日语","FI-芬兰","VN-越南","KR-韩语","AR-阿拉伯","TR-土耳其","TH-泰语","HU-匈牙利","EL-希腊","NL-荷兰","NO-挪威","MA-马来西亚","FA-波斯","DA-丹麦","RO-罗马尼亚","BR-葡萄牙(巴西)","ID-印尼"]
            lit = 0
            for langue in listLangue:
                sht1.write(0,lit,langue)
                lit=lit+1
            listLangue = ["CN-简体","EN-英语","TW-繁体","ES-西班牙","PL-波兰","PT-葡萄牙","RU-俄语","FR-法语","DE-德语","IT-意大利","JA-日语","FI-芬兰","VN-越南","KR-韩语","AR-阿拉伯","TR-土耳其","TH-泰语","HU-匈牙利","EL-希腊","NL-荷兰","NO-挪威","MA-马来西亚","FA-波斯","DA-丹麦","RO-罗马尼亚","BR-葡萄牙(巴西)","ID-印尼"]

            lit = 0
            for langue in listLangue:
                sht2.write(0,lit,langue)
                lit=lit+1
                
            lit = 0
            for langue in listLangue:
                sht3.write(0,lit,langue)
                lit=lit+1
                
            sht4.write(0,0,"")
            xls.save(self.strVehicelxls)
        
        folder = os.path.exists(self.strOutPath)
        if not folder:
            os.makedirs(self.strOutPath)
            
        folder = os.path.exists("UNIT.json")
        if not folder:
            with open("UNIT.json","w") as CreateJS:
                json.dump(self.UNIT_dict,CreateJS,ensure_ascii=False)
                
        with open("UNIT.json",'r',encoding = "utf-8")as json_Read:
            self.UNIT_dict = json.load(json_Read)

    def xls2dict(self):
        self.CheckedDataDict = {}
        self.ExistDataDict = {}
        self.YodaoDataDict = {}
        mywork = xlrd.open_workbook(self.strVehicelxls)
        sheet = mywork.sheet_by_name("已经校验翻译")
        rows = sheet.nrows
        cols = sheet.ncols
        for col in range(1,cols):
            language = sheet.cell_value(0,col)
            if language.find(self.TargeLanguage)!=-1:
                cols = col
                break
        for row in range(1,rows):
            cnData = str(sheet.cell_value(row,0))
            targData = str(sheet.cell_value(row,cols)).replace("\n","\\n")
            self.CheckedDataDict[cnData] = targData
       
        sheet = mywork.sheet_by_name("现有翻译")
        rows = sheet.nrows
        cols = sheet.ncols
        for col in range(1,cols):
            language = sheet.cell_value(0,col)
            if language.find(self.TargeLanguage)!=-1:
                cols = col
                break
        for row in range(1,rows):
            cnData = str(sheet.cell_value(row,0))
            targData = str(sheet.cell_value(row,cols))
            self.ExistDataDict[cnData] = targData

        sheet = mywork.sheet_by_name("有道翻译")
        rows = sheet.nrows
        cols = sheet.ncols
        for col in range(1,cols):
            language = sheet.cell_value(0,col)
            if language.find(self.TargeLanguage)!=-1:
                cols = col
                break
        for row in range(1,rows):
            cnData = str(sheet.cell_value(row,0))
            targData = str(sheet.cell_value(row,cols))
            if len(targData)!=0:
                self.YodaoDataDict[cnData] = targData

    def creatUI(self):
        self.bTrans = 1
        self.f_path=""
        self.MyGUI = Tk()
        self.MyGUI.title("翻译文档")
        self.MyGUI.geometry('890x482+10+10')
        #语言选择
        LangueTips = Label(self.MyGUI,justify = 'left',anchor='n', text='选择翻译的语言：')
        LangueTips.place(x=0,y=0)
        self.LangueSelect = ttk.Combobox(self.MyGUI,width = 13)
        self.LangueSelect['value'] = ("国产四种-英西俄法","ALL-所有","EN-英语","TW-繁体","ES-西班牙","PL-波兰","PT-葡萄牙","RU-俄语","FR-法语","DE-德语","IT-意大利","JA-日语","FI-芬兰","VN-越南","KR-韩语","AR-阿拉伯","TR-土耳其","TH-泰语","HU-匈牙利","EL-希腊","NL-荷兰","NO-挪威","MA-马来西亚","FA-波斯","DA-丹麦","RO-罗马尼亚","BR-葡萄牙(巴西)","ID-印尼")
        self.LangueSelect.current(0)
        self.LangueSelect.grid(row=1, column=0, sticky='NS')
        
        #对象选择
        ObjTips = Label(self.MyGUI,justify = 'left', text='选择翻译的对象：')
        ObjTips.place(x=160,y=0)
        self.ObjSelect = ttk.Combobox(self.MyGUI)
        self.ObjSelect['value'] = ("国产翻译","ALL","DTC.txt","TEXT.txt","DS.txt","DTC_H.txt","MENU.txt","ROOT.txt","EXCEL.xls")
        self.ObjSelect.current(0)
        self.ObjSelect.grid(row=1, column=1, sticky='NS')
        
        #选择文件夹
        FileTips = Label(self.MyGUI,justify = 'left', text='选择翻译的目录：')
        FileTips.grid(row=0, column=2)
        self.DataSelect = ttk.Combobox(self.MyGUI)
        self.DataSelect['value']=("选择文件夹")
        self.DataSelect.grid(row=1, column=2, sticky='NS')
        
        youdao = Label(self.MyGUI,justify = 'left', text="使用翻译平台：")
        youdao.grid(row=0, column=3)
        
        self.Selectyoudao = ttk.Combobox(self.MyGUI)
        self.Selectyoudao['value']=("是","否")
        self.Selectyoudao.current(1)
        self.Selectyoudao.grid(row=1, column=3, sticky='NS')
        
        FileSelect = Button(self.MyGUI,text="开始翻译",bg='lightblue',command = self.TranStart)
        FileSelect.grid(row=1,column=4, sticky='EW')
        
        #FileSelect = Button(self.MyGUI,text="检查乱码",bg='lightblue',command = self.CheckGarbledCode)
        #FileSelect.grid(row=1,column=5, sticky='EW')

        FileSelect = Button(self.MyGUI,text="格式调整",bg='lightblue',command = self.CheckFormatStart)
        FileSelect.grid(row=1,column=5, sticky='EW')
        
        LogTips = Label(self.MyGUI,justify = 'left',text = "实时日志：")
        LogTips.grid(row=3,column = 0)
        self.LogBox = Text(self.MyGUI, width=125, height=30)
        self.LogBox.place(x=2,y=70)
        self.MyGUI.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.UpDataGUI()

    def UpDataGUI(self):
        self.MyGUI.update()
        self.updateId = self.MyGUI.after(1000,self.UpDataGUI)
        if self.Selectyoudao.get()=="是":
            self.Toyoudao = True
        else:
            self.Toyoudao = False
        if self.DataSelect.get()=="选择文件夹":
            self.MyGUI.after_cancel( self.updateId)
            self.GetPath()

    def GetPath(self):
        self.DataPath = ttfiledialog.askdirectory()
        self.DataPath = self.DataPath+"/"
        ipos1 = self.DataPath.find(r"GS_")
        ipos2 = self.DataPath.find(r"/data")
        self.strVehicel = self.DataPath[ipos1+3:ipos2]
        self.DataSelect.set(self.DataPath)
        self.UpDataGUI()
    
    def TranStart(self):
        if self.strVehicel=="":
            return 0
        self.LangueSelect.update()
        LangueList = []
        if self.LangueSelect.get()=="ALL-所有":
            LangueList = ["EN-英语","TW-繁体","ES-西班牙","PL-波兰","PT-葡萄牙","RU-俄语","FR-法语","DE-德语","IT-意大利","JA-日语","FI-芬兰","VN-越南","KR-韩语","AR-阿拉伯","TR-土耳其","TH-泰语","HU-匈牙利","EL-希腊","NL-荷兰","NO-挪威","MA-马来西亚","FA-波斯","DA-丹麦","RO-罗马尼亚","BR-葡萄牙(巴西)","ID-印尼"]
        elif self.LangueSelect.get()=="国产四种-英西俄法":
            LangueList = ["EN-英语","ES-西班牙","RU-俄语","FR-法语"]
        else:
            LangueList = [self.LangueSelect.get()]
        self.ObjSelect.update()
        FileTypeList = []
        if self.ObjSelect.get()=="ALL":
            FileTypeList = ["DTC.txt","TEXT.txt","DS.txt","DTC_H.txt","MENU.txt","ROOT.txt","EXCEL.xls"]
        elif self.ObjSelect.get()=="国产翻译":
            FileTypeList = ["ROOT.txt","MENU.txt","TEXT.txt","DS.txt","DTC_H.txt","DTC.txt"]
        else:
            FileTypeList = [self.ObjSelect.get()]
        
        for Langue in LangueList:
            self.strOutPath = self.DataPath+Langue+"/"
            self.creatFile()
            for FileType in FileTypeList:
                self.TargeLanguage = Langue
                self.FileType = FileType
                self.SelectFun()
        self.UpdataLog("*************************************END OF ALL*************************************")
        self.mylog()

    def UpDict(self):
        with open("UNIT.json","w",encoding = "utf-8") as CreateJS:
            json.dump(self.UNIT_dict,CreateJS,ensure_ascii=False)

    def UpProgramXls(self):
        rd = xlrd.open_workbook(self.strVehicelxls, formatting_info = True)
        #xlrd_sheet1 = rd.sheet_by_name("已经校验翻译")
        xlrd_sheet2 = rd.sheet_by_name("现有翻译")
        xlrd_sheet3 = rd.sheet_by_name("有道翻译")
        xlrd_sheet4 = rd.sheet_by_name("未翻译")
        wt = xlscopy(rd)
        #sheet1 = wt.get_sheet("已经校验翻译")
        sheet2 = wt.get_sheet("现有翻译")
        sheet3 = wt.get_sheet("有道翻译")
        sheet4 = wt.get_sheet("未翻译")
        #已经检验
        #targeCol=0
        #for col in range(0,xlrd_sheet1.ncols):
        #    language = xlrd_sheet1.cell_value(0,col)
        #    if language.startswith(self.TargeLanguage):
        #        targeCol = col
        #        break
        #row=0
        #for k,v in self.CheckedDataDict.items():
        #    sheet1.write(row,0,k)
        #    sheet1.write(row,targeCol,v)
        #    row = row+1
        #现有
        targeCol=0
        for col in range(0,xlrd_sheet2.ncols):
            language = xlrd_sheet2.cell_value(0,col)
            if language.startswith(self.TargeLanguage):
                targeCol = col
                break
        row=1
        for k,v in self.ExistDataDict.items():
            sheet2.write(row,0,k)
            sheet2.write(row,targeCol,v)
            row = row+1
        #有道
        targeCol=0
        for col in range(0,xlrd_sheet3.ncols):
            language = xlrd_sheet3.cell_value(0,col)
            if language.startswith(self.TargeLanguage):
                targeCol = col
                break
        row=1
        for k,v in self.YodaoDataDict.items():
            sheet3.write(row,0,k)
            sheet3.write(row,targeCol,v)
            row = row+1
        #无翻译
        for row in range(0,xlrd_sheet4.nrows):
            sheet4.write(row,0,"")#清空
            sheet4.write(row,1,"")#清空
        row=0
        for k,v in self.NOTranDataDict.items():
            sheet4.write(row,0,k)
            sheet4.write(row,1,v)
            row = row+1
        try:
            wt.save(self.strVehicelxls)
        except:
            self.UpdataLog("大佬，先关闭这个文件哈："+self.strVehicelxls)
    def SelectFun(self):
        self.UpdataLog("*******************************************************开始翻译："+self.TargeLanguage+"_"+self.FileType)
        self.xls2dict()
        self.CNFile = self.DataPath+"CN_"+self.FileType
        self.CNFileNew = self.DataPath+"CN_"+"NEW_"+self.FileType
        Language = self.TargeLanguage.split("-")[0]
        self.OutFile = self.strOutPath+Language+"_"+self.FileType
        self.TargeFile = self.DataPath+Language+"_"+self.FileType
        if self.FileType == "DTC.txt":
            self.CNDTCToOther()
        elif self.FileType == "DTC_H.txt":
            self.CNDTC_HToOther()
        elif self.FileType == "TEXT.txt":
            self.CNTextToOther()
        elif self.FileType == "DS.txt":
            self.CNDSToOther()
        elif self.FileType == "EXCEL.xls":
            self.CNExcelToOther()
        elif self.FileType == "MENU.txt":
            self.CMenuToOther()
        elif self.FileType == "ROOT.txt":
            self.CRootToOther()
        self.UpDict()
        self.UpProgramXls()
        self.UpdataLog("*******************************************************成功翻译："+Language+"_"+self.FileType+"\n")

    def Myfilter(self,content):
        ipos = content.find("//")
        if ipos!=-1:
            if content[ipos-1] != ":":
                content = content.replace(content[ipos:],"")
        content = content.replace("\n","")
        if self.FileType != "MENU.txt" and self.FileType != "ROOT.txt":
            content = content.strip()
        else:
            myData = content
            if (myData.strip()==""):
                return ""
            content = content
        return content
        
    def CNDTCToOther(self):
        #self.DTC2dict()
        try:
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
                with open(self.OutFile,r"w",encoding = self.encodingDict[self.TargeLanguage], errors='replace') as Targe_Write:
                    for CNLine in CN_Read.readlines():
                        CNLine = self.Myfilter(CNLine)
                        #空行、注释行
                        if CNLine=="" or CNLine.find("/*")!=-1:
                            continue
                        if CNLine.find("include")!=-1:
                            Targe_Write.write(CNLine + "\n")
                            continue
                        DTC_List = CNLine.split("\t")
                        Index = DTC_List[0]
                        PCBU = '\t' + DTC_List[1]
                        MyRe = r"\"$|^\""
                        ENValueAll = ""
                        for i in range(2,len(DTC_List)):
                            CNValue = re.sub(MyRe,"",DTC_List[i]).strip()
                            if CNValue!="" and contains_chinese(CNValue) == True:
                                ENValue = self.getTargeValue(CNValue)
                            else:
                                ENValue = CNValue
                            ENValueAll = ENValueAll +'\t' + "\"" + ENValue + "\""
                        Targe_Write.write(Index + PCBU + ENValueAll + "\n")
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+ str(e) + CNLine)
    def DTC2dict(self):
        CN_dict = {}
        Targe_dict = {}
        self.error = "无法打开文件"
        try:
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
                for CNLine in CN_Read.readlines():
                    self.error = CNLine
                    CNLine = self.Myfilter(CNLine)
                    if CNLine=="" or CNLine.find("include")!=-1:
                        continue
                    CNLineList = CNLine.split("\t")
                    if len(CNLineList)>=3:
                        key = CNLineList[0]
                        value = '\t'.join(CNLineList[2:])
                        # 判断并添加‘*’标记直到没有重复的
                        while key in CN_dict:
                            key += '*'  # 在键的末尾添加‘*’
                        CN_dict[key] = value  # 将最终的唯一键与值插入到字典中
                    else:
                        self.UpdataLog(self.FileType+" 格式错误 "+CNLine)
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+self.error+ str(e) + CNLine)
        self.error = "无法打开文件"
        try:
            with open(self.TargeFile,r"r",encoding = self.encodingDict[self.TargeLanguage]) as TargeRead:
                for TargeLine in TargeRead.readlines():
                    self.error = TargeLine
                    TargeLine = self.Myfilter(TargeLine)
                    if TargeLine=="":
                        continue
                    TargeLineList = TargeLine.split("\t")
                    if len(TargeLineList)>=3:
                        Targe_dict[TargeLineList[0].strip()] = TargeLineList[2:].strip()
                    else:
                        self.UpdataLog(self.FileType+" 格式错误 "+TargeLine+ str(e) + CNLine)
        except :
            self.UpdataLog(self.TargeFile+" 错误信息："+self.error + CNLine)

    def CNDTC_HToOther(self):
        if os.path.exists(self.CNFile):
            #self.DTC_H2dict()
            try:
                with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
                    with open(self.OutFile,r"w",encoding = self.encodingDict[self.TargeLanguage], errors='replace') as Targe_Write:
                        for CNLine in CN_Read.readlines():
                            CNLine = self.Myfilter(CNLine)
                            removeBlank = CNLine.strip()
                            #空行、注释行
                            if removeBlank=="":
                                Targe_Write.write("\n")
                                continue
                            if CNLine.find("include")!=-1 or CNLine.find("@$pdf")!=-1:
                                Targe_Write.write(CNLine + "\n")
                                continue
                            DTC_HList = CNLine.split("\t")
                            Index = DTC_HList[0]
                            MyRe = r"\"$|^\""
                            ENValueAll = ""
                            for i in range(1,len(DTC_HList)):
                                CNValue = re.sub(MyRe,"",DTC_HList[i]).strip()
                                if CNValue != "" and contains_chinese(CNValue) == True:
                                    ENValue = self.getTargeValue(CNValue)
                                else:
                                    ENValue = CNValue
                                ENValueAll = ENValueAll +'\t' + "\"" + ENValue + "\""
                            Targe_Write.write(Index + ENValueAll + "\n")
            except Exception as e:
                self.UpdataLog(self.CNFile+" 错误信息："+ str(e))

    def DTC_H2dict(self):
        try:
            Myre = r"\"$|^\""
            CN_dict = {}
            Targe_dict = {}
            self.error = "无法打开文件"
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
                for CNLine in CN_Read.readlines():
                    self.error = CNLine
                    CNLine = self.Myfilter(CNLine)
                    removeBlank = CNLine.strip()
                    #空行、注释行
                    if removeBlank=="":
                        continue
                    if CNLine.find("include")!=-1 or CNLine.find("@$pdf")!=-1:
                        continue
                    CNLineList = CNLine.split("\t")
                    if len(CNLineList)==2:
                        if CNLineList[0] in CN_dict:
                            self.UpdataLog(self.FileType+" Index重复 "+CNLine)
                        CN_dict[CNLineList[0]] = re.sub(Myre,"",CNLineList[1]).strip()
                    elif len(CNLineList)==3:
                       if CNLineList[0] in CN_dict:
                           self.UpdataLog(self.FileType+" Index重复 "+CNLine)
                       CN_dict[CNLineList[0].strip()] = CNLineList[2].strip()    
                    else:
                        self.UpdataLog(self.FileType+" 格式错误 "+CNLine)
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+self.error+ str(e))
        try:
            self.error = "无法打开文件"
            with open(self.TargeFile,r"r",encoding = self.encodingDict[self.TargeLanguage]) as TargeRead:
                for TargeLine in TargeRead.readlines():
                    self.error = TargeLine
                    TargeLine = self.Myfilter(TargeLine)
                    removeBlank = TargeLine.strip()
                    #空行、注释行
                    if removeBlank=="":
                        continue
                    TargeLineList = TargeLine.split("\t")
                    if len(TargeLineList)==2:
                        if TargeLineList[0] in Targe_dict:
                            self.UpdataLog(self.FileType+" Index重复 "+TargeLine)
                        Targe_dict[TargeLineList[0]] = re.sub(Myre,"",TargeLineList[1]).strip()
                    elif len(TargeLineList)==3:
                        Targe_dict[TargeLineList[0].strip()] = TargeLineList[2].strip()
                    else:
                        self.UpdataLog(self.FileType+" 格式错误 "+TargeLine)
        except:
            self.UpdataLog(self.TargeFile+" 错误信息："+self.error)

    def CNTextToOther(self):
        #self.Text2dict()
        try:
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"]) as CN_Read:
                with open(self.OutFile,r"w",encoding = self.encodingDict[self.TargeLanguage], errors='replace') as Targe_Write:
                    for CNLine in CN_Read.readlines():
                        CNLine = self.Myfilter(CNLine)
                        removeBlank = CNLine.strip()
                        if CNLine.find("include")!=-1:
                            Targe_Write.write(CNLine + "\n")
                            continue
                        #空行、注释行
                        if removeBlank=="":
                            Targe_Write.write("\n")
                            continue
                        else:
                            CNLineList = CNLine.split("\t")
                            ENLine = CNLineList[0]+"\t"
                            for i in range(1,len(CNLineList)):
                                MyRe = r"\"$|^\""
                                CNValue = re.sub(MyRe,"",CNLineList[i])
                                CNValue = CNValue.strip()
                                ENValue = "\"\""
                                if CNValue!="":
                                    ENValue = self.getTargeValue(CNValue)
                                if CNValue!="":
                                    ENLine = ENLine+ CNLineList[i].replace(CNValue,ENValue)+"\t"
                                else:
                                    ENLine = ENLine+ CNLineList[i]+"\t"
                            ENLine = ENLine[0:-1]
                            Targe_Write.write(ENLine + "\n")
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+ str(e))

    def Text2dict(self):
        try:
            self.error = "无法打开文件"
            CN_dict = {}
            Targe_dict = {}
            Myre = r"\"$|^\""
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"]) as CN_Read:
                for CNLine in CN_Read.readlines():
                    if CNLine.find("include")!=-1:
                        continue
                    self.error = CNLine
                    CNLine = self.Myfilter(CNLine)
                    removeBlank = CNLine.strip()
                    #空行、注释行
                    if removeBlank=="":
                        continue
                    CNLineList = CNLine.split("\t")
                    if CNLineList[0] in CN_dict:
                        self.UpdataLog(self.FileType+" Index重复 "+CNLine)
                    if len(CNLineList)==2:
                        CN_dict[CNLineList[0]] = re.sub(Myre,"",CNLineList[1]).strip()
                    elif len(CNLineList)==3:
                        CN_dict[CNLineList[0]] =re.sub(Myre,"",CNLineList[1]).strip()
                        CN_dict[CNLineList[0]+"_1"] =re.sub(Myre,"",CNLineList[2]).strip()
                    else:
                        self.UpdataLog(self.FileType+" 格式错误 "+CNLine)
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+self.error + str(e))
        try:
            self.error = "无法打开文件"
            with open(self.TargeFile,r"r",encoding = self.encodingDict[self.TargeLanguage]) as TargeRead:
                for TargeLine in TargeRead.readlines():
                    self.error = TargeLine
                    TargeLine = self.Myfilter(TargeLine)
                    removeBlank = TargeLine.strip()
                    #空行、注释行
                    if removeBlank=="":
                        continue
                    TargeLineList = TargeLine.split("\t")
                    if TargeLineList[0] in Targe_dict:
                        self.UpdataLog(self.FileType+" Index重复 "+TargeLine)
                    if len(TargeLineList)==2:
                        Targe_dict[TargeLineList[0]] = re.sub(Myre,"",TargeLineList[1]).strip()
                    elif len(TargeLineList)==3:
                        Targe_dict[TargeLineList[0]] = re.sub(Myre,"",TargeLineList[1]).strip()
                        Targe_dict[TargeLineList[0]+"_1"] = re.sub(Myre,"",TargeLineList[2]).strip()
                    else:
                        self.UpdataLog(self.FileType+" 格式错误 "+TargeLine)
        except :
            self.UpdataLog(self.TargeFile+" 错误信息："+self.error)
        self.FileToLocalDict(CN_dict,Targe_dict)

    def Menu2dict(self):
        try:
            self.error = "无法打开文件"
            CN_dict = {}
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"]) as CN_Read:
                 for CNLine in CN_Read.readlines():
                    self.error = CNLine
                    if CNLine.startswith("\t")==False:
                        continue
                    CNLine = self.Myfilter(CNLine)#去除注释、换行
                    CNLine = CNLine.replace("\t","")
                    ipos = CNLine.find("<")
                    if ipos!=-1:
                        CNLine = CNLine[0:ipos]
                    #直接翻译
                    if CNLine!="":
                        self.getTargeValue(CNLine)
                    #if CNLine not in self.CheckedDataDict:
                    #   if CNLine not in self.ExistDataDict:
                    #       if CNLine not in self.YodaoDataDict:
                    #           self.NOTranDataDict[CNLine] = self.FileType
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+self.error+ str(e))

    def CMenuToOther(self):
        #self.Menu2dict()
        try:
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"]) as CN_Read:
                with open(self.OutFile,r"w",encoding = self.encodingDict[self.TargeLanguage], errors='replace') as Targe_Write:
                    for CNLine in CN_Read.readlines():
                        CNLine = self.Myfilter(CNLine)#去除注释、换行
                        sourceLine = CNLine
                        if CNLine == "":
                            #Targe_Write.write("\n")
                            continue
                        if CNLine.startswith("\t") == False:
                            Targe_Write.write(CNLine+"\n")
                            continue
                        CNLine = self.Myfilter(CNLine)
                        NeedTran = CNLine.replace("\t","")
                        ipos = NeedTran.find("<")
                        if ipos!=-1:
                            NeedTran = NeedTran[0:ipos]
                        Targe = self.getTargeValue(NeedTran)
                        Targe = sourceLine.replace(NeedTran,Targe)
                        Targe_Write.write(Targe+"\n")
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+ str(e))
    def Root2dict(self):
        self.error = "无法打开文件"
        CN_dict = {}
        Targe_dict = {}
        try:
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
                 for CNLine in CN_Read.readlines():
                    CNLine = self.Myfilter(CNLine)#去除注释、换行
                    CNLine = CNLine.replace("\t","")
                    self.error = CNLine
                    ipos = CNLine.find("<0x")
                    if ipos !=-1:
                        Id = CNLine[ipos:]
                        content = CNLine[0:ipos]
                        CN_dict[Id] = content
                    elif CNLine!="":
                        header = CNLine[0:2]
                        if header.find("@")!=-1:
                            CNLine = CNLine[2:]
                        self.getTargeValue(CNLine)
                        #if CNLine not in self.CheckedDataDict:
                        #    if CNLine not in self.ExistDataDict:
                        #        if CNLine not in self.YodaoDataDict:
                        #            self.NOTranDataDict[CNLine] = self.FileType
                            
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+self.error+ str(e))
        try:
            self.error = "无法打开文件"
            with open(self.TargeFile,r"r",encoding = self.encodingDict[self.TargeLanguage]) as TargeRead:
                 for TargeLine in TargeRead.readlines():
                    self.error = TargeLine
                    TargeLine = self.Myfilter(TargeLine)#去除注释、换行
                    TargeLine = TargeLine.replace("\t","")
                    ipos = TargeLine.find("<0x")
                    if ipos !=-1:
                        Id = TargeLine[ipos:]
                        content = TargeLine[0:ipos]
                        Targe_dict[Id] = content
        except :
            self.UpdataLog(self.TargeFile+" 错误信息："+self.error)
            
    def CRootToOther(self):
        #self.Root2dict()
        try:
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
                with open(self.OutFile,r"w",encoding = self.encodingDict[self.TargeLanguage], errors='replace') as Targe_Write:
                    for CNLine in CN_Read.readlines():
                        CNLine = self.Myfilter(CNLine)#去除注释、换行
                        sourceLine = CNLine
                        if CNLine == "":
                            #Targe_Write.write("\n")
                            continue
                        CNLine = self.Myfilter(CNLine)
                        NeedTran = CNLine.replace("\t","")
                        header = NeedTran[0:2]
                        if header.find("@")!=-1:
                            NeedTran = NeedTran[2:]
                        ipos = NeedTran.find("<")
                        if ipos!=-1:
                            NeedTran = NeedTran[0:ipos]
                        Targe = self.getTargeValue(NeedTran)
                        Targe = sourceLine.replace(NeedTran,Targe)
                        Targe_Write.write(Targe+"\n")
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+ str(e))

    def CNDSToOther(self):
        #self.DS2dict()
        with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
            with open(self.OutFile,r"w",encoding = self.encodingDict[self.TargeLanguage], errors='replace') as Targe_Write:
                for CNLine in CN_Read.readlines():
                    CNLine = self.Myfilter(CNLine)
                    removeBlank = CNLine.strip()
                    #空行、注释行
                    if removeBlank=="":
                        continue
                    if contains_chinese(removeBlank)==False:
                        Targe_Write.write(CNLine + "\n")
                        continue
                    if CNLine.find("include")!=-1:
                        Targe_Write.write(CNLine + "\n")
                        continue
                    CNLineList = removeBlank.split("\t")
                    if len(CNLineList)<6:
                        Targe_Write.write(CNLine + "\n")
                        continue
                    MyRe = r"\"$|^\""
                    CNName = re.sub(MyRe,"",CNLineList[1])
                    ENName = ""
                    if CNName!="":
                        ENName = self.getTargeValue(CNName)
                    CNUnit = CNLineList[2]
                    if re.sub(MyRe,"",CNUnit).strip() in self.UNIT_dict:
                        ENUnit = "\"" +self.UNIT_dict[re.sub(MyRe,"",CNUnit).strip()] +"\""
                    else:
                        ENUnit = "\"" +re.sub(MyRe,"",CNUnit).strip() +"\""
                        self.UNIT_dict[ENUnit] = ENUnit
                    try:
                        CNValue = re.sub(MyRe,"",CNLineList[3])
                        CNValueList = CNValue.split("|")
                        ENValue = ""
                        for CNOneValue in CNValueList:
                            CNOneValue = CNOneValue.strip()
                            
                            if CNOneValue.find("%")!=-1:
                                ENValue = CNOneValue
                            else:
                                if CNOneValue!="":
                                    CNOneValue = self.getTargeValue(CNOneValue)
                                ENValue = ENValue + CNOneValue
                            ENValue = ENValue+"|"
                        ENValue = ENValue[0:-1]
                        ENLine = ""
                        CNLine = ""
                        for i in CNLineList:
                            CNLine = CNLine+ i +"\t"
                        CNLineList[1] = CNLineList[1].replace(CNName,ENName)
                        CNLineList[2] = CNLineList[2].replace(CNUnit,ENUnit)
                        CNLineList[3] = CNLineList[3].replace(CNValue,ENValue)
                        for i in CNLineList:
                            ENLine = ENLine+ i +"\t"
                        ENLine = ENLine[0:-1]
                        CNLine = CNLine[0:-1]
                        Targe_Write.write(ENLine+"\n")
                    except Exception as e:
                        self.UpdataLog("CN_DS、"+CNLine+" 格式有问题"+ str(e))

    def DS2dict(self):
        try:
            self.error = "无法打开文件"
            CN_dict = {}
            Targe_dict = {}
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
                for CNLine in CN_Read.readlines():
                    self.error = CNLine
                    CNLine = self.Myfilter(CNLine)
                    removeBlank = CNLine.strip()
                    #空行、注释行
                    if removeBlank=="":
                        continue
                    if CNLine.find("include")!=-1:
                        continue
                    CNLineList = CNLine.split("\t")

                    if CNLineList[0] in CN_dict:
                        self.UpdataLog(self.FileType+" Index重复 "+CNLine)
                    if len(CNLineList)>3:
                        Index = CNLineList[0]
                        value = '\t'.join(CNLineList[1:])
                        CN_dict[Index] = value
                    else:
                        self.UpdataLog(self.FileType+" 格式错误 ",CNLine)
        except Exception as e:
            self.UpdataLog(self.CNFile+" 错误信息："+self.error+ str(e))
        try:
            self.error = "无法打开文件"
            with open(self.TargeFile,r"r",encoding = self.encodingDict[self.TargeLanguage], errors='replace') as TargeRead:
                for TargeLine in TargeRead.readlines():
                    self.error = TargeLine
                    TargeLine = self.Myfilter(TargeLine)
                    removeBlank = TargeLine.strip()
                    #空行、注释行
                    if removeBlank=="":
                        continue
                    TargeLineList = TargeLine.split("\t")
                    if TargeLineList[0] in Targe_dict:
                        self.UpdataLog(self.FileType+" Index重复 "+TargeLine)
                    if len(TargeLineList)>3:
                        Index = TargeLineList[0]
                        value = '\t'.join(TargeLineList[1:])
                        Targe_dict[Index] = value
                    else:
                        self.UpdataLog(self.FileType+" 格式错误 "+TargeLine)
        except:
            self.UpdataLog(self.TargeFile+" 错误信息："+self.error)
        self.FileToLocalDict(CN_dict,Targe_dict)

    def CNExcelToOther(self):
        Language = self.TargeLanguage.split("-")[0]
        TargeFileNames=os.listdir(self.strOutPath+"\\")
        TargeXlsList=[]
        for oneFile in TargeFileNames:
            if oneFile.endswith(".xls"):
                if oneFile.startswith(Language+"_"):
                    TargeXlsList.append(oneFile)
        SoureFileNames=os.listdir(self.DataPath)
        NeedTranXlsList=[]
        for oneFile in SoureFileNames:
            if oneFile.endswith(".xls"):
                if oneFile.startswith("CN_"):
                    TargeFile = oneFile.replace("CN_",Language+"_")
                    if TargeFile not in TargeXlsList:
                        NeedTranXlsList.append(self.DataPath+oneFile)
        for oneXLS in NeedTranXlsList:
            self.currentxls = oneXLS
            if oneXLS.find("QUICK_SCAN.xls")!=-1:
                continue
            self.CNPath = oneXLS
            self.TargeXlsExsit = False
            self.ENExcelDict = {}
            ENXLS = oneXLS.replace("CN_",Language+"_")
            self.ReadENXls(ENXLS)
            try:
                self.CN_work = xlrd.open_workbook(oneXLS,formatting_info = True)
                self.New_work = xlscopy(self.CN_work)
                self.read_cell("EcuInfo",3)
                self.read_cell("EcuInfo",9)
                
                self.read_cell("ReadCds",2)
                self.read_cell("ReadCds",3)
                self.read_cell("ReadCds",5)
                
                self.read_cell("FreezeFrame",2)
                self.read_cell("FreezeFrame",3)
                self.read_cell("FreezeFrame",5)
                
                self.read_cell("Text",1)
                
                self.read_cell("Stat",1)
                
                self.read_cell("Dtc",2)
                
                self.read_cell("Dtc",3)
                
                oneXLS = oneXLS.replace("CN_",self.TargeLanguage+"/"+Language+"_")
                self.New_work.save(oneXLS)
                self.UpdataLog(Language+" 成功翻译："+self.CNPath)
            except Exception as e:
                self.UpdataLog(Language+" 未能翻译："+self.CNPath+ str(e))

    def read_cell(self,sheetName,col):
        CN_sheet = self.CN_work.sheet_by_name(sheetName)
        self.New_sheet = self.New_work.get_sheet(sheetName)
        CN_rows = CN_sheet.nrows
        for row in range(1,CN_rows):
            my_cell_value = CN_sheet.cell_value(row,col)
            strid = str(CN_sheet.cell_value(row,0)).replace(".0","").strip()
            if len(my_cell_value)==0:
                continue
            elif col ==3 and (sheetName=="ReadCds" or sheetName=="FreezeFrame"):
                self.write_xls(strid,row,col,my_cell_value,sheetName,True)
            else:
                self.write_xls(strid,row,col,my_cell_value,sheetName)

    def write_xls(self,strid,row,col,value,SheetName,Unit = False):
        New_word = ""
        if Unit:#单位翻译
            if value in self.UNIT_dict:
                New_word = self.UNIT_dict[value]
            else:
                self.UNIT_dict[value] = value
                self.UpdataLog(SheetName+" "+strid+" 新的单位出现："+value)
                return 0
        elif col==5 and (SheetName=="ReadCds" or SheetName=="FreezeFrame"):#数据流格式翻译
            valueList = value.split("|")
            NewValueList = []
            if self.TargeXlsExsit:
                if len(valueList)!=len(self.ENExcelDict[SheetName]["Value"][strid]):
                    self.UpdataLog("中英文库的值不一样："+self.CNPath+"/t"+SheetName+"/t"+strid)
                    for oneValue in valueList:
                        XLSOneValue = oneValue
                        if oneValue.find("%")!=-1:
                            NewValueList.append(oneValue+"|")
                        else:
                            if XLSOneValue!="":
                                XLSOneValue = self.getTargeValue(XLSOneValue)
                            NewValueList.append(XLSOneValue+"|")
                else:
                    i = 0
                    for oneValue in valueList:
                        XLSOneValue = oneValue
                        myWord = ""
                        if oneValue.find("%")!=-1:
                            myWord = oneValue
                        else:
                            if oneValue in self.CheckedDataDict:
                                myWord = self.CheckedDataDict[XLSOneValue]
                            if myWord=="" and (oneValue in self.ExistDataDict):
                                myWord = self.ExistDataDict[XLSOneValue]
                            if myWord=="":
                                myWord = self.ENExcelDict[SheetName]["Value"][strid][strid+"_"+str(i)]
                                self.AddValue2Dict(XLSOneValue,myWord)
                            if myWord=="" and (oneValue in self.YodaoDataDict):
                                myWord = self.YodaoDataDict[XLSOneValue]
                            if myWord=="":
                                myWord = self.youdaoTranslate(XLSOneValue)
                        NewValueList.append(myWord+"|")
                        i = i+1
            else:
                for oneValue in valueList:
                    XLSOneValue = oneValue
                    if oneValue.find("%")!=-1:
                        NewValueList.append(oneValue+"|")
                    else:
                        if XLSOneValue!="":
                            XLSOneValue = self.getTargeValue(XLSOneValue)
                        NewValueList.append(XLSOneValue+"|")
            if len(NewValueList)!=0:
                NewValueList[-1] = NewValueList[-1].replace("|","")
            for oneValue in NewValueList:
                New_word = New_word+oneValue
        elif value.find("%")!=-1 and SheetName=="EcuInfo":#信息版本，没有文本
            return 0
        else:#正常串翻译
            XLSOneValue = value
            if XLSOneValue in self.CheckedDataDict:
                New_word = self.CheckedDataDict[XLSOneValue]
            if New_word == "" and (XLSOneValue in self.ExistDataDict):
                New_word = self.ExistDataDict[XLSOneValue]
            if New_word == "":
                if self.TargeXlsExsit:
                    if SheetName=="EcuInfo" or SheetName=="Text" or SheetName=="Stat":
                        if strid in self.ENExcelDict[SheetName]:
                            New_word = self.ENExcelDict[SheetName][strid]
                            if New_word!="":
                                self.AddValue2Dict(XLSOneValue,New_word)
                        if New_word=="" and (XLSOneValue in self.YodaoDataDict):
                            New_word = self.YodaoDataDict[XLSOneValue]
                        if New_word=="":
                            New_word = self.youdaoTranslate(XLSOneValue,self.CNPath)
                    elif SheetName=="ReadCds" or SheetName=="FreezeFrame":
                        if strid in self.ENExcelDict[SheetName]["Name"]:
                            New_word = self.ENExcelDict[SheetName]["Name"][strid]
                            if New_word!="":
                                self.AddValue2Dict(XLSOneValue,New_word)
                        if New_word=="" and (XLSOneValue in self.YodaoDataDict):
                            New_word = self.YodaoDataDict[XLSOneValue]
                        if New_word=="":
                            New_word = self.youdaoTranslate(XLSOneValue,self.CNPath)
                    elif SheetName=="Dtc":
                        if strid in self.ENExcelDict[SheetName]["Description"]:
                            if col==2:
                                New_word = self.ENExcelDict[SheetName]["Description"][strid]
                            else:
                                New_word = self.ENExcelDict[SheetName]["Help"][strid]
                            if len(New_word)!=0:
                                self.AddValue2Dict(XLSOneValue,New_word)
                            else:
                                if XLSOneValue in self.YodaoDataDict:
                                    New_word = self.YodaoDataDict[XLSOneValue]
                                if New_word == "":
                                    New_word = self.youdaoTranslate(XLSOneValue,self.CNPath)
                        elif XLSOneValue in self.YodaoDataDict:
                            New_word = self.YodaoDataDict[XLSOneValue]
                        if New_word=="":
                            New_word = self.youdaoTranslate(XLSOneValue,self.CNPath)
                
                if New_word=="" and XLSOneValue!="":
                    New_word = self.getTargeValue(XLSOneValue)
                    
        self.New_sheet.write(row,col,New_word)#重写

    def ReadENXls(self,EnXlsName):
        self.ENExcelDict["EcuInfo"] = {}
        self.ENExcelDict["ReadCds"] = {}
        self.ENExcelDict["ReadCds"]["Name"] = {}
        self.ENExcelDict["ReadCds"]["Value"] = {}
        self.ENExcelDict["FreezeFrame"] = {}
        self.ENExcelDict["FreezeFrame"]["Name"] = {}
        self.ENExcelDict["FreezeFrame"]["Value"] = {}
        self.ENExcelDict["Text"] = {}
        self.ENExcelDict["Stat"] = {}
        self.ENExcelDict["Dtc"] = {}
        self.ENExcelDict["Dtc"]["Description"] = {}
        self.ENExcelDict["Dtc"]["Help"] = {}
        try:
            ENWork = xlrd.open_workbook(EnXlsName,formatting_info = True)
            EcuInfoSheet = ENWork.sheet_by_name("EcuInfo")
            for i in range(0,EcuInfoSheet.nrows):
                strid = str(EcuInfoSheet.cell_value(i,0)).replace(".0","").strip()
                strName = str(EcuInfoSheet.cell_value(i,3)).strip()
                self.ENExcelDict["EcuInfo"][strid] = strName
            EcuInfoSheet = ENWork.sheet_by_name("ReadCds")
            for i in range(0,EcuInfoSheet.nrows):
                strid = str(EcuInfoSheet.cell_value(i,0)).replace(".0","").strip()
                strName = str(EcuInfoSheet.cell_value(i,2)).strip()
                strValue = str(EcuInfoSheet.cell_value(i,5)).strip()
                self.ENExcelDict["ReadCds"]["Name"][strid] = strName
                ValueList = strValue.split("|")
                self.ENExcelDict["ReadCds"]["Value"][strid] = {}
                for j in range(0,len(ValueList)):
                    strNewid = strid +"_"+ str(j).replace(".0","")
                    self.ENExcelDict["ReadCds"]["Value"][strid][strNewid] = ValueList[j]
            EcuInfoSheet = ENWork.sheet_by_name("FreezeFrame")
            for i in range(0,EcuInfoSheet.nrows):
                strid = str(EcuInfoSheet.cell_value(i,0)).replace(".0","").strip()
                strName = str(EcuInfoSheet.cell_value(i,2)).strip()
                strValue = str(EcuInfoSheet.cell_value(i,5)).strip()
                self.ENExcelDict["FreezeFrame"]["Name"][strid] = strName
                ValueList = strValue.split("|")
                self.ENExcelDict["FreezeFrame"]["Value"][strid] = {}
                for j in range(0,len(ValueList)):
                    strNewid = strid +"_"+ str(j).replace(".0","")
                    self.ENExcelDict["FreezeFrame"]["Value"][strid][strNewid] = ValueList[j]
            EcuInfoSheet = ENWork.sheet_by_name("Text")
            for i in range(0,EcuInfoSheet.nrows):
                strid = str(EcuInfoSheet.cell_value(i,0)).replace(".0","").strip()
                strName = str(EcuInfoSheet.cell_value(i,1)).strip()
                self.ENExcelDict["Text"][strid] = strName
            EcuInfoSheet = ENWork.sheet_by_name("Stat")
            for i in range(0,EcuInfoSheet.nrows):
                strid = str(EcuInfoSheet.cell_value(i,0)).replace(".0","").strip()
                strName = str(EcuInfoSheet.cell_value(i,1)).strip()
                self.ENExcelDict["Stat"][strid] = strName
            EcuInfoSheet = ENWork.sheet_by_name("Dtc")
            for i in range(0,EcuInfoSheet.nrows):
                strid = str(EcuInfoSheet.cell_value(i,0)).replace(".0","").strip()
                strDescription = str(EcuInfoSheet.cell_value(i,2)).strip()
                strHlep = str(EcuInfoSheet.cell_value(i,3)).strip()
                self.ENExcelDict["Dtc"]["Description"][strid] = strDescription
                self.ENExcelDict["Dtc"]["Help"][strid] = strHlep
            self.TargeXlsExsit = True
        except:
            self.UpdataLog("无法读取对应的表格："+EnXlsName)

    def UpdataLog(self,CurrentInfo):
        self.LogBox.insert(END,CurrentInfo+"\n")
        self.LogBox.see(END)
        self.LogBox.update()
        self.MyGUI.update()
        self.logInfo.append(CurrentInfo+"\n")

    def FileToLocalDict(self,CN_dict,Targe_dict):
        MyRe = r"\"$|^\""#去除字符串两端的冒号
        for key,value in CN_dict.items():
            if self.FileType=="DS.txt":
                CNValueList = value.split("\t")
                CNName = re.sub(MyRe,"",CNValueList[0]).strip()
                CNUnit = re.sub(MyRe,"",CNValueList[1]).strip()
                CNSubValue = re.sub(MyRe,"",CNValueList[2])
                if key in Targe_dict:#有对应的翻译
                    ENValueList = Targe_dict[key].split("\t")
                    ENName = re.sub(MyRe,"",ENValueList[0]).strip()
                    self.AddValue2Dict(CNName,ENName)
                    ENUnit = re.sub(MyRe,"",ENValueList[1]).strip()
                    if CNUnit not in self.UNIT_dict:
                        self.UpdataLog(self.FileType+" 新单位出现："+key)
                        self.UNIT_dict[CNUnit] = ENUnit
                    ENSubValue = re.sub(MyRe,"",ENValueList[2])
                    CNSubValueList = CNSubValue.split("|")
                    ENSubValueList = ENSubValue.split("|")
                    if len(CNSubValueList)!=len(ENSubValueList):
                        self.UpdataLog(self.FileType+"Index {strkey} 中英文库的值不一样".format(strkey = key))
                        if CNSubValueList[i].strip() not in self.ExistDataDict:
                            self.youdaoTranslate(CNSubValueList[i].strip())
                    else:
                        for i in range(len(CNSubValueList)):
                            if CNSubValueList[i].find("%")!=-1:#特殊字符不翻译
                                pass
                            else:
                                if CNSubValueList[i].strip() not in self.ExistDataDict:
                                    self.AddValue2Dict(CNSubValueList[i].strip(),ENSubValueList[i].strip())
                else:#无对应翻译
                    if CNName in self.ExistDataDict:
                        pass
                    else:
                        self.youdaoTranslate(CNName)
                            
                    if CNUnit not in self.UNIT_dict:
                        self.UpdataLog(self.FileType+" 新单位出现："+key)
                        self.UNIT_dict[CNUnit] = CNUnit
                    CNSubValueList = CNSubValue.split("|")
                    for CNOneSubValue in CNSubValueList:
                        if CNOneSubValue.find("%")!=-1:#特殊字符不翻译
                            continue
                        elif CNOneSubValue in self.ExistDataDict:
                            continue
                        else:
                            self.youdaoTranslate(CNOneSubValue)
            elif self.FileType=="TEXT.txt" or self.FileType=="DTC.txt" or self.FileType=="ROOT.txt" or self.FileType=="DTC_H.txt":
                CNValue = re.sub(MyRe,"",value).strip()
                ENValue = ""
                if key in Targe_dict:
                    ENValue = re.sub(MyRe,"",Targe_dict[key]).strip()
                    self.AddValue2Dict(CNValue,ENValue)
                else:
                    self.youdaoTranslate(CNValue)

    def checkChar(self,text):
        if self.TargeLanguage=="EN-英语":
            if len(text)==0:
                return 2
            for i in range(0,len(text)):
                if 0<ord (text[i]) and ord (text[i])>127:#英语字符范围0-127
                    #print(text[i])
                    return 0
            return 1
        return 0

    def youdaoTranslate(self,SorseText,xls=""):
        ENText = ""
        if SorseText in self.CheckedDataDict:
            ENText = self.CheckedDataDict[SorseText]
            if ENText!="":
                return ENText
        if SorseText in self.ExistDataDict:
            ENText = self.ExistDataDict[SorseText]
            if ENText!="":
                return ENText
        if SorseText in self.YodaoDataDict:
            ENText = self.YodaoDataDict[SorseText]
            if ENText!="":
                #self.UpdataLog("已有翻译："+SorseText+"->"+ENText)
                return ENText
        
        if self.checkChar(SorseText)==1 or len(SorseText)==0:
            self.UpdataLog("实时翻译："+SorseText+"->"+SorseText)
            self.YodaoDataDict[SorseText]=SorseText
            return SorseText
        ENText = self.NoTranFlah
        if self.Toyoudao:
            #ENText = self.Mytran(SorseText)
            #if ENText!=SorseText:
            #    self.YodaoDataDict[SorseText]=ENText
            #else:
            #    self.NOTranDataDict[SorseText]=self.FileType
            ENText = gooletranslate(SorseText,target_language = self.TargeLanguage)
            self.UpdataLog("实时翻译："+SorseText+"->"+ENText)
            self.YodaoDataDict[SorseText]=ENText
        else:
            self.NOTranDataDict[SorseText]= self.FileType+" " +xls
        return ENText

    def Mytran(self,cn):
        #print(pyautogui.position())
        #time.sleep(100)
        for i in range(0,1):
            i=0
            en = cn
            MyClipboard.copy(cn)
            pyautogui.moveTo(171,1061)
            pyautogui.click()
            pyautogui.moveTo(174,659)
            pyautogui.click()
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5+i)#翻译延迟
            pyautogui.moveTo(442,1014)
            pyautogui.click()
            time.sleep(0.2+i)#翻译延迟
            en = MyClipboard.paste()
            pyautogui.moveTo(171,1061)
            pyautogui.click()
            if en!=cn:
                return en
            i= i + 1
        return cn

    def getTargeValue(self,CN):
        EN = ""
        if CN.find("include")!=-1:
            return CN
        if CN in self.CheckedDataDict:
            EN = self.CheckedDataDict[CN]
            if len(EN)==0:
                if CN in self.ExistDataDict:
                    EN = self.ExistDataDict[CN]
                    if len(EN)==0:
                        if CN in self.YodaoDataDict:
                            EN = self.YodaoDataDict[CN]
        if len(EN)==0:
            EN = self.youdaoTranslate(CN)
        if len(EN)==0:
            EN = self.NoTranFlah
        return EN

    def mylog(self):
        current_time = self.get_current_time().strip().replace(" ","_").replace("-","_").replace(":","_")+".txt"
        folder = os.path.exists("Log")
        if not folder:
            os.makedirs("Log")
        with open("Log//"+current_time,'w',encoding = "utf-8")as lgwrite:
            for line in self.logInfo:
                lgwrite.write(line)
        self.logInfo.clear()

    def CheckReadDtc(self):
        CN_sheet = self.CN_work.sheet_by_name("ReadDtc")
        CN_rows = CN_sheet.nrows
        for row in range(0,CN_rows):
            list5 = str(CN_sheet.cell_value(row,5)).replace(".0","").strip()
            list6 = str(CN_sheet.cell_value(row,6)).replace(".0","").strip()
            list7 = str(CN_sheet.cell_value(row,7)).replace(".0","").strip()
            list8 = str(CN_sheet.cell_value(row,8)).replace(".0","").strip()
            if list5=="3" and list6=="4" and list8!="4" and list7.find("[2]")!=-1:
                self.UpdataLog("存在问题："+self.CNPath)
                break

    def get_current_time(self):
        current_time = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        return current_time

    def CheckGarbledCode(self):
        self.strOutPath="36h.txt"
        self.LangueSelect.update()
        self.TargeLanguage = self.LangueSelect.get()
        if self.TargeLanguage=="EN-英语":
            self.creatFile()
            self.xls2dict()
            Warn = "存在异常："
            i=1
            for k,v in self.CheckedDataDict.items():
                if i:
                    i=0
                    continue
                if self.checkChar(v)!=1 and len(k)!=0:
                    self.UpdataLog(Warn+k+"->"+v)
                    
            i=1
            for k,v in self.ExistDataDict.items():
                if i:
                    i=0
                    continue
                if self.checkChar(v)!=1 and len(k)!=0:
                    self.UpdataLog(Warn+k+"->"+v)
            i=1
            for k,v in self.YodaoDataDict.items():
                if i:
                    i=0
                    continue
                if self.checkChar(v)!=1 and len(k)!=0:
                    self.UpdataLog(Warn+k+"->"+v)
            i=1
            for k,v in self.UNIT_dict.items():
                if i:
                    i=0
                    continue
                if self.checkChar(v)==0:
                    self.UpdataLog("单位乱码："+k+"->"+v)
        else:
            self.UpdataLog("未支持该语言检查")

    def Compare(self):
        strPath = r"D:\内网通接收\2024-09-09\ZJLPZY(2).xls"
        mywork = xlrd.open_workbook(strPath)
        sheet = mywork.sheet_by_name("已经校验翻译")
        rows = sheet.nrows
        CheckedDataDict = {}
        for row in range(1,rows):
            CheckedDataDict[sheet.cell_value(row,0)]=""
            
        sheet = mywork.sheet_by_name("现有翻译")
        rows = sheet.nrows
        ExistDataDict = {}
        for row in range(1,rows):
            ExistDataDict[sheet.cell_value(row,0)]=""
        with open(r"词条差异.txt",r"w",encoding = "utf-8") as Targe_Write:
            for k,v in ExistDataDict.items():
                if k not in CheckedDataDict:
                    Targe_Write.write(k+"\n")
    
    def CheckFormatStart(self):
        if self.strVehicel=="":
            return 0
        self.ObjSelect.update()
        if self.ObjSelect.get()=="国产翻译":
            FileTypeList = ["DS.txt","DTC_H.txt","DTC.txt"]
        else:
            FileTypeList = [self.ObjSelect.get()]
        
        for FileType in FileTypeList:
            self.FileType = FileType
            self.CheckFun()
        self.UpdataLog("********************************格式调整完成，建议用Beyond Compare对比检查一下********************************")
        self.mylog()

    def CheckFun(self):
        self.UpdataLog("*******************************************************开始检查："+self.FileType)
        self.CNFile = self.DataPath+"CN_"+self.FileType
        self.CNFileNew = self.DataPath+"CN_"+"NEW_"+self.FileType
        self.CheckFormat()
        self.UpdataLog("*******************************************************检查完成："+self.FileType+"\n")

    def CheckFormat(self):
        try:
            with open(self.CNFile,r"r",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Read:
                with open(self.CNFileNew,r"w",encoding = self.encodingDict["CN-简体"], errors='replace') as CN_Write:
                    CN_Lines = CN_Read.readlines()
                    for line in CN_Lines:
                        #line = self.Myfilter(line)
                        line = line.strip()
                        if line.find("include")!=-1 or line=="" or line=="//":
                            CN_Write.write(line + "\n")
                            continue
                        CheckLine = process_string(line)
                        if has_invalid_characters(CheckLine, self.encodingDict["CN-简体"]):
                            self.UpdataLog(CheckLine+" 存在乱码")
                        CN_Write.write(CheckLine + "\n")
        except Exception as e:
            self.UpdataLog(line+" 运行错误"+ str(e))

def contains_chinese(text):
    # 定义正则表达式，匹配中文字符
    pattern = re.compile(r'[\u4e00-\u9fa5]')
    return bool(pattern.search(text))

def process_string(input_string):
    # 定义一个函数用来处理引号中的内容
    def clean_quoted_content(match):
        # 获取引号中的内容
        content = match.group(1)
        # 去除首尾空格，并将连续的空格替换为一个空格
        cleaned_content = re.sub(r'\s+', ' ', content.strip())
        return f'"{cleaned_content}"'  # 重新加上引号
    
    # 将连续的空格（两个或更多）替换为一个制表符
    processed_string = re.sub(r' {2,}', '\t', input_string)

    # 将连续的制表符替换为一个制表符
    processed_string = re.sub(r'\t+', '\t', processed_string)

    # 使用正则表达式匹配引号中的内容并进行处理
    processed_string = re.sub(r'"([^"]*)"', clean_quoted_content, processed_string)
    
    # 更新引号内部的内容，将符号","替换为制表符
    processed_string = re.sub(r'"\s*,\s*"', '"\t"', processed_string)

    # 更新引号内部的内容，将符号" "替换为制表符
    processed_string = re.sub(r'"\s* \s*"', '"\t"', processed_string)

    # 将处理后的字符串按制表符切割
    parts = processed_string.split('\t')
    
    # 去除每部分的前后空白字符
    cleaned_parts = [part.strip() for part in parts]

    # 去除首尾的逗号
    cleaned_parts = [part.strip(',') for part in parts]
    
    # 重新用制表符连接这些部分
    return '\t'.join(cleaned_parts)

def has_invalid_characters(input_string, encoding='utf-8'):
    try:
        # 尝试将字符串编码为指定的编码格式
        input_string.encode(encoding)
    except UnicodeEncodeError:
        # 如果抛出 UnicodeEncodeError，表示存在不支持的字符
        return True
    return False

def find_char_positions(s, char):
    positions = []
    for index, c in enumerate(s):
        if c == char:
            positions.append(index)
    return positions

def remove_consecutive_spaces_and_tabs(input_string):
    # 使用正则表达式替换连续的空格或制表符
    # ' + '表示一个或多个空格，'\t+'表示一个或多个制表符
    result = re.sub(r'[ \t]+', ' ', input_string)  # 将多个空格或制表符替换为一个空格
    return result.strip()  # 去掉字符串两端的空白字符

if __name__=='__main__':
    translator = myTranslator()  # 实例化 myTranslator 类
    translator.creatUI()  # 创建并显示用户界面
    translator.MyGUI.mainloop()  # 进入消息循环
