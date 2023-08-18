Dim xlApp,xlFile,xlsheet
Dim YDHXGUID
Dim checkbs
ZDCode = "9130223"
Dim ZDID
Dim id0,id1,id2
#include ".\function\Encryption.vbs"
Dim  Registrationkeyidlist,HardID,usbkeyidlist,usbkeyid

Sub OnClick()
    RegistrationMode = SSProcess.ReadEpsGlobalIni("SoftRegister", "Mode" , "")
    
    If   RegistrationMode = 1 Then
        RegistrationMode1
        If Registrationkeyidlist = Replace(Registrationkeyidlist,HardID,"") Or HardID = "0"  Then  MsgBox "软件未正常授权，请确认注册是否正确！"
        Exit Sub
    ElseIf RegistrationMode = 2 Then
        RegistrationMode2
        If usbkeyidlist = Replace(usbkeyidlist,usbkeyid,"") Or usbkeyid = 0  Then    MsgBox "软件未正常授权，请确认注册是否正确！"
        Exit Sub
    Else
        MsgBox "软件未正常授权，请确认注册是否正确！"
        Exit Sub
    End If
    
    mapHandle = SSProject.GetActiveMap
    mapType = SSProject.GetMapInfo(mapHandle, "MapType")
    If mapType <> 2 Then
        MsgBox "本功能只支持在地形图窗口执行！"
        Exit Sub
    End If
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "=", ZDCode
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount
    If geoCount = 0 Then
        GetZDID = 0
        Exit Sub
    ElseIf geoCount = 1 Then
        ZDID = SSProcess.GetSelGeoValue (0, "SSObj_ID")
        YDHXGUID = SSProcess.GetSelGeoValue (0, "[YDHXGUID]")
        FeatureGUID = SSProcess.GetSelGeoValue (0, "[FeatureGUID]")
        If YDHXGUID = "{00000000-0000-0000-0000-000000000000}"  Then  SSProcess.SetObjectAttr ZDID, "[YDHXGUID]", FeatureGUID
        YDHXGUID = FeatureGUID
    Else
        MsgBox "图上有多个地!"
        Exit Sub
    End If
    
    aa = MsgBox("将覆盖已有数据，是否导入信息？",4 + 64)'是6 否7
    If aa = 7 Then  Exit Sub
    
    EXCELFILE = SSProcess.SelectFileName(1,"选择excel文件",0,"EXCEL Files(*.xls)|*.xls|EXCEL Files(*.xlsx)|*.xlsx|All Files (*.*)|*.*||")
    If EXCELFILE = "" Then Exit Sub
    Set xlApp = CreateObject("Excel.Application")
    Set xlFile = xlApp.Workbooks.Open(EXCELFILE)
    'YDHX()
    CheckExcel()
    
    If checkbs <> ""  Then MsgBox checkbs & "表中存在非法数值，请检查！"
    Exit Sub
    
    GHXK()
    
    DTXXLR()
    
    GNQXXLR()
    
    DTMJZZB()
    
    CXXLR()
    xlApp.Quit
    
    getsygn()
    MsgBox "调入完成，请在”规划指标信息管理“中查看！"
    SSProcess.MapMethod "clearattrbuffer", "ZD_宗地基本信息属性表"
End Sub

Function CheckExcel()
    checkbs = ""
    Set xlsheet = xlFile.Worksheets("工程信息")
    xlsheet.Activate
    ii = 10
    str = Replace(  xlApp.Cells(ii,1),"'","")
    While str <> ""
        str1 = Replace(  xlApp.Cells(ii,2),"'","")
        If str1 = "" Then
            'checkbs1="工程信息"
        Else
            If IsNumeric(str1) = False  Then checkbs1 = "工程信息"
        End If
        ii = ii + 1
        str = Replace(  xlApp.Cells(ii,1),"'","")
    WEnd ifcheckbs1 <> ""  Then checkbs = checkbs1
    
    Set xlsheet = xlFile.Worksheets("单体基本信息")
    xlsheet.Activate
    ii = 2
    str = Replace(  xlApp.Cells(ii,1),"'","")
    While str <> ""
        For j = 3 To 9
            str1 = Replace(  xlApp.Cells(ii,j),"'","")
            
            If str1 = "" Then
                'checkbs2="单体基本信息"
            Else
                If IsNumeric(str1) = False  Then checkbs2 = "单体基本信息"
            End If
        Next
        ii = ii + 1
        str = Replace(  xlApp.Cells(ii,1),"'","")
    WEnd ifcheckbs2 <> ""  Then
    If checkbs = ""  Then checkbs = checkbs2
Else checkbs = checkbs & "，" & checkbs2
End If
End If

Set xlsheet = xlFile.Worksheets("总体指标")
xlsheet.Activate
ii = 2
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""

str1 = Replace(  xlApp.Cells(ii,2),"'","")
If str1 = "" Then
    'checkbs3="总体指标"
Else
    If IsNumeric(str1) = False  Then checkbs3 = "总体指标"
End If
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd ifcheckbs3 <> ""  Then
If checkbs = ""  Then checkbs = checkbs3
Else checkbs = checkbs & "，" & checkbs3
End If
End If

Set xlsheet = xlFile.Worksheets("单体面积指标")
xlsheet.Activate
ii = 2
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""

str1 = Replace(  xlApp.Cells(ii,3),"'","")
If str1 = "" Then
' checkbs4="单体面积指标"
Else
If IsNumeric(str1) = False  Then checkbs4 = "单体面积指标"
End If
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd ifcheckbs4 <> ""  Then
If checkbs = ""  Then checkbs = checkbs4
Else checkbs = checkbs & "，" & checkbs4
End If
End If
For i = 1 To xlApp.sheets.count
checkbsn = ""
If xlApp.sheets(i).Name <> "工程信息"   And xlApp.sheets(i).Name <> "单体面积指标"  And xlApp.sheets(i).Name <> "总体指标"   And  xlApp.sheets(i).Name <> "单体基本信息"  Then
Set xlsheet = xlFile.Worksheets(xlApp.sheets(i).Name)
xlsheet.Activate
ii = 2
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""
For j = 1 To 4
If j = 1 Or j = 3 Or j = 4 Then
str1 = Replace(  xlApp.Cells(ii,j),"'","")
If str1 = "" Then
    ' checkbsn=xlApp.sheets(i).Name
Else
    If IsNumeric(str1) = False  Then checkbsn = xlApp.sheets(i).Name
End If
End If
Next
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd ifcheckbsn <> ""  Then
If checkbs = ""  Then checkbs = checkbsn
Else checkbs = checkbs & "，" & checkbsn
End If
End If
End If
Next

End Function


Function getsygn()
Dim arRecordList()
SYGNS = ""
ProjectName = SSProcess.GetProjectFileName()
sql = "SELECT DISTINCT GongNLX FROM JG_建筑物单体计容面积指标核实信息属性表"
GetSQLRecordAll ProjectName,sql,arRecordList,RecordListCount
For i = 0 To RecordListCount - 1
If SYGNS = ""  Then
SYGNS = arRecordList(i)
Else
SYGNS = SYGNS & "," & arRecordList(i)
End If
Next
sql = "SELECT DISTINCT GongNLX FROM JG_建筑物单体不计容面积指标核实信息属性表"
GetSQLRecordAll ProjectName,sql,arRecordList,RecordListCount
For i = 0 To RecordListCount - 1
If SYGNS = ""  Then
SYGNS = arRecordList(i)
Else
SYGNS = SYGNS & "," & arRecordList(i)
End If
Next
sql = "SELECT DISTINCT GongNLX FROM JG_建筑物单体建筑面积指标核实信息属性表"
GetSQLRecordAll ProjectName,sql,arRecordList,RecordListCount
For i = 0 To RecordListCount - 1
If SYGNS = ""  Then
SYGNS = arRecordList(i)
Else
SYGNS = SYGNS & "," & arRecordList(i)
End If
Next
SYGN = ""
artemp = Split(SYGNS,",")
For i = 0 To UBound(artemp)
If SYGN = ""  Then
SYGN = "|" & artemp(i) & "|"
Else
If Replace(SYGN,"|" & artemp(i) & "|","") = SYGN  Then
SYGN = SYGN & ",|" & artemp(i) & "|"
End If
End If
Next
SYGN = Replace(SYGN,"|","") & ",公建"
infile = "DefaultValue"
sql = "Select " & infile & " From SourceTableFieldInfoTB Where SourceTableFieldInfoTB.SourceTable='FC_面积块信息属性表' and SourceTableFieldInfoTB.SourceField='SYGN'"
inAttr2 sql,infile,SYGN
End Function

'获取指定sql语句下的  搜索结果
Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
SSProcess.OpenAccessMDB mdbName
iRecordCount =  - 1
sql = StrSqlStatement
'打开记录集
SSProcess.OpenAccessRecordset mdbName, sql
'获取记录总数
RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
If RecordCount > 0 Then
iRecordCount = 0
ReDim arSQLRecord(RecordCount)
'将记录游标移到第一行
SSProcess.AccessMoveFirst mdbName, sql
'浏览记录
While SSProcess.AccessIsEOF (mdbName, sql) = 0
fields = ""
values = ""
'获取当前记录内容
SSProcess.GetAccessRecord mdbName, sql, fields, values
If values <> ""  And values <> "*"  Then
arSQLRecord(iRecordCount) = values                                        '查询记录
iRecordCount = iRecordCount + 1
End If'查询记录数
'移动记录游标
SSProcess.AccessMoveNext mdbName, sql
WEnd
End If
'关闭记录集
SSProcess.CloseAccessRecordset mdbName, sql
SSProcess.CloseAccessMDB mdbName
End Function

'修改表信息
Function inAttr2(sql,infile,invalues)
projectName = SSProcess.GetSysPathName (1) & "温州基础地理标准-17_房产分层图.mdt"
SSProcess.OpenAccessMdb projectName
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount > 0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.ModifyAccessRecord  projectName, sql, infile , invalues'输出到mdb表中
SSProcess.AccessMoveNext projectName, sql
WEnd
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Function


ydClassTableName0 = "ZD_宗地基本信息属性表"
ghClassTableName0 = "JG_建设工程规划许可证信息属性表"
dtClassTableName0 = "JG_建设工程建筑单体信息属性表"
gnqClassTableName0 = "JG_建设工程规划许可计容面积指标核实信息属性表"
gnqClassTableName1 = "JG_建设工程规划许可不计容面积指标核实信息属性表"
gnqClassTableName2 = "JG_建设工程规划许可建筑面积指标核实信息属性表"
cClassTableName0 = "JG_建筑物单体楼层高度规划指标核实信息属性表"
dtmjClassTableName0 = "JG_建筑物单体不计容面积指标核实信息属性表"
dtmjClassTableName1 = "JG_建筑物单体计容面积指标核实信息属性表"
dtmjClassTableName2 = "JG_建筑物单体建筑面积指标核实信息属性表"


'用地信息
Function YDHX()
'读excel数据到变量里
Set xlsheet = xlFile.Worksheets("用地信息")
xlsheet.Activate
ydxx = ""
For i = 1 To 20
str = Replace( xlApp.Cells(2,i),"'","")
If ydxx = "" Then
ydxx = str
Else
ydxx = ydxx & "," & str
End If
Next
infile = "GuiHYDXKZBH,XiangMMC,XiangMDZ,JianSDW,SheJDW,WeiTDW,PZMJ,GuiHSPZJZMJ,GuiHSPDSJZMJ,GuiHSPDXJZMJ,GuiHSPZDMJ,GuiHSPRJL,GuiHSPJZMD,GuiHSPLHL,GuiHSPLDMJ,GuiHSPDSJTCWSL,GuiHSPDXJTCWSL,GuiHSPDSFJTCWSL,GuiHSPDXFJTCWSL,GuiHSPZTS"

sql = "Select " & infile & " From " & ydClassTableName0 & " Where " & ydClassTableName0 & ".YDHXGUID =" & YDHXGUID
inAttr sql,infile,ydxx
'while str<>""
'str = xlApp.Cells(h,1)
'wend
End Function

'工程信息

Dim    dtglxx
Dim   GuiHXKZBH

Function GHXK()
Set xlsheet = xlFile.Worksheets("工程信息")
xlsheet.Activate
id0 = getmaxid (ghClassTableName0)

ii = 1
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""
If str = "工程规划许可证编号"  Then GuiHXKZBH = xlApp.Cells(ii,2)
If str = "项目名称"  Then XiangMMC = xlApp.Cells(ii,2)
If str = "工程编号"  Then GongCBH = xlApp.Cells(ii,2)
If str = "项目地址"  Then XiangMDZ = xlApp.Cells(ii,2)
If str = "建设单位"  Then JianSDW = xlApp.Cells(ii,2)
If str = "设计单位"  Then SheJDW = xlApp.Cells(ii,2)
If str = "委托单位"  Then WeiTDW = xlApp.Cells(ii,2)
If str = "审批单位"  Then ShenPDW = xlApp.Cells(ii,2)
If str = "审批时间"  Then ShenPSJ = xlApp.Cells(ii,2)
If str = "规划总用地面积"  Then PZMJ = xlApp.Cells(ii,2)
If PZMJ = ""  Then PZMJ = 0
If str = "规划总建筑面积"  Then GuiHSPZJZMJ = xlApp.Cells(ii,2)
If GuiHSPZJZMJ = ""  Then GuiHSPZJZMJ = 0
If str = "规划地上建筑面积"  Then GuiHSPDSJZMJ = xlApp.Cells(ii,2)
If GuiHSPDSJZMJ = ""  Then GuiHSPDSJZMJ = 0
If str = "规划地下建筑面积"  Then GuiHSPDXJZMJ = xlApp.Cells(ii,2)
If GuiHSPDXJZMJ = ""  Then GuiHSPDXJZMJ = 0
If str = "规划建筑占地面积"  Then GuiHSPZDMJ = xlApp.Cells(ii,2)
If GuiHSPZDMJ = ""  Then GuiHSPZDMJ = 0
If str = "规划容积率"  Then GuiHSPRJL = xlApp.Cells(ii,2)
If GuiHSPRJL = ""  Then GuiHSPRJL = 0
If str = "规划绿化率"  Then GuiHSPLHL = xlApp.Cells(ii,2)
If GuiHSPLHL = ""  Then GuiHSPLHL = 0
If str = "规划建筑密度"  Then GuiHSPJZMD = xlApp.Cells(ii,2)
If GuiHSPJZMD = ""  Then GuiHSPJZMD = 0
If str = "住宅总户数"  Then GuiHSPZTS = xlApp.Cells(ii,2)
If GuiHSPZTS = ""  Then GuiHSPZTS = 0
If str = "规划绿化面积"  Then GuiHSPLDMJ = xlApp.Cells(ii,2)
If GuiHSPLDMJ = ""  Then GuiHSPLDMJ = 0
If str = "规划地上机动停车位"  Then GuiHSPDSJTCWSL = xlApp.Cells(ii,2)
If GuiHSPDSJTCWSL = ""  Then GuiHSPDSJTCWSL = 0
If str = "规划地下机动停车位"  Then GuiHSPDXJTCWSL = xlApp.Cells(ii,2)
If GuiHSPDXJTCWSL = ""  Then GuiHSPDXJTCWSL = 0
If str = "规划地上非机动停车位"  Then GuiHSPDSFJTCWSL = xlApp.Cells(ii,2)
If GuiHSPDSFJTCWSL = ""  Then GuiHSPDSFJTCWSL = 0
If str = "规划地下非机动停车位"  Then GuiHSPDXFJTCWSL = xlApp.Cells(ii,2)
If GuiHSPDXFJTCWSL = ""  Then GuiHSPDXFJTCWSL = 0
If str = "规划总计容面积"  Then GuiHSPZJRMJ = xlApp.Cells(ii,2)
If GuiHSPZJRMJ = ""  Then GuiHSPZJRMJ = 0
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd ifGuiHXKZBH <> ""  Then
scbxx ghClassTableName0,"GuiHXKZBH",GuiHXKZBH
FeatureGUID = GenNewGUID'获取新的featureguid
id0 = 1 + id0
dtglxx = FeatureGUID & "," & YDHXGUID & ",'" & GuiHXKZBH

fields = "ID,FeatureGUID,YDHXGUID,JSGHXKZGUID,GuiHXKZBH,XiangMMC,XiangMDZ,JianSDW,GuiHSPZJZMJ,GuiHSPDSJZMJ,GuiHSPDXJZMJ,GuiHSPZDMJ,GuiHSPRJL,GuiHSPJZMD,GuiHSPLHL,GuiHSPLDMJ,GuiHSPDSJTCWSL,GuiHSPDXJTCWSL,GuiHSPDSFJTCWSL,GuiHSPDXFJTCWSL,GuiHSPZTS,GuiHSPZJRMJ"'
values = id0 & "," & FeatureGUID & "," & YDHXGUID & "," & FeatureGUID & ",'" & GuiHXKZBH & "','" & XiangMMC & "','" & XiangMDZ & "','" & JianSDW & "'," & GuiHSPZJZMJ & "," & GuiHSPDSJZMJ & "," & GuiHSPDXJZMJ & "," & GuiHSPZDMJ & "," & GuiHSPRJL & "," & GuiHSPJZMD & "," & GuiHSPLHL & "," & GuiHSPLDMJ & "," & GuiHSPDSJTCWSL & "," & GuiHSPDXJTCWSL & "," & GuiHSPDSFJTCWSL & "," & GuiHSPDXFJTCWSL & "," & GuiHSPZTS & "," & GuiHSPZJRMJ
InsertRecord ghClassTableName0, fields, values

SSProcess.SetObjectAttr ZDID, "[XiangMMC],[XiangMDZ],[JianSDW],[ShenPDW],[ShenPSJ],[GuiHSPZJZMJ],[GuiHSPDSJZMJ],[GuiHSPDXJZMJ],[GuiHSPZDMJ],[GuiHSPRJL],[GuiHSPJZMD],[GuiHSPLHL],[GuiHSPLDMJ],[GuiHSPDSJTCWSL],[GuiHSPDXJTCWSL],[GuiHSPDSFJTCWSL],[GuiHSPDXFJTCWSL],[GuiHSPZTS],[SheJDW],[WeiTDW],[GongCBH],[PZMJ],[GuiHSPZJRMJ]",XiangMMC & "," & XiangMDZ & "," & JianSDW & "," & ShenPDW & "," & ShenPSJ & "," & GuiHSPZJZMJ & "," & GuiHSPDSJZMJ & "," & GuiHSPDXJZMJ & "," & GuiHSPZDMJ & "," & GuiHSPRJL & "," & GuiHSPJZMD & "," & GuiHSPLHL & "," & GuiHSPLDMJ & "," & GuiHSPDSJTCWSL & "," & GuiHSPDXJTCWSL & "," & GuiHSPDSFJTCWSL & "," & GuiHSPDXFJTCWSL & "," & GuiHSPZTS & "," & SheJDW & "," & WeiTDW & "," & GongCBH & "," & PZMJ & "," & GuiHSPZJRMJ

End If

End Function
'单体信息
Dim  gnvlcglxx(1000),dtcount
Function DTXXLR()

If GuiHXKZBH <> "" Then
'infile="JSGHXKZGUID,YDHXGUID,GuiHYDXKZBH,GuiHXKZBH"
'sql = "Select "&infile&" From "&dtClassTableName0&" Where "&dtClassTableName0&".GuiHXKZBH ='"&GuiHXKZBH&"'"
'inAttr sql,infile,replace(dtglxx,"'","")
scbxx dtClassTableName0,"GuiHXKZBH",GuiHXKZBH
End If


Set xlsheet = xlFile.Worksheets("单体基本信息")
xlsheet.Activate
id0 = getmaxid (dtClassTableName0)
'getmaxid dtClassTableName0,id0
ii = 2
dtcount = 0
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""
gcxx = ""

jzwmc = ",'" & Replace(  xlApp.Cells(ii,1),"'","") & "'"
For i = 1 To 11
str = Replace(   xlApp.Cells(ii,i) ,"'","")
If str = ""  Then str = 0
'if i<6 then
str = "'" & str & "'"
'end if
If gcxx = "" Then
gcxx = str
Else
gcxx = gcxx & "," & str
End If
Next

If    dtglxx <> ""  Then
scbxxdt Replace(  xlApp.Cells(ii,1),"'",""),Replace(  xlApp.Cells(ii,2),"'","")
FeatureGUID = GenNewGUID'获取新的featureguid
id0 = 1 + id0
gnvlcglxx(ii - 2) = dtglxx & "'," & FeatureGUID & jzwmc
dtcount = dtcount + 1
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,JZWMCGUID,JianZWMC,GuiHSPJGLX,GuiHSPZFL,JunGCLZFL,GuiHSPZGD,JunGCLZGD,GuiHSPDSCS,GuiHSPDXCS,GuiHSPZDMJ,FenQH,BeiZ"
values = id0 & "," & dtglxx & "'," & FeatureGUID & "," & FeatureGUID & "," & gcxx

InsertRecord dtClassTableName0, fields, values
End If
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd
End Function

'功能区信息录入
Function GNQXXLR()

If GuiHXKZBH <> "" Then
scbxx gnqClassTableName0,"GuiHXKZBH",GuiHXKZBH
scbxx gnqClassTableName1,"GuiHXKZBH",GuiHXKZBH
scbxx gnqClassTableName2,"GuiHXKZBH",GuiHXKZBH
End If

Set xlsheet = xlFile.Worksheets("总体指标")
xlsheet.Activate
id0 = getmaxid(gnqClassTableName0)
'getmaxid gnqClassTableName0,id0
id1 = getmaxid(gnqClassTableName1)
'getmaxid gnqClassTableName1,id1
id2 = getmaxid(gnqClassTableName2)
'getmaxid gnqClassTableName2,id2
ii = 2
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""
gcxx = ""
xxgljl = ""
ghxkzh = ",'" & Replace(  xlApp.Cells(ii,1),"'","") & "'"
For i = 1 To 2
str = Replace( xlApp.Cells(ii,i) ,"'","")
If str = ""  Then str = 0
If i < 2 Then
str = "'" & str & "'"
End If
If gcxx = "" Then
gcxx = str
Else
gcxx = gcxx & "," & str
End If
If i = 1 Then  gcxx = gcxx & "," & str
If i = 2 Then  gcxx = gcxx & "," & str
Next

If    dtglxx <> ""   Then
FeatureGUID = GenNewGUID'获取新的featureguid
If Replace( xlApp.Cells(ii,3) ,"'","") = "是"  Then
id0 = 1 + id0
If Replace( xlApp.Cells(ii,4) ,"'","") = "地上"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPJRMJ,GuiHSPDSJRMJ"
ElseIf Replace( xlApp.Cells(ii,4) ,"'","") = "地下"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPJRMJ,GuiHSPDXJRMJ"
End If
values = id0 & "," & dtglxx & "'," & FeatureGUID & "," & gcxx
InsertRecord gnqClassTableName0, fields, values
End If
If Replace( xlApp.Cells(ii,3) ,"'","") = "否"  Then
id1 = 1 + id1
If Replace( xlApp.Cells(ii,4) ,"'","") = "地上"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPBJRMJ,GuiHSPDSBJRMJ"
ElseIf Replace( xlApp.Cells(ii,4) ,"'","") = "地下"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPBJRMJ,GuiHSPDXBJRMJ"
End If
values = id0 & "," & dtglxx & "'," & FeatureGUID & "," & gcxx
InsertRecord gnqClassTableName1, fields, values
End If
If Replace( xlApp.Cells(ii,3) ,"'","") <> "否"  And  Replace( xlApp.Cells(ii,3) ,"'","") <> "是" Then
id2 = 1 + id2
If Replace( xlApp.Cells(ii,4) ,"'","") = "地上"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPJZMJ,GuiHSPDSJZMJ"
ElseIf Replace( xlApp.Cells(ii,4) ,"'","") = "地下"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPJZMJ,GuiHSPDXJZMJ"
End If
values = id0 & "," & dtglxx & "'," & FeatureGUID & "," & gcxx
InsertRecord gnqClassTableName2, fields, values
End If
End If
ii = ii + 1
str = Replace( xlApp.Cells(ii,1) ,"'","")
WEnd
End Function

'单体面积指标
Function DTMJZZB()
Set xlsheet = xlFile.Worksheets("单体面积指标")
xlsheet.Activate
ii = 2
dtnames = ""
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""
If dtnames = ""  Then
dtnames = "'" & str & "'"
Else
dtnames = dtnames & ",'" & str & "'"
End If
ii = ii + 1
str = Replace( xlApp.Cells(ii,1) ,"'","")
WEnd ifdtnames <> ""  Then
scbxxdtzb dtmjClassTableName0,"JianZWMC",dtnames
scbxxdtzb dtmjClassTableName1,"JianZWMC",dtnames
scbxxdtzb dtmjClassTableName2,"JianZWMC",dtnames
id0 = getmaxid(dtmjClassTableName1)
id1 = getmaxid(dtmjClassTableName0)
id2 = getmaxid(dtmjClassTableName2)

ii = 2
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""
dtzbglxx = ""
gcxx = ""
jzwmc = "'" & Replace(  xlApp.Cells(ii,1),"'","") & "'"
For i = 2 To 3
str = Replace( xlApp.Cells(ii,i) ,"'","")
If str = ""  Then str = 0
If i < 3 Then
str = "'" & str & "'"
End If
If gcxx = "" Then
gcxx = str
Else
gcxx = gcxx & "," & str
End If
If i = 2 Then  gcxx = gcxx & "," & str
If i = 3 Then  gcxx = gcxx & "," & str
Next
For j = 0 To dtcount - 1
If InStr(gnvlcglxx(j),jzwmc) > 0 Then  dtzbglxx = gnvlcglxx(j)
Exit For
Next

If    dtzbglxx <> ""  Then
FeatureGUID = GenNewGUID'获取新的featureguid
If Replace( xlApp.Cells(ii,4) ,"'","") = "是"  Then
id0 = 1 + id0
If Replace( xlApp.Cells(ii,5) ,"'","") = "地上"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPJRMJ,GuiHSPDSJRMJ"
ElseIf Replace( xlApp.Cells(ii,5) ,"'","") = "地下"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPJRMJ,GuiHSPDXJRMJ"
End If
values = id0 & "," & dtzbglxx & "," & FeatureGUID & "," & gcxx
InsertRecord dtmjClassTableName1, fields, values
End If
If Replace( xlApp.Cells(ii,4) ,"'","") = "否"  Then
id1 = 1 + id1
If Replace( xlApp.Cells(ii,5) ,"'","") = "地上"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPBJRMJ,GuiHSPDSBJRMJ"
ElseIf Replace( xlApp.Cells(ii,5) ,"'","") = "地下"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPBJRMJ,GuiHSPDXBJRMJ"
End If
values = id0 & "," & dtzbglxx & "," & FeatureGUID & "," & gcxx
InsertRecord dtmjClassTableName0, fields, values
End If
If Replace( xlApp.Cells(ii,4) ,"'","") <> "否"  And  Replace( xlApp.Cells(ii,4) ,"'","") <> "是" Then
id2 = 1 + id2
If Replace( xlApp.Cells(ii,5) ,"'","") = "地上"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPJZMJ,GuiHSPDSJZMJ"
ElseIf Replace( xlApp.Cells(ii,5) ,"'","") = "地下"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPJZMJ,GuiHSPDXJZMJ"
End If
values = id0 & "," & dtzbglxx & "," & FeatureGUID & "," & gcxx
InsertRecord dtmjClassTableName2, fields, values
End If
End If
ii = ii + 1
str = Replace( xlApp.Cells(ii,1) ,"'","")
WEnd
End If
End Function

'层信息录入
Function CXXLR()
For m = 1 To xlApp.sheets.count
bs = 0
For j = 0 To dtcount - 1
If InStr(gnvlcglxx(j),xlApp.sheets(m).Name) > 0 Then bs = 1
sheetname = xlApp.sheets(m).Name
cglxx = gnvlcglxx(j)
Exit For
Next
If bs = 1 Then
scbxx cClassTableName0,"JianZWMC",sheetname
Set xlsheet = xlFile.Worksheets(sheetname)
xlsheet.Activate
id0 = getmaxid(cClassTableName0)
' getmaxid cClassTableName0,id0
ii = 2
str = Replace( xlApp.Cells(ii,1),"'","")
While str <> ""
gcxx = ""
ch = ""
cm = ""
For i = 1 To 4
str = Replace( xlApp.Cells(ii,i) ,"'","")
If str = ""  Then str = 0
If i = 1 Then
ch = str
ElseIf  i = 2 Then
cm = str
Else
If gcxx = "" Then
gcxx = str
Else
gcxx = gcxx & "," & str
End If
End If
Next
If InStr(ch,"+") = 0  Then
ch = "'" & ch & "'"
cm = "'" & cm & "'"
FeatureGUID = GenNewGUID'获取新的featureguid
id0 = 1 + id0
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,CengH,CengM,JianSXKCG,YanSCLCG"
values = id0 & "," & cglxx & "," & FeatureGUID & "," & ch & "," & cm & "," & gcxx
InsertRecord cClassTableName0, fields, values
Else
artemp = Split(ch,"+")
For j = CDbl(artemp(0)) To CDbl(artemp(1))
ch = "'" & j & "'"
FeatureGUID = GenNewGUID'获取新的featureguid
id0 = 1 + id0
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,CengH,CengM,JianSXKCG,YanSCLCG"
ReplaceNum j,t
NUMtoZW t,cm
cm = "'" & cm & "'"
values = id0 & "," & cglxx & "," & FeatureGUID & "," & j & "," & cm & "," & gcxx
InsertRecord cClassTableName0, fields, values
Next
End If
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd
End If
Next
End Function

Function ReplaceNum(ByVal xh1,ByRef xh2)
xh1 = Replace (xh1,"-","地下")
xh1 = Replace (xh1,"1","一")
xh1 = Replace (xh1,"2","二")
xh1 = Replace (xh1,"3","三")
xh1 = Replace (xh1,"4","四")
xh1 = Replace (xh1,"5","五")
xh1 = Replace (xh1,"6","六")
xh1 = Replace (xh1,"7","七")
xh1 = Replace (xh1,"8","八")
xh1 = Replace (xh1,"9","九")
xh1 = Replace (xh1,"0","零")
xh2 = xh1
End Function
Function NUMtoZW(ByVal xh3,ByRef xh4) '数字转中文
weiS = Array("","十","百","千","万","十")
sffs = 0
If InStr(xh3,"地下") > 0 Then
xh3 = Replace(xh3,"地下","")
sffs = 1
End If
If InStr(xh3,"夹层") > 0 Then
xh3 = Replace(xh3,"夹层","")
sffs = 2
End If
If InStr(xh3,"顶层") > 0 Then
sffs = 3
End If

length = Len(xh3)
xh4 = ""
If Len(xh3) = 2 And Left(xh3,1) = "一" And Right(xh3,1) <> "零" Then
xh4 = "十" & Right(xh3,1) & ""
ElseIf xh3 = "一零" Then
xh4 = "十"
Else
For i = 1 To length
txh1 = Left(xh3,i)
xh11 = Right(txh1,1)
If i <> length Then
txhCheck = Left(xh3,i + 1)
xhCheck = Right(txhCheck,1)
If xh11 = "零" And xhCheck = "零" Then
xh11 = Replace (xh11,"零","")
End If
End If
If xh11 = "零" And i <> length Then
xh4 = xh4 & xh11 '& weiS(length-i)
ElseIf xh11 <> "零" And xh11 <> "" Then
xh4 = xh4 & xh11 & weiS(length - i)
ElseIf xh11 = ""And length > 5 And i = 2 Then
xh4 = xh4 & xh11 & weiS(length - i)
End If
Next
End If
If sffs = 1 Then xh4 = "地下" & xh4
If sffs = 2 Then xh4 = xh4 & "层夹层"
If sffs = 3 Then xh4 = "顶层"
If sffs = 0 Then xh4 = xh4 & "层"
End Function

'修改表信息
Function inAttr(sql,infile,invalues)
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount > 0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.ModifyAccessRecord  projectName, sql, infile , invalues'输出到mdb表中
SSProcess.AccessMoveNext projectName, sql
WEnd
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Function

'********插入新纪录
Function InsertRecord( tableName, fields, values)
sqlString = "insert into " & tableName & " (" & fields & ") values (" & values & ")"
InsertRecord = SSProcess.ExecuteSql (sqlString)
End Function

'取最新FeatureGUID
Function GenNewGUID()
Set TypeLib = CreateObject("Scriptlet.TypeLib")
GenNewGUID = Left(TypeLib.Guid,38)
Set TypeLib = Nothing
End Function

'删除表信息
Function scbxx(tablename,field,value)
sql = "SELECT * FROM " & tablename & " where " & tablename & "." & field & " = '" & value & "';"

mdbName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb mdbName
SSProcess.OpenAccessRecordset mdbName, sql  '打开数据库

While  SSProcess.AccessIsEOF (mdbName, sql) = False
SSProcess.DelAccessRecord mdbName, sql
WEnd
SSProcess.CloseAccessRecordset mdbName, sql '关库
SSProcess.CloseAccessMdb mdbName
End Function

'删除表信息
Function scbxxdt(ghxkzh,jzwmc)
sql = "SELECT * FROM JG_建设工程建筑单体信息属性表 where JG_建设工程建筑单体信息属性表.GuiHXKZBH = '" & ghxkzh & "' and JG_建设工程建筑单体信息属性表.JianZWMC = '" & jzwmc & "';"
mdbName = SSProcess.GetProjectFileName

SSProcess.OpenAccessMdb mdbName
SSProcess.OpenAccessRecordset mdbName, sql  '打开数据库

While  SSProcess.AccessIsEOF (mdbName, sql) = False
SSProcess.DelAccessRecord mdbName, sql
WEnd
SSProcess.CloseAccessRecordset mdbName, sql '关库
SSProcess.CloseAccessMdb mdbName
End Function

'删除表信息
Function scbxxdtzb(tablename,field,value)
sql = "SELECT * FROM " & tablename & " where " & tablename & "." & field & " in (" & value & ");"

mdbName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb mdbName
SSProcess.OpenAccessRecordset mdbName, sql  '打开数据库

While  SSProcess.AccessIsEOF (mdbName, sql) = False
SSProcess.DelAccessRecord mdbName, sql
WEnd
SSProcess.CloseAccessRecordset mdbName, sql '关库
SSProcess.CloseAccessMdb mdbName
End Function

Function getmaxid(tablename)

mdbName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb mdbName
sql = "SELECT Max(" & tablename & ".ID) AS ID之最大值 FROM " & tablename & ";"
SSProcess.OpenAccessRecordset mdbName, sql
SSProcess.GetAccessRecord mdbName,sql,fields,idvalues
If idvalues <> "" Then
getmaxid = idvalues                            '返回最大ID——ZD_建设用地使用权信息表
Else
getmaxid = 0
End If
SSProcess.CloseAccessRecordset mdbName, sql
SSProcess.CloseAccessMdb mdbName
End Function

