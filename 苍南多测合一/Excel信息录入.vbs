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
        If Registrationkeyidlist = Replace(Registrationkeyidlist,HardID,"") Or HardID = "0"  Then  MsgBox "���δ������Ȩ����ȷ��ע���Ƿ���ȷ��"
        Exit Sub
    ElseIf RegistrationMode = 2 Then
        RegistrationMode2
        If usbkeyidlist = Replace(usbkeyidlist,usbkeyid,"") Or usbkeyid = 0  Then    MsgBox "���δ������Ȩ����ȷ��ע���Ƿ���ȷ��"
        Exit Sub
    Else
        MsgBox "���δ������Ȩ����ȷ��ע���Ƿ���ȷ��"
        Exit Sub
    End If
    
    mapHandle = SSProject.GetActiveMap
    mapType = SSProject.GetMapInfo(mapHandle, "MapType")
    If mapType <> 2 Then
        MsgBox "������ֻ֧���ڵ���ͼ����ִ�У�"
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
        MsgBox "ͼ���ж����!"
        Exit Sub
    End If
    
    aa = MsgBox("�������������ݣ��Ƿ�����Ϣ��",4 + 64)'��6 ��7
    If aa = 7 Then  Exit Sub
    
    EXCELFILE = SSProcess.SelectFileName(1,"ѡ��excel�ļ�",0,"EXCEL Files(*.xls)|*.xls|EXCEL Files(*.xlsx)|*.xlsx|All Files (*.*)|*.*||")
    If EXCELFILE = "" Then Exit Sub
    Set xlApp = CreateObject("Excel.Application")
    Set xlFile = xlApp.Workbooks.Open(EXCELFILE)
    'YDHX()
    CheckExcel()
    
    If checkbs <> ""  Then MsgBox checkbs & "���д��ڷǷ���ֵ�����飡"
    Exit Sub
    
    GHXK()
    
    DTXXLR()
    
    GNQXXLR()
    
    DTMJZZB()
    
    CXXLR()
    xlApp.Quit
    
    getsygn()
    MsgBox "������ɣ����ڡ��滮ָ����Ϣ�����в鿴��"
    SSProcess.MapMethod "clearattrbuffer", "ZD_�ڵػ�����Ϣ���Ա�"
End Sub

Function CheckExcel()
    checkbs = ""
    Set xlsheet = xlFile.Worksheets("������Ϣ")
    xlsheet.Activate
    ii = 10
    str = Replace(  xlApp.Cells(ii,1),"'","")
    While str <> ""
        str1 = Replace(  xlApp.Cells(ii,2),"'","")
        If str1 = "" Then
            'checkbs1="������Ϣ"
        Else
            If IsNumeric(str1) = False  Then checkbs1 = "������Ϣ"
        End If
        ii = ii + 1
        str = Replace(  xlApp.Cells(ii,1),"'","")
    WEnd ifcheckbs1 <> ""  Then checkbs = checkbs1
    
    Set xlsheet = xlFile.Worksheets("���������Ϣ")
    xlsheet.Activate
    ii = 2
    str = Replace(  xlApp.Cells(ii,1),"'","")
    While str <> ""
        For j = 3 To 9
            str1 = Replace(  xlApp.Cells(ii,j),"'","")
            
            If str1 = "" Then
                'checkbs2="���������Ϣ"
            Else
                If IsNumeric(str1) = False  Then checkbs2 = "���������Ϣ"
            End If
        Next
        ii = ii + 1
        str = Replace(  xlApp.Cells(ii,1),"'","")
    WEnd ifcheckbs2 <> ""  Then
    If checkbs = ""  Then checkbs = checkbs2
Else checkbs = checkbs & "��" & checkbs2
End If
End If

Set xlsheet = xlFile.Worksheets("����ָ��")
xlsheet.Activate
ii = 2
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""

str1 = Replace(  xlApp.Cells(ii,2),"'","")
If str1 = "" Then
    'checkbs3="����ָ��"
Else
    If IsNumeric(str1) = False  Then checkbs3 = "����ָ��"
End If
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd ifcheckbs3 <> ""  Then
If checkbs = ""  Then checkbs = checkbs3
Else checkbs = checkbs & "��" & checkbs3
End If
End If

Set xlsheet = xlFile.Worksheets("�������ָ��")
xlsheet.Activate
ii = 2
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""

str1 = Replace(  xlApp.Cells(ii,3),"'","")
If str1 = "" Then
' checkbs4="�������ָ��"
Else
If IsNumeric(str1) = False  Then checkbs4 = "�������ָ��"
End If
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd ifcheckbs4 <> ""  Then
If checkbs = ""  Then checkbs = checkbs4
Else checkbs = checkbs & "��" & checkbs4
End If
End If
For i = 1 To xlApp.sheets.count
checkbsn = ""
If xlApp.sheets(i).Name <> "������Ϣ"   And xlApp.sheets(i).Name <> "�������ָ��"  And xlApp.sheets(i).Name <> "����ָ��"   And  xlApp.sheets(i).Name <> "���������Ϣ"  Then
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
Else checkbs = checkbs & "��" & checkbsn
End If
End If
End If
Next

End Function


Function getsygn()
Dim arRecordList()
SYGNS = ""
ProjectName = SSProcess.GetProjectFileName()
sql = "SELECT DISTINCT GongNLX FROM JG_�����ﵥ��������ָ���ʵ��Ϣ���Ա�"
GetSQLRecordAll ProjectName,sql,arRecordList,RecordListCount
For i = 0 To RecordListCount - 1
If SYGNS = ""  Then
SYGNS = arRecordList(i)
Else
SYGNS = SYGNS & "," & arRecordList(i)
End If
Next
sql = "SELECT DISTINCT GongNLX FROM JG_�����ﵥ�岻�������ָ���ʵ��Ϣ���Ա�"
GetSQLRecordAll ProjectName,sql,arRecordList,RecordListCount
For i = 0 To RecordListCount - 1
If SYGNS = ""  Then
SYGNS = arRecordList(i)
Else
SYGNS = SYGNS & "," & arRecordList(i)
End If
Next
sql = "SELECT DISTINCT GongNLX FROM JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�"
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
SYGN = Replace(SYGN,"|","") & ",����"
infile = "DefaultValue"
sql = "Select " & infile & " From SourceTableFieldInfoTB Where SourceTableFieldInfoTB.SourceTable='FC_�������Ϣ���Ա�' and SourceTableFieldInfoTB.SourceField='SYGN'"
inAttr2 sql,infile,SYGN
End Function

'��ȡָ��sql����µ�  �������
Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
SSProcess.OpenAccessMDB mdbName
iRecordCount =  - 1
sql = StrSqlStatement
'�򿪼�¼��
SSProcess.OpenAccessRecordset mdbName, sql
'��ȡ��¼����
RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
If RecordCount > 0 Then
iRecordCount = 0
ReDim arSQLRecord(RecordCount)
'����¼�α��Ƶ���һ��
SSProcess.AccessMoveFirst mdbName, sql
'�����¼
While SSProcess.AccessIsEOF (mdbName, sql) = 0
fields = ""
values = ""
'��ȡ��ǰ��¼����
SSProcess.GetAccessRecord mdbName, sql, fields, values
If values <> ""  And values <> "*"  Then
arSQLRecord(iRecordCount) = values                                        '��ѯ��¼
iRecordCount = iRecordCount + 1
End If'��ѯ��¼��
'�ƶ���¼�α�
SSProcess.AccessMoveNext mdbName, sql
WEnd
End If
'�رռ�¼��
SSProcess.CloseAccessRecordset mdbName, sql
SSProcess.CloseAccessMDB mdbName
End Function

'�޸ı���Ϣ
Function inAttr2(sql,infile,invalues)
projectName = SSProcess.GetSysPathName (1) & "���ݻ��������׼-17_�����ֲ�ͼ.mdt"
SSProcess.OpenAccessMdb projectName
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount > 0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.ModifyAccessRecord  projectName, sql, infile , invalues'�����mdb����
SSProcess.AccessMoveNext projectName, sql
WEnd
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Function


ydClassTableName0 = "ZD_�ڵػ�����Ϣ���Ա�"
ghClassTableName0 = "JG_���蹤�̹滮���֤��Ϣ���Ա�"
dtClassTableName0 = "JG_���蹤�̽���������Ϣ���Ա�"
gnqClassTableName0 = "JG_���蹤�̹滮��ɼ������ָ���ʵ��Ϣ���Ա�"
gnqClassTableName1 = "JG_���蹤�̹滮��ɲ��������ָ���ʵ��Ϣ���Ա�"
gnqClassTableName2 = "JG_���蹤�̹滮��ɽ������ָ���ʵ��Ϣ���Ա�"
cClassTableName0 = "JG_�����ﵥ��¥��߶ȹ滮ָ���ʵ��Ϣ���Ա�"
dtmjClassTableName0 = "JG_�����ﵥ�岻�������ָ���ʵ��Ϣ���Ա�"
dtmjClassTableName1 = "JG_�����ﵥ��������ָ���ʵ��Ϣ���Ա�"
dtmjClassTableName2 = "JG_�����ﵥ�彨�����ָ���ʵ��Ϣ���Ա�"


'�õ���Ϣ
Function YDHX()
'��excel���ݵ�������
Set xlsheet = xlFile.Worksheets("�õ���Ϣ")
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

'������Ϣ

Dim    dtglxx
Dim   GuiHXKZBH

Function GHXK()
Set xlsheet = xlFile.Worksheets("������Ϣ")
xlsheet.Activate
id0 = getmaxid (ghClassTableName0)

ii = 1
str = Replace(  xlApp.Cells(ii,1),"'","")
While str <> ""
If str = "���̹滮���֤���"  Then GuiHXKZBH = xlApp.Cells(ii,2)
If str = "��Ŀ����"  Then XiangMMC = xlApp.Cells(ii,2)
If str = "���̱��"  Then GongCBH = xlApp.Cells(ii,2)
If str = "��Ŀ��ַ"  Then XiangMDZ = xlApp.Cells(ii,2)
If str = "���赥λ"  Then JianSDW = xlApp.Cells(ii,2)
If str = "��Ƶ�λ"  Then SheJDW = xlApp.Cells(ii,2)
If str = "ί�е�λ"  Then WeiTDW = xlApp.Cells(ii,2)
If str = "������λ"  Then ShenPDW = xlApp.Cells(ii,2)
If str = "����ʱ��"  Then ShenPSJ = xlApp.Cells(ii,2)
If str = "�滮���õ����"  Then PZMJ = xlApp.Cells(ii,2)
If PZMJ = ""  Then PZMJ = 0
If str = "�滮�ܽ������"  Then GuiHSPZJZMJ = xlApp.Cells(ii,2)
If GuiHSPZJZMJ = ""  Then GuiHSPZJZMJ = 0
If str = "�滮���Ͻ������"  Then GuiHSPDSJZMJ = xlApp.Cells(ii,2)
If GuiHSPDSJZMJ = ""  Then GuiHSPDSJZMJ = 0
If str = "�滮���½������"  Then GuiHSPDXJZMJ = xlApp.Cells(ii,2)
If GuiHSPDXJZMJ = ""  Then GuiHSPDXJZMJ = 0
If str = "�滮����ռ�����"  Then GuiHSPZDMJ = xlApp.Cells(ii,2)
If GuiHSPZDMJ = ""  Then GuiHSPZDMJ = 0
If str = "�滮�ݻ���"  Then GuiHSPRJL = xlApp.Cells(ii,2)
If GuiHSPRJL = ""  Then GuiHSPRJL = 0
If str = "�滮�̻���"  Then GuiHSPLHL = xlApp.Cells(ii,2)
If GuiHSPLHL = ""  Then GuiHSPLHL = 0
If str = "�滮�����ܶ�"  Then GuiHSPJZMD = xlApp.Cells(ii,2)
If GuiHSPJZMD = ""  Then GuiHSPJZMD = 0
If str = "סլ�ܻ���"  Then GuiHSPZTS = xlApp.Cells(ii,2)
If GuiHSPZTS = ""  Then GuiHSPZTS = 0
If str = "�滮�̻����"  Then GuiHSPLDMJ = xlApp.Cells(ii,2)
If GuiHSPLDMJ = ""  Then GuiHSPLDMJ = 0
If str = "�滮���ϻ���ͣ��λ"  Then GuiHSPDSJTCWSL = xlApp.Cells(ii,2)
If GuiHSPDSJTCWSL = ""  Then GuiHSPDSJTCWSL = 0
If str = "�滮���»���ͣ��λ"  Then GuiHSPDXJTCWSL = xlApp.Cells(ii,2)
If GuiHSPDXJTCWSL = ""  Then GuiHSPDXJTCWSL = 0
If str = "�滮���Ϸǻ���ͣ��λ"  Then GuiHSPDSFJTCWSL = xlApp.Cells(ii,2)
If GuiHSPDSFJTCWSL = ""  Then GuiHSPDSFJTCWSL = 0
If str = "�滮���·ǻ���ͣ��λ"  Then GuiHSPDXFJTCWSL = xlApp.Cells(ii,2)
If GuiHSPDXFJTCWSL = ""  Then GuiHSPDXFJTCWSL = 0
If str = "�滮�ܼ������"  Then GuiHSPZJRMJ = xlApp.Cells(ii,2)
If GuiHSPZJRMJ = ""  Then GuiHSPZJRMJ = 0
ii = ii + 1
str = Replace(  xlApp.Cells(ii,1),"'","")
WEnd ifGuiHXKZBH <> ""  Then
scbxx ghClassTableName0,"GuiHXKZBH",GuiHXKZBH
FeatureGUID = GenNewGUID'��ȡ�µ�featureguid
id0 = 1 + id0
dtglxx = FeatureGUID & "," & YDHXGUID & ",'" & GuiHXKZBH

fields = "ID,FeatureGUID,YDHXGUID,JSGHXKZGUID,GuiHXKZBH,XiangMMC,XiangMDZ,JianSDW,GuiHSPZJZMJ,GuiHSPDSJZMJ,GuiHSPDXJZMJ,GuiHSPZDMJ,GuiHSPRJL,GuiHSPJZMD,GuiHSPLHL,GuiHSPLDMJ,GuiHSPDSJTCWSL,GuiHSPDXJTCWSL,GuiHSPDSFJTCWSL,GuiHSPDXFJTCWSL,GuiHSPZTS,GuiHSPZJRMJ"'
values = id0 & "," & FeatureGUID & "," & YDHXGUID & "," & FeatureGUID & ",'" & GuiHXKZBH & "','" & XiangMMC & "','" & XiangMDZ & "','" & JianSDW & "'," & GuiHSPZJZMJ & "," & GuiHSPDSJZMJ & "," & GuiHSPDXJZMJ & "," & GuiHSPZDMJ & "," & GuiHSPRJL & "," & GuiHSPJZMD & "," & GuiHSPLHL & "," & GuiHSPLDMJ & "," & GuiHSPDSJTCWSL & "," & GuiHSPDXJTCWSL & "," & GuiHSPDSFJTCWSL & "," & GuiHSPDXFJTCWSL & "," & GuiHSPZTS & "," & GuiHSPZJRMJ
InsertRecord ghClassTableName0, fields, values

SSProcess.SetObjectAttr ZDID, "[XiangMMC],[XiangMDZ],[JianSDW],[ShenPDW],[ShenPSJ],[GuiHSPZJZMJ],[GuiHSPDSJZMJ],[GuiHSPDXJZMJ],[GuiHSPZDMJ],[GuiHSPRJL],[GuiHSPJZMD],[GuiHSPLHL],[GuiHSPLDMJ],[GuiHSPDSJTCWSL],[GuiHSPDXJTCWSL],[GuiHSPDSFJTCWSL],[GuiHSPDXFJTCWSL],[GuiHSPZTS],[SheJDW],[WeiTDW],[GongCBH],[PZMJ],[GuiHSPZJRMJ]",XiangMMC & "," & XiangMDZ & "," & JianSDW & "," & ShenPDW & "," & ShenPSJ & "," & GuiHSPZJZMJ & "," & GuiHSPDSJZMJ & "," & GuiHSPDXJZMJ & "," & GuiHSPZDMJ & "," & GuiHSPRJL & "," & GuiHSPJZMD & "," & GuiHSPLHL & "," & GuiHSPLDMJ & "," & GuiHSPDSJTCWSL & "," & GuiHSPDXJTCWSL & "," & GuiHSPDSFJTCWSL & "," & GuiHSPDXFJTCWSL & "," & GuiHSPZTS & "," & SheJDW & "," & WeiTDW & "," & GongCBH & "," & PZMJ & "," & GuiHSPZJRMJ

End If

End Function
'������Ϣ
Dim  gnvlcglxx(1000),dtcount
Function DTXXLR()

If GuiHXKZBH <> "" Then
'infile="JSGHXKZGUID,YDHXGUID,GuiHYDXKZBH,GuiHXKZBH"
'sql = "Select "&infile&" From "&dtClassTableName0&" Where "&dtClassTableName0&".GuiHXKZBH ='"&GuiHXKZBH&"'"
'inAttr sql,infile,replace(dtglxx,"'","")
scbxx dtClassTableName0,"GuiHXKZBH",GuiHXKZBH
End If


Set xlsheet = xlFile.Worksheets("���������Ϣ")
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
FeatureGUID = GenNewGUID'��ȡ�µ�featureguid
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

'��������Ϣ¼��
Function GNQXXLR()

If GuiHXKZBH <> "" Then
scbxx gnqClassTableName0,"GuiHXKZBH",GuiHXKZBH
scbxx gnqClassTableName1,"GuiHXKZBH",GuiHXKZBH
scbxx gnqClassTableName2,"GuiHXKZBH",GuiHXKZBH
End If

Set xlsheet = xlFile.Worksheets("����ָ��")
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
FeatureGUID = GenNewGUID'��ȡ�µ�featureguid
If Replace( xlApp.Cells(ii,3) ,"'","") = "��"  Then
id0 = 1 + id0
If Replace( xlApp.Cells(ii,4) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPJRMJ,GuiHSPDSJRMJ"
ElseIf Replace( xlApp.Cells(ii,4) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPJRMJ,GuiHSPDXJRMJ"
End If
values = id0 & "," & dtglxx & "'," & FeatureGUID & "," & gcxx
InsertRecord gnqClassTableName0, fields, values
End If
If Replace( xlApp.Cells(ii,3) ,"'","") = "��"  Then
id1 = 1 + id1
If Replace( xlApp.Cells(ii,4) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPBJRMJ,GuiHSPDSBJRMJ"
ElseIf Replace( xlApp.Cells(ii,4) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPBJRMJ,GuiHSPDXBJRMJ"
End If
values = id0 & "," & dtglxx & "'," & FeatureGUID & "," & gcxx
InsertRecord gnqClassTableName1, fields, values
End If
If Replace( xlApp.Cells(ii,3) ,"'","") <> "��"  And  Replace( xlApp.Cells(ii,3) ,"'","") <> "��" Then
id2 = 1 + id2
If Replace( xlApp.Cells(ii,4) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,FeatureGUID,GongNLX,GongNMC,GuiHSPJZMJ,GuiHSPDSJZMJ"
ElseIf Replace( xlApp.Cells(ii,4) ,"'","") = "����"   Then
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

'�������ָ��
Function DTMJZZB()
Set xlsheet = xlFile.Worksheets("�������ָ��")
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
FeatureGUID = GenNewGUID'��ȡ�µ�featureguid
If Replace( xlApp.Cells(ii,4) ,"'","") = "��"  Then
id0 = 1 + id0
If Replace( xlApp.Cells(ii,5) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPJRMJ,GuiHSPDSJRMJ"
ElseIf Replace( xlApp.Cells(ii,5) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPJRMJ,GuiHSPDXJRMJ"
End If
values = id0 & "," & dtzbglxx & "," & FeatureGUID & "," & gcxx
InsertRecord dtmjClassTableName1, fields, values
End If
If Replace( xlApp.Cells(ii,4) ,"'","") = "��"  Then
id1 = 1 + id1
If Replace( xlApp.Cells(ii,5) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPBJRMJ,GuiHSPDSBJRMJ"
ElseIf Replace( xlApp.Cells(ii,5) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPBJRMJ,GuiHSPDXBJRMJ"
End If
values = id0 & "," & dtzbglxx & "," & FeatureGUID & "," & gcxx
InsertRecord dtmjClassTableName0, fields, values
End If
If Replace( xlApp.Cells(ii,4) ,"'","") <> "��"  And  Replace( xlApp.Cells(ii,4) ,"'","") <> "��" Then
id2 = 1 + id2
If Replace( xlApp.Cells(ii,5) ,"'","") = "����"   Then
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,GongNLX,GongNMC,GuiHSPJZMJ,GuiHSPDSJZMJ"
ElseIf Replace( xlApp.Cells(ii,5) ,"'","") = "����"   Then
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

'����Ϣ¼��
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
FeatureGUID = GenNewGUID'��ȡ�µ�featureguid
id0 = 1 + id0
fields = "ID,JSGHXKZGUID,YDHXGUID,GuiHXKZBH,JZWMCGUID,JianZWMC,FeatureGUID,CengH,CengM,JianSXKCG,YanSCLCG"
values = id0 & "," & cglxx & "," & FeatureGUID & "," & ch & "," & cm & "," & gcxx
InsertRecord cClassTableName0, fields, values
Else
artemp = Split(ch,"+")
For j = CDbl(artemp(0)) To CDbl(artemp(1))
ch = "'" & j & "'"
FeatureGUID = GenNewGUID'��ȡ�µ�featureguid
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
xh1 = Replace (xh1,"-","����")
xh1 = Replace (xh1,"1","һ")
xh1 = Replace (xh1,"2","��")
xh1 = Replace (xh1,"3","��")
xh1 = Replace (xh1,"4","��")
xh1 = Replace (xh1,"5","��")
xh1 = Replace (xh1,"6","��")
xh1 = Replace (xh1,"7","��")
xh1 = Replace (xh1,"8","��")
xh1 = Replace (xh1,"9","��")
xh1 = Replace (xh1,"0","��")
xh2 = xh1
End Function
Function NUMtoZW(ByVal xh3,ByRef xh4) '����ת����
weiS = Array("","ʮ","��","ǧ","��","ʮ")
sffs = 0
If InStr(xh3,"����") > 0 Then
xh3 = Replace(xh3,"����","")
sffs = 1
End If
If InStr(xh3,"�в�") > 0 Then
xh3 = Replace(xh3,"�в�","")
sffs = 2
End If
If InStr(xh3,"����") > 0 Then
sffs = 3
End If

length = Len(xh3)
xh4 = ""
If Len(xh3) = 2 And Left(xh3,1) = "һ" And Right(xh3,1) <> "��" Then
xh4 = "ʮ" & Right(xh3,1) & ""
ElseIf xh3 = "һ��" Then
xh4 = "ʮ"
Else
For i = 1 To length
txh1 = Left(xh3,i)
xh11 = Right(txh1,1)
If i <> length Then
txhCheck = Left(xh3,i + 1)
xhCheck = Right(txhCheck,1)
If xh11 = "��" And xhCheck = "��" Then
xh11 = Replace (xh11,"��","")
End If
End If
If xh11 = "��" And i <> length Then
xh4 = xh4 & xh11 '& weiS(length-i)
ElseIf xh11 <> "��" And xh11 <> "" Then
xh4 = xh4 & xh11 & weiS(length - i)
ElseIf xh11 = ""And length > 5 And i = 2 Then
xh4 = xh4 & xh11 & weiS(length - i)
End If
Next
End If
If sffs = 1 Then xh4 = "����" & xh4
If sffs = 2 Then xh4 = xh4 & "��в�"
If sffs = 3 Then xh4 = "����"
If sffs = 0 Then xh4 = xh4 & "��"
End Function

'�޸ı���Ϣ
Function inAttr(sql,infile,invalues)
projectName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb projectName
SSProcess.OpenAccessRecordset projectName, sql
rscount = SSProcess.GetAccessRecordCount (projectName, sql)
If rscount > 0 Then
SSProcess.AccessMoveFirst projectName, sql
While (SSProcess.AccessIsEOF (projectName, sql ) = False)
SSProcess.ModifyAccessRecord  projectName, sql, infile , invalues'�����mdb����
SSProcess.AccessMoveNext projectName, sql
WEnd
End If
SSProcess.CloseAccessRecordset projectName, sql
SSProcess.CloseAccessMdb projectName
End Function

'********�����¼�¼
Function InsertRecord( tableName, fields, values)
sqlString = "insert into " & tableName & " (" & fields & ") values (" & values & ")"
InsertRecord = SSProcess.ExecuteSql (sqlString)
End Function

'ȡ����FeatureGUID
Function GenNewGUID()
Set TypeLib = CreateObject("Scriptlet.TypeLib")
GenNewGUID = Left(TypeLib.Guid,38)
Set TypeLib = Nothing
End Function

'ɾ������Ϣ
Function scbxx(tablename,field,value)
sql = "SELECT * FROM " & tablename & " where " & tablename & "." & field & " = '" & value & "';"

mdbName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb mdbName
SSProcess.OpenAccessRecordset mdbName, sql  '�����ݿ�

While  SSProcess.AccessIsEOF (mdbName, sql) = False
SSProcess.DelAccessRecord mdbName, sql
WEnd
SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
SSProcess.CloseAccessMdb mdbName
End Function

'ɾ������Ϣ
Function scbxxdt(ghxkzh,jzwmc)
sql = "SELECT * FROM JG_���蹤�̽���������Ϣ���Ա� where JG_���蹤�̽���������Ϣ���Ա�.GuiHXKZBH = '" & ghxkzh & "' and JG_���蹤�̽���������Ϣ���Ա�.JianZWMC = '" & jzwmc & "';"
mdbName = SSProcess.GetProjectFileName

SSProcess.OpenAccessMdb mdbName
SSProcess.OpenAccessRecordset mdbName, sql  '�����ݿ�

While  SSProcess.AccessIsEOF (mdbName, sql) = False
SSProcess.DelAccessRecord mdbName, sql
WEnd
SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
SSProcess.CloseAccessMdb mdbName
End Function

'ɾ������Ϣ
Function scbxxdtzb(tablename,field,value)
sql = "SELECT * FROM " & tablename & " where " & tablename & "." & field & " in (" & value & ");"

mdbName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb mdbName
SSProcess.OpenAccessRecordset mdbName, sql  '�����ݿ�

While  SSProcess.AccessIsEOF (mdbName, sql) = False
SSProcess.DelAccessRecord mdbName, sql
WEnd
SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
SSProcess.CloseAccessMdb mdbName
End Function

Function getmaxid(tablename)

mdbName = SSProcess.GetProjectFileName
SSProcess.OpenAccessMdb mdbName
sql = "SELECT Max(" & tablename & ".ID) AS ID֮���ֵ FROM " & tablename & ";"
SSProcess.OpenAccessRecordset mdbName, sql
SSProcess.GetAccessRecord mdbName,sql,fields,idvalues
If idvalues <> "" Then
getmaxid = idvalues                            '�������ID����ZD_�����õ�ʹ��Ȩ��Ϣ��
Else
getmaxid = 0
End If
SSProcess.CloseAccessRecordset mdbName, sql
SSProcess.CloseAccessMdb mdbName
End Function

