'========================================================Excel����������ļ�·����������======================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Excel��������
Dim ExcelObj
Set ExcelObj = CreateObject("Excel.Application")

'============================================================����¼����====================================================================

'������ֵ
Dim Threshold
Threshold = 5

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���߾��ȼ��"

'��鼯������
Dim strCheckName
strCheckName = "����ȼ��"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->����ȼ��"

'�������
Dim strDescription
strDescription = "����ȳ���"


'=============================================================�������=======================================================================

Sub OnClick()
    
    AllVisible
    
    JcrInfo DateTime,PersonName
    
    ExcelFilePath = SSProcess.SelectFileName(1,"ѡ��Excel�ļ�",0,"EXCEL Files(*.xlsx)|*.xlsx|EXCEL Files(*.xls)|*.xls|All Files (*.*)|*.*||")
    
    If  ExcelFilePath = "" Then
        MsgBox "δѡ���ļ������˳�"
        Exit Sub
    End If
    
    FileName = Right(ExcelFilePath,Len(ExcelFilePath) - InStrRev(ExcelFilePath,"\"))
    FileSysObj.CopyFile  ExcelFilePath,SSProcess.GetSysPathName(5) & "��������\" & FileName
    
    ClearCheckRecord
    
    OpenExcel SSProcess.GetSysPathName(5) & FileName,ExcelFile
    
    SetTableHeader
    
    InsertExcel ExcelFile,4,GetEndRow - 1,ErrorIds
    
    GetMaxDeep ExcelFile,4,GetEndRow - 1
    
    SetFooter DateTime,PersonName,GetEndRow - 1
    
    CloseExcel ExcelFile
    
    AddRecord ErrorIds
    
    Ending
    
End Sub' OnClick

'=============================================================������ֵ==============================================================

'�������˵���
Function JcrInfo(ByRef DateTime,ByRef PersonName)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "�����" , "" , 0 , "" , ""
    SSProcess.AddInputParameter "����" , GetNowTime , 0 , "" , ""
    Result = SSProcess.ShowInputParameterDlg ("��Ϣ¼��")
    If Result = 1 Then
        DateTime = SSProcess.GetInputParameter ("����")
        PersonName = SSProcess.GetInputParameter ("�����")
    End If
End Function' JcrInfo

'��ȡ��ǰϵͳʱ��
Function GetNowTime()
    GetNowTime = CStr(FormatDateTime(Now(),1))
End Function' GetNowTime

'���ñ���
Function SetTableHeader()
    SqlStr = "Select XMMC From ������Ŀ��Ϣ�� Where ������Ŀ��Ϣ��.ID = 1"
    GetSQLRecordAll SqlStr,Xmmc,Count
    ExcelObj.Cells(2,2) = Xmmc(0)
End Function' SetTableHeader

'��ȡ�������
Function GetEndRow()
    GetEndRow = 4
    Poisition = InStr(ExcelObj.Cells(GetEndRow,1),"���ϲ�Ϊ��")
    Do While ExcelObj.Cells(GetEndRow,1) <> ""
        Poisition = InStr(ExcelObj.Cells(GetEndRow,1),"���ϲ�Ϊ��")
        'MsgBox Poisition
        If Poisition = 0 Then
            GetEndRow = GetEndRow + 1
        Else
            Exit Function
        End If
    Loop
End Function' GetEndRow

Function InsertExcel(ByVal ExcelFile,ByVal StartRow,ByVal EndRow,ByRef ErrorIds)
    ErrorIds = ""
    For i = StartRow To EndRow
        SqlStr = "Select ���¹��������Ա�.ID,���¹��������Ա�.GXQDMS,���¹��������Ա�.GXQDDH From ���¹��������Ա� inner join GeoLineTB on ���¹��������Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0 And GXQDDH = " & AddDenote(ExcelObj.Cells(i,1)) & " And GXZDDH = " & AddDenote(ExcelObj.Cells(i,2))
        GetSQLRecordAll SqlStr,DeepArr,DeepCount
        If DeepCount >= 1 Then
            TempArr = Split(DeepArr(0),",", - 1,1)
            SqlStr = "Select ���¹��ߵ����Ա�.FSW From ���¹��ߵ����Ա� inner join GeoPointTB on ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And ���¹��ߵ����Ա�.WTDH = " & "'" & TempArr(2) & "'"
            GetSQLRecordAll SqlStr,PoiFWSArr,Count
            ExcelObj.Cells(i,3) = Round(Transform(TempArr(1)) * 100,2)
            Diff = Abs(Round(Transform(TempArr(1)) * 100 - Transform(ExcelObj.Cells(i,4)),2))
            ExcelObj.Cells(i,5) = Diff
            If PoiFWSArr(0) = "" Then
                If Diff > 0.15 * Transform(TempArr(1)) * 100 Then
                    ExcelObj.Cells(i,6) = "�����"
                    If ErrorIds = "" Then
                        ErrorIds = TempArr(0)
                    Else
                        ErrorIds = ErrorIds & "," & TempArr(0)
                    End If
                End If
            ElseIf PoiFWSArr(0) = "*" Then
                If Diff > 0.15 * Transform(TempArr(1)) * 100 Then
                    ExcelObj.Cells(i,6) = "�����"
                    If ErrorIds = "" Then
                        ErrorIds = TempArr(0)
                    Else
                        ErrorIds = ErrorIds & "," & TempArr(0)
                    End If
                End If
            ElseIf PoiFWSArr(0) = Null Then
                If Diff > 0.15 * Transform(TempArr(1)) * 100 Then
                    ExcelObj.Cells(i,6) = "�����"
                    If ErrorIds = "" Then
                        ErrorIds = TempArr(0)
                    Else
                        ErrorIds = ErrorIds & "," & TempArr(0)
                    End If
                End If
            Else
                If Diff > Threshold Then
                    ExcelObj.Cells(i,6) = "�����"
                    If ErrorIds = "" Then
                        ErrorIds = TempArr(0)
                    Else
                        ErrorIds = ErrorIds & "," & TempArr(0)
                    End If
                End If
            End If
        End If
    Next 'i
End Function' InsertExcel

'��ȡ�����Ȳ�
Function GetMaxDeep(ByVal ExcelFile,ByVal StartRow,ByVal EndRow)
    GetMaxDeep = 0
    For i = StartRow To EndRow
        If GetMaxDeep < Transform(ExcelObj.Cells(i,5)) Then
            GetMaxDeep = Transform(ExcelObj.Cells(i,5))
        End If
    Next 'i
    ExcelObj.Cells(EndRow + 1,1) = "���ϲ�Ϊ��" & GetMaxDeep & "CM"
    TIANzdjc "MSZDJC",GetMaxDeep
End Function' GetMaxDeep

'��д����˺�����
Function SetFooter(ByVal DateTime,ByVal PersonName,ByVal EndRow)
    ExcelObj.Cells(EndRow + 2,2) = PersonName
    ExcelObj.Cells(EndRow + 2,6) = DateTime
End Function' SetFooter

'��Ӽ���¼
Function AddRecord(ByVal ErrorIds)
    If ErrorIds <> "" Then
        ErrorArr = Split(ErrorIds,",", - 1,1)
        For i = 0 To UBound(ErrorArr)
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(ErrorArr(i),"SSObj_X"),SSProcess.GetObjectAttr(ErrorArr(i),"SSObj_Y"),0,1,ErrorArr(i),""
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' AddRecord

'���������ֵ������Ϣ��
Function TIANzdjc (zdmc,zdjc)
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    sql = "update  ������Ŀ��Ϣ�� set " & zdmc & " = " & zdjc & " where ������Ŀ��Ϣ��.ID= 1"
    SSProcess.ExecuteAccessSql  mdbName,sql
    SSProcess.CloseAccessMdb mdbName
End Function

'==========================================================�����ຯ��====================================================================

'������ͼ��
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'��������ת��
Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

'��Excel��
Function OpenExcel(ByVal FilePath,ByRef ExcleFile)
    ExcelObj.Application.Visible = False
    Set ExcleFile = ExcelObj.WorkBooks.Open(FilePath)
    Set ExcelSheet = ExcleFile.WorkSheets(1)
    ExcelSheet.Activate
End Function

'��ȡExcel����
Function GetExcelName(ByVal ExcelFilePath)
    ExcelFilePathArr = Split(ExcelFilePath,"\", - 1,1)
    GetExcelName = ExcelFilePathArr(UBound(ExcelFilePathArr))
End Function' GetExcelName

'����ر�Excel���
Function CloseExcel(ByVal ExcelFile)
    ExcelFile.Save
    ExcelFile.Close
    ExcelObj.Quit
End Function' CloseExcel

'��ȡ���м�¼
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset SSProcess.GetProjectFileName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (SSProcess.GetProjectFileName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst SSProcess.GetProjectFileName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (SSProcess.GetProjectFileName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord SSProcess.GetProjectFileName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext SSProcess.GetProjectFileName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset SSProcess.GetProjectFileName, StrSqlStatement
    SSProcess.CloseAccessMdb SSProcess.GetProjectFileName
End Function

'��ӵ�����
Function AddDenote(ByVal Value)
    AddDenote = "'" & Value & "'"
End Function' AddDenote

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'������ʾ
Function Ending()
    MsgBox "������"
End Function' Ending