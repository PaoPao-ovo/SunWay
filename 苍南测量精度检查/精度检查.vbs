
' 1����ģ�帴�Ƶ���ǰ��Ŀ·����
' 2��ͨ��SQL��ѯ���¹��ߵ����Ա�ġ���̽��š���DEAFULTͼ�㡾0����ĵ�����ͬ�ĵ�

' [���߱�����Ϣ]
' ��� = ""
' ��Ŀ���� = ""
' ��Ŀ��ַ = ""
' ��Ƶ�λ = ""
' ���赥λ = ""
' ί�е�λ = ""
' ��ҵʱ�� = ""
' ���ʱ�� = ""
' �����ϲ�ֵ = ""
' �߳����ϲ�ֵ = ""

'========================================================Excel����������ļ�·����������======================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Excel��������
Dim ExcelObj
Set ExcelObj = CreateObject("Excel.Application")

'============================================================����¼����====================================================================

'��鼯��Ŀ����
Dim strGroupName
strGroupName = "���߾��ȼ��"

'��鼯������
Dim strCheckName
strCheckName = "���꾫�ȼ��"

'���ģ������
Dim CheckmodelName
CheckmodelName = "�Զ���ű������->���꾫�ȼ��"

'�������
Dim strDescription
strDescription = "���꾫�ȳ���"

'=============================================================�������=======================================================================

Sub OnClick()
    
    AllVisible
    
    FileSysObj.CopyFile  SSProcess.GetSysPathName (7) & "���ģ��\" & "�������ȵ����ģ��.xlsx",SSProcess.GetSysPathName(5) & "�������ȵ����.xlsx"
    
    OpenExcel SSProcess.GetSysPathName(5) & "�������ȵ����.xlsx",ExcleFile
    
    GetPoiNameZero ZeorCount,ZeroArr,ZeroStr
    
    GetGXDDH GXDDHArr,DhLxCount,ZeroStr
    
    InsertExcle 3,EndRow,ZeorCount,ZeroArr,GXDDHArr,DhLxCount
    
    InsertDiff 3,EndRow
    
    GetResultInfo ZeorCount,3,EndRow,OverPoiCount,OverHeightCount,AveragePoi,AverageHeight,MiddlePoi,MiddleHei,OverPercent,OverPoiName,OverHeiName,MaxLen,MaxHei
    
    GXEWB MaxLen,MaxHei

    InsertResult EndRow + 1,ZeorCount,OverPoiCount,OverHeightCount,AveragePoi,AverageHeight,MiddlePoi,MiddleHei,OverPercent
    
    DelSelCol 6
    
    CloseExcel ExcleFile
    
    ClearCheckRecord
    
    AddRecord OverPoiName,OverHeiName
    
End Sub' OnClick

'==========================================================Excel��ֵ===================================================================

'��дͳ�ƽ��
Function InsertResult(ByVal ResultRow,ByVal ZeroCount,ByVal OverPoiCount,ByVal OverHeightCount,ByVal AveragePoi,ByVal AverageHeight,ByVal MiddlePoi,ByVal MiddleHei,ByVal OverPercent)
    If ZeroCount < 15 Then
        If OverPoiCount > 0 And OverHeightCount > 0 Then
            ResultString = "ͳ�ƽ������������" & ZeroCount & "������λ����ĵ���" & OverPoiCount & "�����̳߳���ĸ���" & OverHeightCount & "���������λռ�ܵ����ٷֱ�" & OverPercent & "%����λ���ƽ��ֵ��" & AveragePoi & "��M�����߳����ƽ��ֵ��" & AverageHeight & "��M��"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount > 0 And OverHeightCount = 0 Then
            ResultString = "ͳ�ƽ������������" & ZeroCount & "������λ����ĵ���" & OverPoiCount & "���������λռ�ܵ����ٷֱ�" & OverPercent & "%����λ���ƽ��ֵ��" & AveragePoi & "��M�����߳����ƽ��ֵ��" & AverageHeight & "��M��"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount = 0 And OverHeightCount > 0 Then
            ResultString = "ͳ�ƽ������������" & ZeroCount & "��" & "�̳߳���ĸ���" & OverHeightCount & "���������λռ�ܵ����ٷֱ�" & OverPercent & "%����λ���ƽ��ֵ��" & AveragePoi & "��M�����߳����ƽ��ֵ��" & AverageHeight & "��M��"
            ExcelObj.Cells(ResultRow,1) = ResultString
        Else
            ResultString = "ͳ�ƽ������������" & ZeroCount & "��" & "��λ���ƽ��ֵ��" & AveragePoi & "��M�����߳����ƽ��ֵ��" & AverageHeight & "��M��"
            ExcelObj.Cells(ResultRow,1) = ResultString
        End If
    ElseIf ZeroCount >= 15 Then
        If OverPoiCount > 0 And OverHeightCount > 0 Then
            ResultString = "ͳ�ƽ������������" & ZeroCount & "������λ����ĵ���" & OverPoiCount & "�����̳߳���ĸ���" & OverHeightCount & "���������λռ�ܵ����ٷֱ�" & OverPercent & "%����λ�������" & MiddlePoi & "��M�����߳��������" & MiddleHei & "��M��"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount > 0 And OverHeightCount = 0 Then
            ResultString = "ͳ�ƽ������������" & ZeroCount & "������λ����ĵ���" & OverPoiCount & "���������λռ�ܵ����ٷֱ�" & OverPercent & "%����λ�������" & MiddlePoi & "��M�����߳��������" & MiddleHei & "��M��"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount = 0 And OverHeightCount > 0 Then
            ResultString = "ͳ�ƽ������������" & ZeroCount & "�����̳߳���ĸ���" & OverHeightCount & "���������λռ�ܵ����ٷֱ�" & OverPercent & "%����λ�������" & MiddlePoi & "��M�����߳��������" & MiddleHei & "��M��"
            ExcelObj.Cells(ResultRow,1) = ResultString
        Else
            ResultString = "ͳ�ƽ������������" & ZeroCount & "��" & "��λ�������" & MiddlePoi & "��M�����߳��������" & MiddleHei & "��M��"
            ExcelObj.Cells(ResultRow,1) = ResultString
        End If
    End If
End Function' InsertResult

'��ȡ�����Ϣ
Function GetResultInfo(ByVal ZeroCount,ByVal StartRow,ByVal EndRow,ByRef OverPoiCount,ByRef OverHeightCount,ByRef AveragePoi,ByRef AverageHeight,ByRef MiddlePoi,ByRef MiddleHei,ByRef OverPercent,ByRef OverPoiName,ByRef OverHeiName,ByRef MaxLen,ByRef MaxHei)
    OverPoiCount = 0
    OverHeightCount = 0
    OverNum = 0
    AveragePoi = 0.00
    AverageHeight = 0.00
    SquarePoi = 0
    SquareHei = 0
    MiddlePoi = 0.00
    MiddleHei = 0.00
    OverPoiName = ""
    OverHeiName = ""
    MaxLen = 0.00
    MaxHei = 0.00
    For i = StartRow To EndRow
        If Transform(ExcelObj.Cells(i,10)) > 0.05 Then
            OverPoiCount = OverPoiCount + 1
            If OverPoiName = "" Then
                OverPoiName = "'" & ExcelObj.Cells(i,2) & "'"
            Else
                OverPoiName = OverPoiName & "," & "'" & ExcelObj.Cells(i,2) & "'"
            End If
        End If
        If Transform(ExcelObj.Cells(i,10)) > MaxLen Then
            MaxLen = Transform(ExcelObj.Cells(i,10))
        End If
        If Transform(ExcelObj.Cells(i,11)) > MaxHei Then
            MaxHei = Transform(ExcelObj.Cells(i,10))
        End If
        If Transform(ExcelObj.Cells(i,11)) > 0.03 Then
            OverHeightCount = OverHeightCount + 1
            If OverHeiName = "" Then
                OverHeiName = "'" & ExcelObj.Cells(i,2) & "'"
            Else
                OverHeiName = OverHeiName & "," & "'" & ExcelObj.Cells(i,2) & "'"
            End If
        End If
        TotalPoi = TotalPoi + Transform(ExcelObj.Cells(i,10))
        TotalHei = TotalHei + Transform(ExcelObj.Cells(i,11))
        SquarePoi = SquarePoi + Transform(ExcelObj.Cells(i,10)) ^ 2
        SquareHei = SquareHei + Transform(ExcelObj.Cells(i,11)) ^ 2
        If Transform(ExcelObj.Cells(i,10)) > 0.05 Or Transform(ExcelObj.Cells(i,11)) > 0.03 Then
            OverNum = OverNum + 1
        End If
    Next 'i
    OverPercent = Round((OverNum / ZeroCount),2) * 100
    AveragePoi = Round(TotalPoi / ZeroCount,3)
    AverageHeight = Round(TotalHei / ZeroCount,3)
    MiddlePoi = Round(Sqr(SquarePoi / ZeroCount),3)
    MiddleHei = Round(Sqr(SquareHei / ZeroCount),3)
End Function' GetResultInfo

'��ȡ���Ȳ�
Function GetDiffS(x1,y1,x2,y2)
    GetDiffS = Round(Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2),3)
End Function' GetDiffS

'��д�̲߳�;����
Function InsertDiff(ByVal StartRow,ByVal EndRow)
    For i = StartRow To EndRow
        ExcelObj.Cells(i,10) = GetDiffS(Transform(ExcelObj.Cells(i,3)),Transform(ExcelObj.Cells(i,4)),Transform(ExcelObj.Cells(i,7)),Transform(ExcelObj.Cells(i,8)))
        ExcelObj.Cells(i,11) = Round(Abs(Transform(ExcelObj.Cells(i,5)) - Transform(ExcelObj.Cells(i,9))),3)
    Next 'i
End Function' InsertDiff

'��дExcel
Function InsertExcle(ByVal StartRow,ByRef EndRow,ByVal ZeroCount,ByVal ZeroArr(),ByVal GXDDHArr(),ByVal DhLxCount)
    EndRow = ZeroCount + 2
    If EndRow <= 5  Then
        For i = StartRow To EndRow
            ExcelObj.Cells(i,1) = i - 2
            ExcelObj.Cells(i,6) = ZeroArr(i - StartRow,0)
            InsertZeroXYZ i,i - StartRow,ZeroArr
            InsertGxPoint StartRow,EndRow,GXDDHArr,DhLxCount
        Next 'i
    ElseIf EndRow > 5 Then
        InsertRows StartRow,ZeroCount - 3
        For i = StartRow To EndRow
            ExcelObj.Cells(i,1) = i - 2
            ExcelObj.Cells(i,6) = ZeroArr(i - StartRow,0)
            InsertZeroXYZ i,i - StartRow,ZeroArr
            InsertGxPoint StartRow,EndRow,GXDDHArr,DhLxCount
        Next 'i
    End If
End Function' InsertExcle

'��д���ߵ�ֵ
Function InsertGxPoint(ByVal StartRow,ByVal EndRow,ByVal GXDDHArr,ByVal DhLxCount)
    For i = 0 To DhLxCount
        For j = StartRow To EndRow
            If ExcelObj.Cells(j,6) = SSProcess.GetObjectAttr(GXDDHArr(i),"[WTDH]") Then
                ExcelObj.Cells(j,2) = SSProcess.GetObjectAttr(GXDDHArr(i),"[WTDH]")
                ExcelObj.Cells(j,3) = Round(Transform(SSProcess.GetObjectAttr(GXDDHArr(i),"SSObj_Y")),3)
                ExcelObj.Cells(j,4) = Round(Transform(SSProcess.GetObjectAttr(GXDDHArr(i),"SSObj_X")),3)
                ExcelObj.Cells(j,5) = Round(Transform(SSProcess.GetObjectAttr(GXDDHArr(i),"SSObj_Z")),3)
            End If
        Next 'j
    Next 'i
End Function' InsertGxPoint

'��д0���XYZֵ��Excel�����
Function InsertZeroXYZ(ByVal InsertRow,ByVal Index,ByVal ZeroArr())
    ExcelObj.Cells(InsertRow,7) = Round(Transform(SSProcess.GetObjectAttr(ZeroArr(Index,1),"SSObj_Y")),3)
    ExcelObj.Cells(InsertRow,8) = Round(Transform(SSProcess.GetObjectAttr(ZeroArr(Index,1),"SSObj_X")),3)
    ExcelObj.Cells(InsertRow,9) = Round(Transform(SSProcess.GetObjectAttr(ZeroArr(Index,1),"SSObj_Z")),3)
End Function' InsertZeroXYZ

'��ȡ���е�0��ĵ�����ID
Function GetPoiNameZero(ByRef ZeorCount,ByRef ZeroArr(),ByRef ZeroStr)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "0"
    SSProcess.SetSelectCondition "SSObj_PointName", "<>", ""
    SSProcess.SelectFilter
    ZeorCount = SSProcess.GetSelGeoCount
    ReDim  ZeroArr(ZeorCount - 1,1)
    For i = 0 To ZeorCount - 1
        ZeroArr(i,0) = SSProcess.GetSelGeoValue(i,"SSObj_PointName")
        ZeroArr(i,1) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        If ZeroStr = "" Then
            ZeroStr = "'" & SSProcess.GetSelGeoValue(i,"SSObj_PointName") & "'"
        Else
            ZeroStr = ZeroStr & "," & "'" & SSProcess.GetSelGeoValue(i,"SSObj_PointName") & "'"
        End If
    Next 'i
End Function' GetPoiName_0

Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

'��ȡ����������������̽��ŵ�ID
Function GetGXDDH(ByRef GXDDHArr(),ByRef DhLxCount,ByVal ZeoStr)
    SqlString = "Select ���¹��ߵ����Ա�.ID From ���¹��ߵ����Ա� Inner Join GeoPointTB on ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And ���¹��ߵ����Ա�.WTDH In " & "(" & ZeoStr & ")"
    GetSQLRecordAll SqlString,GXDDHArr,DhLxCount
End Function' GetGXDDH

'��Ӽ���¼
Function AddRecord(ByVal OverPoiName,ByVal OverHeiName)
    If OverPoiName <> "" Then
        SqlString = "Select ���¹��ߵ����Ա�.ID From ���¹��ߵ����Ա� Inner Join GeoPointTB on ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And ���¹��ߵ����Ա�.WTDH In " & "(" & OverPoiName & ")"
        GetSQLRecordAll SqlString,OverPoiArr,OverPoiCount
        For i = 0 To OverPoiCount - 1
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(OverPoiArr(i),"SSObj_X"),SSProcess.GetObjectAttr(OverPoiArr(i),"SSObj_Y"),0,0,OverPoiArr(i),""
        Next 'i
    End If
    strCheckName = "�߳̾��ȼ��"
    CheckmodelName = "�Զ���ű������->�߳̾��ȼ��"
    strDescription = "�߳̾��ȳ���"
    If OverHeiName <> "" Then
        SqlString = "Select ���¹��ߵ����Ա�.ID From ���¹��ߵ����Ա� Inner Join GeoPointTB on ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And ���¹��ߵ����Ա�.WTDH In " & "(" & OverHeiName & ")"
        GetSQLRecordAll SqlString,OverHeiArr,OverHeiCount
        For i = 0 To OverHeiCount - 1
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(OverHeiArr(i),"SSObj_X"),SSProcess.GetObjectAttr(OverHeiArr(i),"SSObj_Y"),0,0,OverHeiArr(i),""
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' AddRecord

'��ռ�鼯
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'��ȡ���м�¼
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "��ѯ���Ϊ�գ�����ֹͣ��",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset ProJectName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (ProJectName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst ProJectName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (ProJectName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord ProJectName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext ProJectName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProJectName, StrSqlStatement
    SSProcess.CloseAccessMdb ProJectName
End Function

'����ָ������
Function InsertRows(ByVal StartRow,ByVal InsertCount)
    For i = 0 To InsertCount - 1
        ExcelObj.ActiveSheet.Rows(StartRow).Insert
    Next 'i
End Function' InsertRows

'��Excel��
Function OpenExcel(ByVal FilePath,ByRef ExcleFile)
    ExcelObj.Application.Visible = True
    Set ExcleFile = ExcelObj.WorkBooks.Open(FilePath)
    Set ExcelSheet = ExcleFile.WorkSheets("���ȵ����")
    ExcelSheet.Activate
End Function

'������ͼ��
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'����ر�Excel���
Function CloseExcel(ByVal ExcleFile)
    ExcleFile.Save
    ExcelObj.Quit
End Function' CloseExcel

'ɾ��ָ������
Function DelSelCol(ByVal ColNum)
    ExcelObj.ActiveSheet.Columns(ColNum).Delete
End Function' DelSelCol

'ˢ�¶�ά��
Function GXEWB(DWZDJC,GCZDJC)
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    sql = "update  ������Ŀ��Ϣ�� set DWZDJC = " & DWZDJC & "where ������Ŀ��Ϣ��.ID= 1"
    SSProcess.ExecuteAccessSql  mdbName,sql
    sql = "update  ������Ŀ��Ϣ�� set GCZDJC = " & GCZDJC & "where ������Ŀ��Ϣ��.ID= 1"
    SSProcess.ExecuteAccessSql  mdbName,sql
    SSProcess.CloseAccessMdb mdbName
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function