
'========================================================Excel����������ļ�·����������======================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Excel��������
Dim ExcelObj
Set ExcelObj = CreateObject("Excel.Application")

'��׼���ģ��Excel�ļ�·��
Dim TempLateFilePath
TempLateFilePath = SSProcess.GetSysPathName(7) & "���ģ��\" & "��׼���.xls"

'===========================================�������========================================================

'�����
Sub OnClick()
    
    '1�������յ�Excel�ļ�
    FilePath = SSProcess.GetSysPathName(5) & "��׼���.xls"
    FileSysObj.CopyFile TempLateFilePath,FilePath
    
    '2����д����Sheet,��EPS�ܵ�Sheet
    OpenExcel FilePath,ExcelFile,ExcelSheet,1
    
    '3����ȡWTDH��X,Y,Zֵ
    SqlStr = "Select ���¹��ߵ����Ա�.ID From ���¹��ߵ����Ա� INNER JOIN GeoPointTB ON ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE ([GeoPointTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll SqlStr,PointArr,PointCount
    SearchRow = PointCount + 1
    StartPointRow = PointCount + 1
    
    '4����дEPS���ߵ�Sheet
    For i = 0 To PointCount - 1
        ExcelObj.Cells(i + 1,1) = SSProcess.GetObjectAttr(PointArr(i),"[WTDH]")
        ExcelObj.Cells(i + 1,2) = SSProcess.GetObjectAttr(PointArr(i),"SSObj_Y")
        ExcelObj.Cells(i + 1,3) = SSProcess.GetObjectAttr(PointArr(i),"SSObj_X")
        ExcelObj.Cells(i + 1,4) = SSProcess.GetObjectAttr(PointArr(i),"SSObj_Z")
    Next 'i
    
    '5����EPS����Sheet
    ChangSheet ExcelFile,ExcelSheet,2
    
    '6����ȡ����Ϣ����ֵ
    Filds = "ID,GXQDDH,GXZDDH,GJ,GC,WYKS,ZKS,LX,DYZ,D_Dia,SHGL,GXQDMS,GXZDMS,JCNY,FSFS,QSDW,BZ,SJYL" '��Ӧ������0,1,2,5,6,7,8,9,10,11,12,13,14,15,16,17,19
    ColNum = "0,1,2,5,6,7,8,9,10,11,12,13,14,15,16,17,19,21"
    ColArr = Split(ColNum,",", - 1,1)
    
    SqlStr = "Select ���¹��������Ա�." & Filds & " From ���¹��������Ա� INNER JOIN GeoLineTB ON ���¹��������Ա�.ID = GeoLineTB.ID WHERE ([GeoLineTB].[Mark] Mod 2)<>0"
    
    GetSQLRecordAll SqlStr,LineArr,LineCount
    Dim LineIdArr()
    ArrSize = 0
    ReDim LineIdArr(ArrSize)
    Row = 0
    
    For i = 0 To LineCount - 1
        InfoArr = Split(LineArr(i),",", - 1,1)
        If InfoArr(16) <> "���߿�" Then
            LineIdArr(i) = InfoArr(0)
            ArrSize = ArrSize + 1
            ReDim Preserve LineIdArr(ArrSize)
            For j = 1 To UBound(ColArr)
                If InfoArr(j) <> "*" Then
                    ExcelObj.Cells(Row + 2,Transform(ColArr(j))) = InfoArr(j)
                End If
            Next 'j
            Row = Row + 1
        Else
            Row = Row
        End If
    Next 'i
    
    '7����ȡ�����������ƫ�ĵ㾮�ţ��㱸ע
    Filds = "TZ,FSW,PXJW,BZ"
    ColNum = "3,4,18,20"
    ColArr_Point = Split(ColNum,",", - 1,1)
    For i = 0 To UBound(LineIdArr) - 1
        WTDH = SSProcess.GetObjectAttr(LineIdArr(i),"[GXQDDH]")
        SqlStr = "Select ���¹��ߵ����Ա�." & Filds & " From ���¹��ߵ����Ա� INNER JOIN GeoPointTB ON ���¹��ߵ����Ա�.ID = GeoPointTB.ID WHERE ([GeoPointTB].[Mark] Mod 2)<>0 And ���¹��ߵ����Ա�.WTDH = " & "'" & WTDH & "'"
        GetSQLRecordAll SqlStr,Line_PointArr,PointCount
        Info_PointArr = Split(Line_PointArr(0),",", - 1,1)
        For j = 0 To 3
            ExcelObj.Cells(i + 2,Transform(ColArr_Point(j))) = Info_PointArr(j)
        Next 'j
    Next 'i
    
    '8���߱�עΪ�������ߵ��ߵ���ֵ��ʼ��
    StartRow = Row + 2
    
    '9����ȡ����Ϣ����ֵ
    Filds = "ID,GXQDDH,GXZDDH,GJ,GC,WYKS,ZKS,LX,DYZ,D_Dia,SHGL,GXQDMS,GXZDMS,JCNY,FSFS,QSDW,BZ,SJYL" '��Ӧ������0,1,2,5,6,7,8,9,10,11,12,13,14,15,16,17,19
    ColNum = "0,1,2,5,6,7,8,9,10,11,12,13,14,15,16,17,19,21"
    JNColArr = Split(ColNum,",", - 1,1)
    
    SqlStr = "Select ���¹��������Ա�." & Filds & " From ���¹��������Ա� INNER JOIN GeoLineTB ON ���¹��������Ա�.ID = GeoLineTB.ID WHERE ([GeoLineTB].[Mark] Mod 2)<>0 And ���¹��������Ա�.BZ = " & "'" & "���߿�" & "'"
    GetSQLRecordAll SqlStr,JNLineArr,LineCount
    
    Dim JNIdArr()
    ReDim JNIdArr(LineCount - 1)
    
    '10����д�߱�עΪ�������ߵ��ߵ���Ϣ
    For i = 0 To LineCount - 1
        JNInfoArr = Split(JNLineArr(i),",", - 1,1)
        JNIdArr(i) = JNInfoArr(0)
        For j = 1 To UBound(JNColArr)
            If JNInfoArr(j) <> "*" Then
                ExcelObj.Cells(StartRow + i,Transform(JNColArr(j))) = JNInfoArr(j)
                ExcelObj.Cells(StartRow + i,3) = "���ߵ�"
                'ExcelObj.Cells(StartRow + i,4) = ""
                'ExcelObj.Cells(StartRow + i,18) = ""
                ExcelObj.Cells(StartRow + i,20) = "���ߵ�"
            End If
        Next 'j
    Next 'i
    
    '11���л�ΪEPS���ߵ�Sheet
    ChangSheet ExcelFile,ExcelSheet,1
    
    '12����д�߽ڵ��X,Y,Zֵ
    For i = 0 To UBound(JNIdArr)
        PointCount = SSProcess.GetObjectAttr(JNIdArr(i),"SSObj_PointCount")
        GXQDDH = SSProcess.GetObjectAttr(JNIdArr(i),"[GXQDDH]")
        GXZDDH = SSProcess.GetObjectAttr(JNIdArr(i),"[GXZDDH]")
        For j = 0 To PointCount - 1
            SSProcess.GetObjectPoint JNIdArr(i),j,X,Y,Z,PointType,Name
            X = Round(Transform(X),3)
            Y = Round(Transform(Y),3)
            Z = Round(Transform(Z),3)
            If j = 0 Then
                ExcelObj.Cells(StartPointRow,1) = GXQDDH
                ExcelObj.Cells(StartPointRow,2) = Y
                ExcelObj.Cells(StartPointRow,3) = X
                ExcelObj.Cells(StartPointRow,4) = Z
                StartPointRow = StartPointRow + 1
            Else
                ExcelObj.Cells(S tartPointRow,1) = GXZDDH
                ExcelObj.Cells(StartPointRow,2) = Y
                ExcelObj.Cells(StartPointRow,3) = X
                ExcelObj.Cells(StartPointRow,4) = Z
                StartPointRow = StartPointRow + 1
            End If
        Next 'j
    Next 'i
    
    '13��ɾ���ظ��ĵ�
    For i = SearchRow To StartPointRow - 2
        For j = i + 1 To StartPointRow - 1
            If ExcelObj.Cells(i,1) = ExcelObj.Cells(j,1) Then
                ExcelObj.ActiveSheet.Rows(j).Delete
            End If
        Next 'j
    Next 'i
    
    '14�����沢�ر�Excel
    CloseExcel ExcelFile
    
End Sub' OnClick

'=================================================�����ຯ��=====================================================

'��Excel��
Function OpenExcel(ByVal FilePath,ByRef ExcleFile,ByRef ExcelSheet,ByVal Num)
    ExcelObj.Application.Visible = False
    Set ExcleFile = ExcelObj.WorkBooks.Open(FilePath)
    Set ExcelSheet = ExcleFile.WorkSheets(Num)
    ExcelSheet.Activate
End Function

'�л�Sheet
Function ChangSheet(ByVal ExcleFile,ByVal ExcelSheet,ByVal Num)
    Set ExcelSheet = ExcleFile.WorkSheets(Num)
    ExcelSheet.Activate
End Function' ChangSheet

'����ر�Excel���
Function CloseExcel(ByVal ExcelFile)
    ExcelFile.Save
    ExcelFile.Close
    ExcelObj.Quit
End Function' CloseExcel

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

'��������ת��
Function Transform(ByVal Values)
    If Values <> "" Then
        If IsNumeric(Values) = True Then
            Values = CDbl(Values)
        End If
    Else
        Values = 0
    End If
    Transform = Values
End Function'Transform