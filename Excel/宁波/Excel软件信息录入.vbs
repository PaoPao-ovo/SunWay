
'Excel����
Dim xlApp,xlFile,xlsheet

'�õغ���GUID �� �ڵش���
Dim YDHXGUID
ZDCode = "9410001"
'���蹤�̹滮���֤GUID
Dim JSGHXKZGUID

'����ID
Dim DTid(10000000)
Sub OnClick()
    
    'ѡȡ�ڵ�
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    'SSProcess.SetSelectCondition "SSObj_Type", "=", "AREA"
    SSProcess.SetSelectCondition "SSObj_Code", "=", ZDCode
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount
    If geoCount = 0 Then
        GetZDID = 0
        Exit Sub
    ElseIf geoCount = 1 Then
        ZDID = SSProcess.GetSelGeoValue (0, "SSObj_ID")
        YDHXGUID = SSProcess.GetSelGeoValue (0, "[YDHXGUID]")
        If YDHXGUID = "{00000000-0000-0000-0000-000000000000}"  Then
            YDHXGUID = GenNewGUID
            SSProcess.SetObjectAttr ZDID, "[YDHXGUID]", YDHXGUID
        End If
    Else
        MsgBox "ͼ���ж����!"
        Exit Sub
    End If
    aa = MsgBox("�������������ݣ��Ƿ�����Ϣ��",4 + 64)'��6 ��7
    If aa = 7 Then  Exit Sub
    
    '��Excel���
    ExcelFile = SSProcess.SelectFileName(1,"ѡ��excel�ļ�",0,"EXCEL Files(*.xlsx)|*.xlsx|EXCEL Files(*.xls)|*.xls|All Files (*.*)|*.*||")
    If ExcelFile = "" Then Exit Sub
    Set xlApp = CreateObject("Excel.Application")
    Set xlFile = xlApp.Workbooks.Open(ExcelFile)
    
    GZLTJ()
    RJPZ()
    RYXX()
    XMKZCLCG()
    YQSB()
    xlApp.quit
End Sub

'���Ա�����
Table_GZLTJ = "INFO_GZLTJ"
Table_RJPZ = "INFO_RJPZ"
Table_RYXX = "INFO_RYXX"
Table_XMKZCLCG = "INFO_XMKZCLCG"
Table_YQSB = "INFO_YQSB"

'����ͳ������Ϣ���
Function GZLTJ()
    EmptyGZLTJInfo()
    Set xlsheet = xlFile.Worksheets("������ͳ�Ʊ�")
    xlsheet.Activate
    gztjlxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 5
                str = xlApp.Cells(i,k)
                If gztjlxx = ""  Then
                    gztjlxx = str
                ElseIf k = 1 Then
                    gztjlxx = gztjlxx & str
                ElseIf k = 5 Then
                    gztjlxx = gztjlxx & "," & str & ";"
                Else
                    gztjlxx = gztjlxx & "," & str
                End If
            Next
        End If
    Next
    Infile = "YDHXGUID,���,��������,������,��������λ,��ע"
    Sql = "select " & Infile & " from " & Table_GZLTJ & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString gztjlxx, ";", arr, Count
    For y = 0 To Count - 2
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'���������Ϣ���
Function RJPZ()
    EmptyRJPZInfo()
    Set xlsheet = xlFile.Worksheets("������ñ�")
    xlsheet.Activate
    rjpzxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 3
                str = xlApp.Cells(i,k)
                If rjpzxx = ""  Then
                    rjpzxx = str
                ElseIf k = 1 Then
                    rjpzxx = rjpzxx & str
                ElseIf k = 3 Then
                    rjpzxx = rjpzxx & "," & str & ";"
                Else
                    rjpzxx = rjpzxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,���,�������,�����;"
    Sql = "select " & Infile & " from " & Table_RJPZ & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString rjpzxx, ";", arr, Count
    For y = 0 To Count - 2
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'��Ա��Ϣ�����
Function RYXX()
    EmptyRYXXInfo()
    Set xlsheet = xlFile.Worksheets("��Ա��Ϣ��")
    xlsheet.Activate
    ryxxxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 6
                str = xlApp.Cells(i,k)
                If ryxxxx = ""  Then
                    ryxxxx = str
                ElseIf k = 1 Then
                    ryxxxx = ryxxxx & str
                ElseIf k = 6 Then
                    ryxxxx = ryxxxx & "," & str & ";"
                Else
                    ryxxxx = ryxxxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,���,����,ְ�ƻ�ְҵ�ʸ�,�ϸ�֤���Ż�ְҵ�ʸ�֤���,��Ҫ����ְ��,��ע"
    Sql = "select " & Infile & " from " & Table_RYXX & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString ryxxxx, ";", arr, Count
    For y = 0 To Count - 2
        
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'��Ա��Ϣ�����
Function RYXX()
    EmptyRYXXInfo()
    Set xlsheet = xlFile.Worksheets("��Ա��Ϣ��")
    xlsheet.Activate
    ryxxxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 6
                str = xlApp.Cells(i,k)
                If ryxxxx = ""  Then
                    ryxxxx = str
                ElseIf k = 1 Then
                    ryxxxx = ryxxxx & str
                ElseIf k = 6 Then
                    ryxxxx = ryxxxx & "," & str & ";"
                Else
                    ryxxxx = ryxxxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,���,����,ְ�ƻ�ְҵ�ʸ�,�ϸ�֤���Ż�ְҵ�ʸ�֤���,��Ҫ����ְ��,��ע"
    Sql = "select " & Infile & " from " & Table_RYXX & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString ryxxxx, ";", arr, Count
    For y = 0 To Count - 2
        FeatureGUID = GenNewGUID
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'��Ŀ���Ʋ����ɹ����
Function XMKZCLCG()
    EmptyXMKZCLCGInfo()
    Set xlsheet = xlFile.Worksheets("��Ŀ���Ʋ����ɹ�")
    xlsheet.Activate
    xmkzclcgxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 6
                str = xlApp.Cells(i,k)
                If xmkzclcgxx = ""  Then
                    xmkzclcgxx = str
                ElseIf k = 1 Then
                    xmkzclcgxx = xmkzclcgxx & str
                ElseIf k = 6 Then
                    xmkzclcgxx = xmkzclcgxx & "," & str & ";"
                Else
                    xmkzclcgxx = xmkzclcgxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,��ע,�߳�H,ƽ������Y,ƽ������X,�ȼ�,���"
    Sql = "select " & Infile & " from " & Table_XMKZCLCG & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString xmkzclcgxx, ";", arr, Count
    For y = 0 To Count - 2
        FeatureGUID = GenNewGUID
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'�����豸���
Function YQSB()
    EmptyYQSBInfo()
    Set xlsheet = xlFile.Worksheets("�����豸")
    xlsheet.Activate
    yqsbxx = ""
    excelhs = xlApp.ActiveSheet.UsedRange.Rows.Count
    For i = 2 To excelhs
        If xlApp.Cells(i,1) <> "" Then
            For k = 1 To 6
                str = xlApp.Cells(i,k)
                If yqsbxx = ""  Then
                    yqsbxx = str
                ElseIf k = 1 Then
                    yqsbxx = yqsbxx & str
                ElseIf k = 6 Then
                    yqsbxx = yqsbxx & "," & str & ";"
                Else
                    yqsbxx = yqsbxx & "," & str
                End If
            Next
        End If
    Next
    
    Infile = "YDHXGUID,���,��������,Ʒ���ͺ�,�������,�ȼ�����,�����춨��Ч��"
    Sql = "select " & Infile & " from " & Table_YQSB & " where ID>0"
    Dim arr(100000),Count,Info
    SSFunc.ScanString yqsbxx, ";", arr, Count
    For y = 0 To Count - 2
        FeatureGUID = GenNewGUID
        Values = YDHXGUID & "," & arr(y)
        InsertInfo Sql,Infile,Values
    Next
End Function

'�޸ı���Ϣ
Function inAttr(sql,infile,invalues)
    ProjectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProjectName
    SSProcess.OpenAccessRecordset ProjectName, sql
    rscount = SSProcess.GetAccessRecordCount (ProjectName, sql)
    If rscount > 0 Then
        SSProcess.AccessMoveFirst ProjectName, sql
        While (SSProcess.AccessIsEOF (ProjectName, sql ) = False)
            SSProcess.ModifyAccessRecord  ProjectName, sql, infile , invalues'�����mdb����
            SSProcess.AccessMoveNext ProjectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProjectName, sql
    SSProcess.CloseAccessMdb ProjectName
End Function

'��ȡ���µ�FeatureGUID
Function GenNewGUID()
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GenNewGUID = Left(TypeLib.Guid,38)
    Set TypeLib = Nothing
End Function

'********�����¼�¼
Function InsertRecord( tableName, fields, values)
    sqlString = "insert into " & tableName & " (" & fields & ") values (" & values & ")"
    InsertRecord = SSProcess.ExecuteSql (sqlString)
End Function



'������֤����Ϣ
Function EmptyRYXXInfo()
    sql = "SELECT * FROM INFO_RYXX where INFO_RYXX.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function

'�����¼
Function InsertInfo(sql,Infile,Values)
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    recordc = SSProcess.GetAccessRecordCount(mdbName, sql)
    SSProcess.AddAccessRecord mdbName,sql,Infile,Values
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
End Function

'���������ñ�����
Function EmptyRJPZInfo()
    sql = "SELECT * FROM INFO_RJPZ where INFO_RJPZ.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
    SSProcess.CloseAccessMdb mdbName
End Function

'��չ�����ͳ�Ʊ���Ϣ
Function EmptyGZLTJInfo()
    sql = "SELECT * FROM INFO_GZLTJ where INFO_GZLTJ.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql  '�����ݿ�
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    
    SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
    SSProcess.CloseAccessMdb mdbName
End Function

'�����Ŀ���Ʋ����ɹ�
Function EmptyXMKZCLCGInfo()
    sql = "SELECT * FROM INFO_XMKZCLCG where INFO_XMKZCLCG.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
    SSProcess.CloseAccessMdb mdbName
End Function

'��������豸
Function EmptyYQSBInfo()
    sql = "SELECT * FROM INFO_YQSB where INFO_YQSB.ID > " & "0" & ";"
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    
    While  SSProcess.AccessIsEOF (mdbName, sql) = False
        SSProcess.DelAccessRecord mdbName, sql
    WEnd
    SSProcess.CloseAccessRecordset mdbName, sql '�ؿ�
    SSProcess.CloseAccessMdb mdbName
End Function