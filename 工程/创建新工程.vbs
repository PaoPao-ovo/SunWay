
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")

Sub OnClick()
    CurrentProPath = Mid(SSProcess.GetProjectFileName,1,Len(SSProcess.GetProjectFileName) - 4) & "����" & ".edb"
    Set FormerFileObj = FileSystemObject.GetFile(SSProcess.GetProjectFileName)
    FormerFileObj.Copy CurrentProPath
    SSProcess.OpenDatabase   CurrentProPath
    
    SqlStr = "Select ���¹��������Ա�.ID,���¹��������Ա�.FSFS From ���¹��������Ա� Inner Join GeoLineTB On ���¹��������Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1) 'ID,[FSFS]
        If IsNumeric(SingleLineArr(1)) Then
            '���Զ��գ�0,1,2,3,4,5,6,7,8,9,10,11,12;
            'ֱ��,����,�ܿ�,�ܹ�,�ܿ�,����,�ϼ�,Сͨ��,�ۺϹ��ȣ�����,�˷�,��������,����,ˮ��
            Select Case SingleLineArr(1)
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","ֱ��"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","����"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ܿ�"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ܹ�"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ܿ�"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","����"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ϼ�"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","Сͨ��"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�ۺϹ��ȣ�����"
                Case "9"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","�˷�"
                Case "10"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","��������"
                Case "11"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","����"
                Case "12"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","ˮ��"
            End Select
        End If
    Next 'i
    
    SqlStr = "Select ���¹��������Ա�.ID,���¹��������Ա�.SJYL From ���¹��������Ա� Inner Join GeoLineTB On ���¹��������Ա�.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1) 'ID,[SJYL]
        If IsNumeric(SingleLineArr(1)) Then
            '���Զ��գ�0,1,2,3,4,5,6,7,8;
            '��ѹ,��ѹA��,��ѹB��,�θ�ѹA��,�θ�ѹB��,��ѹ,��ѹA��,��ѹB��,��ѹ,�˷�,��������,����,ˮ��
            Select Case SingleLineArr(1)
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹ"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹA��"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹB��"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","�θ�ѹA��"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","�θ�ѹB��"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹ"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹA��"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹB��"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","��ѹ"
            End Select
        End If
    Next 'i
    
    SSProcess.CloseDatabase   CurrentProPath

End Sub

'========================================���ߺ���==================================================

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