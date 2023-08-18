
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")

Sub OnClick()
    CurrentProPath = Mid(SSProcess.GetProjectFileName,1,Len(SSProcess.GetProjectFileName) - 4) & "副本" & ".edb"
    Set FormerFileObj = FileSystemObject.GetFile(SSProcess.GetProjectFileName)
    FormerFileObj.Copy CurrentProPath
    SSProcess.OpenDatabase   CurrentProPath
    
    SqlStr = "Select 地下管线线属性表.ID,地下管线线属性表.FSFS From 地下管线线属性表 Inner Join GeoLineTB On 地下管线线属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1) 'ID,[FSFS]
        If IsNumeric(SingleLineArr(1)) Then
            '属性对照：0,1,2,3,4,5,6,7,8,9,10,11,12;
            '直埋,管埋,管块,管沟,架空,地面,上架,小通道,综合管廊（沟）,人防,井内连线,顶管,水下
            Select Case SingleLineArr(1)
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","直埋"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","管埋"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","管块"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","管沟"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","架空"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","地面"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","上架"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","小通道"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","综合管廊（沟）"
                Case "9"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","人防"
                Case "10"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","井内连线"
                Case "11"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","顶管"
                Case "12"
                SSProcess.SetObjectAttr SingleLineArr(0),"[FSFS]","水下"
            End Select
        End If
    Next 'i
    
    SqlStr = "Select 地下管线线属性表.ID,地下管线线属性表.SJYL From 地下管线线属性表 Inner Join GeoLineTB On 地下管线线属性表.ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0"
    GetSQLRecordAll SqlStr,LineArr,LineCount
    For i = 0 To LineCount - 1
        SingleLineArr = Split(LineArr(i),",", - 1,1) 'ID,[SJYL]
        If IsNumeric(SingleLineArr(1)) Then
            '属性对照：0,1,2,3,4,5,6,7,8;
            '高压,高压A级,高压B级,次高压A级,次高压B级,中压,中压A级,中压B级,低压,人防,井内连线,顶管,水下
            Select Case SingleLineArr(1)
                Case "0"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","高压"
                Case "1"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","高压A级"
                Case "2"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","高压B级"
                Case "3"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","次高压A级"
                Case "4"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","次高压B级"
                Case "5"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","中压"
                Case "6"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","中压A级"
                Case "7"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","中压B级"
                Case "8"
                SSProcess.SetObjectAttr SingleLineArr(0),"[SJYL]","低压"
            End Select
        End If
    Next 'i
    
    SSProcess.CloseDatabase   CurrentProPath

End Sub

'========================================工具函数==================================================

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
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