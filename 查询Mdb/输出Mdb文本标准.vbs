
'========================================================Doc操作对象和文件路径操作对象================================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'============================================================配置SQL=========================================================================

'保存Txt文件夹路径
TxtFolderPath = "C:\Users\ASUS\Desktop\文本文件\"

'===========================================功能入口==========================================================================================

'总入口
Sub OnClick()
    
    GetAllMdbName MdbNameStr
    
    MdbNameArr = Split(MdbNameStr,",", - 1,1)
    For k = 0 To UBound(MdbNameArr)
        
        MdbFileName = MdbNameArr(k)
        
        GetAllTableName MdbFileName,TableName
        
        TextFileName = Right(MdbFileName,Len(MdbFileName) - InStrRev(MdbFileName,"\"))
        TextFileName = TxtFolderPath & Replace(TextFileName,".mdb",".txt")
        
        CreatTxt TextFileName,TxtFile
        
        TableNameArr = Split(TableName,",", - 1,1)
        
        For i = 0 To UBound(TableNameArr)
            
            GetFildsName MdbFileName,TableNameArr(i),YSDMBool
            
            If YSDMBool = True Then
                SqlStr = "Select DISTINCT " & "YSDM" & " FROM " & TableNameArr(i)
                GetSQLRecordAll MdbFileName,SqlStr,DicArr,LxCount
                If LxCount > 0 Then
                    For j = 0 To UBound(DicArr)
                        If DicArr(j) <> "" Then
                            LenNum = Len(Trim(DicArr(j)))
                            If LenNum > 10 Then
                                TxtFile.WriteLine Trim(DicArr(j)) & "," & Right(Trim(DicArr(j)),LenNum - 4)
                            Else
                                TxtFile.WriteLine Trim(DicArr(j)) & ","
                            End If
                        End If
                    Next 'j
                End If
            End If
        Next 'i
        
        TxtFile.Close
        
    Next 'k
    
End Sub' OnClick

'获取所有Mdb文件名
Function GetAllMdbName(ByRef MdbNameStr)
    MdbNameStr = ""
    Set FolderObj = FileSysObj.GetFolder("C:\Users\ASUS\Desktop\输出Mdb")
    Set AllFiles = FolderObj.Files
    For Each SingleFile In AllFiles
        If InStrRev(SingleFile.Name,".mdb") > 0 Then
            If MdbNameStr = "" Then
                MdbNameStr = SingleFile.Path
            Else
                MdbNameStr = MdbNameStr & "," & SingleFile.Path
            End If
        End If
    Next ' SingleFile
End Function' GetAllMdbName

'是否包含YSDM字段
Function GetFildsName(ByVal MdbFileName,ByVal TableName,ByRef YSDMBool)
    YSDMBool = False
    SSProcess.OpenAccessMdb MdbFileName
    SSProcess.GetAccessFieldInfo MdbFileName, TableName, FieldsInfo
    SSProcess.CloseAccessMdb MdbFileName
    FieldsArr = Split(FieldsInfo,",", - 1,1)
    For i = 0 To UBound(FieldsArr)
        If FieldsArr(i) = "YSDM" Then
            YSDMBool = True
        End If
    Next 'i
End Function' GetFildsName

'输出文本文件
Function CreatTxt(ByVal TextFileName,ByRef TxtFile)
    Set TxtFile = FileSysObj.CreateTextFile(TextFileName,True)
End Function' CreatTxt

'获取所有表名称
Function GetAllTableName(ByVal MdbFileName,ByRef TableName)
    SSProcess.OpenAccessMdb MdbFileName
    SSProcess.GetAccessTableNames MdbFileName,TableName
    SSProcess.CloseAccessMdb MdbFileName
End Function' GetAllTableName

'获取所有记录
Function GetSQLRecordAll(ByVal MdbFileName,ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    SSProcess.OpenAccessMdb MdbFileName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset MdbFileName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (MdbFileName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst MdbFileName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (MdbFileName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord MdbFileName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext MdbFileName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset MdbFileName, StrSqlStatement
    SSProcess.CloseAccessMdb MdbFileName
End Function