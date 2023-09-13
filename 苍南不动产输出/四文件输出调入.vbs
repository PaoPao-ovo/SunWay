' /*
'  * @Description: 请填写简介
'  * @Author: LHY
'  * @Date: 2023-09-13 09:09:54
'  * @LastEditors: LHY
'  * @LastEditTime: 2023-09-13 11:28:31
'  */

Dim filenames(100000),SWJfilename(1000000)
Dim filecount,SWJfilenameCount
filecount = 0
zdnr = ""
Sub OnClick()
    
    '当前的项目路径
    Dim ProJectName

    ProJectName = SSProcess.GetProjectFileName
    
    Export EdbName,ExportPath
    
    FormerPath = Replace(ExportPath,EdbName,"")
    
    ZhengHePath = FormerPath & "整合文件-" & EdbName
    
    ZDANDZRZCODE = "9130223,9210123"
    '输出四文件
    Dim Fzdnr(1000000)
    Dim FzdnrCount
    mark = 0
    
    SSProcess.ClearDataXParameter
    SSProcess.SetDataXParameter "DataType","21"
    CurProjectName = FormerPath
    'CurProjectName = SSProcess.GetImportFileName() '获取四文件路径
    XGwz1 = InStrRev(CurProjectName,"\")  '获取“\”最后一次出现的位置
    CurProjectName = Left(CurProjectName,XGwz1)
    If CurProjectName <> "" Then
        TemplateFileName = SSProcess.GetTemplateFileName  '获取当前模板名称
        GetAllFiles CurProjectName,"edb"  '获取目录下四文件
        
        For i = 0 To  filecount - 1
            If Replace (filenames(i),"_Add.edb","") <> filenames(i) Then
                SWJfilename (SWJfilenameCount) = filenames(i)
                SWJfilenameCount = SWJfilenameCount + 1
            End If
            If Replace (filenames(i),"_Delete.edb","") <> filenames(i) Then
                SWJfilename (SWJfilenameCount) = filenames(i)
                SWJfilenameCount = SWJfilenameCount + 1
            End If
            If Replace (filenames(i),"_Edit.edb","") <> filenames(i) Then
                SWJfilename (SWJfilenameCount) = filenames(i)
                SWJfilenameCount = SWJfilenameCount + 1
            End If
            If Replace (filenames(i),"_None.edb","") <> filenames(i) Then
                SWJfilename (SWJfilenameCount) = filenames(i)
                SWJfilenameCount = SWJfilenameCount + 1
            End If
        Next
        
        
        ZLBfileName = ZhengHePath
        IsFileExist ZLBfileName,IsFolderExistMark
        If IsFolderExistMark = True Then DeleteFile ZLBfileName
        SSProcess.CreateDatabase  TemplateFileName, ZLBfileName  '创建新的工程文件
        For j = 0 To SWJfilenameCount - 1
            SSProcess.SetDataXParameter "ImportPathName", SWJfilename(j)
            MaxidDRQ = SSProcess.GetGeoMaxID '调入前最大id
            SSProcess.ImportData
            MaxidDRH = SSProcess.GetGeoMaxID '调入后最大id
            If Replace (SWJfilename(j),"_Add.edb","") <> SWJfilename(j) Then '判断是否为增加的地物
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.ClearSelectConditionGroups
                SSProcess.SetSelectCondition "SSObj_ID", ">", CLng(MaxidDRQ)
                SSProcess.SetSelectCondition "SSObj_ID", "<", CLng(MaxidDRH) + 1
                SSProcess.SelectFilter
                geocount0 = SSProcess.GetSelGeoCount()
                For i1 = 0 To geoCount0 - 1
                    polygonID = SSProcess.GetSelGeoValue( i1, "SSObj_ID" )
                    state = SSProcess.GetSelGeoValue( i1, "[State]" )
                    Code = SSProcess.GetSelGeoValue( i1, "SSObj_Code" )
                    If state = "" Or state = "0"Or  state = "Null" Then
                        If InStr(ZDANDZRZCODE,Code) = 0  Then
                            SSProcess.SetObjectAttr polygonID, "[State]","1"
                        End If
                    End If
                Next
            End If
            If Replace (SWJfilename(j),"_Delete.edb","") <> SWJfilename(j) Then '判断是否为删除的地物
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.ClearSelectConditionGroups
                SSProcess.SetSelectCondition "SSObj_ID", ">", CLng(MaxidDRQ)
                SSProcess.SetSelectCondition "SSObj_ID", "<", CLng(MaxidDRH) + 1
                SSProcess.SelectFilter
                geocount1 = SSProcess.GetSelGeoCount()
                For i3 = 0 To geocount1 - 1
                    polygonID = SSProcess.GetSelGeoValue( i3, "SSObj_ID" )
                    State = SSProcess.GetSelGeoValue( i3, "[State]" )
                    Code = SSProcess.GetSelGeoValue( i3, "SSObj_Code" )
                    'if state="" or state="0" or state="Null" or state="1"  then              
                    SSProcess.SetObjectAttr polygonID, "[State]","3"
                    ' end if 
                Next
            End If
            If Replace (SWJfilename(j),"_Edit.edb","") <> SWJfilename(j) Then '判断是否为改变的地物
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.ClearSelectConditionGroups
                SSProcess.SetSelectCondition "SSObj_ID", ">", CLng(MaxidDRQ)
                SSProcess.SetSelectCondition "SSObj_ID", "<", CLng(MaxidDRH) + 1
                SSProcess.SelectFilter
                geocount2 = SSProcess.GetSelGeoCount()
                For i2 = 0 To geoCount2 - 1
                    polygonID = SSProcess.GetSelGeoValue( i2, "SSObj_ID" )
                    Code = SSProcess.GetSelGeoValue( i2, "SSObj_Code" )
                    State = SSProcess.GetSelGeoValue( i2, "[State]" )
                    If state = "" Or state = "0" Or state = "NULL"Then
                        If InStr(ZDANDZRZCODE,Code) = 0  Then
                            SSProcess.SetObjectAttr polygonID, "[State]","2"
                        End If
                    End If
                Next
            End If
            If Replace (SWJfilename(j),"_None.edb","") <> SWJfilename(j) Then '判断是否为不变的地物
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.ClearSelectConditionGroups
                SSProcess.SetSelectCondition "SSObj_ID", ">", CLng(MaxidDRQ)
                SSProcess.SetSelectCondition "SSObj_ID", "<", CLng(MaxidDRH) + 1
                SSProcess.SelectFilter
                geocount3 = SSProcess.GetSelGeoCount()
                For i0 = 0 To geoCount3 - 1
                    polygonID = SSProcess.GetSelGeoValue( i0, "SSObj_ID" )
                    State = SSProcess.GetSelGeoValue( i0, "[State]" )
                    If state = "" Or state = "0"Or state = "NULL" Then
                        SSProcess.SetObjectAttr polygonID, "[State]","0"
                    End If
                Next
                'SSProcess.DeleteSelectionObj
            End If
        Next
        
        '删除宗地state=3的宗地面
        DeleteStatestree()
        '删除宗地state=3的自然幢
        DeleteStatestreezrz()
    End If
    
    '表名字符串
    Dim TableStr

    '参数初始化
    TableStr = "ZD_XM_信息表,FC_LPB_户信息表"

    TableMigration TableStr,ProJectName,ZhengHePath
    
End Sub


' /**
'  * @description: 表数据迁移,将OuterDataBase的TableNameStr表迁移至InnerDataBase
'  * @return {*} Void - 无返回值
'  * @param {String} TableNameStr - 需要迁移的表名字符串
'  * @param {String} OuterDataBase - 原始的数据库
'  * @param {String} InnerDataBase - 迁入的数据库
'  */
Function TableMigration(ByVal TableNameStr,ByVal OuterDataBase,ByVal InnerDataBase)
    
    TableArr = Split(TableNameStr,",", - 1,1)
    
    '打开输入数据库
    SSProcess.OpenAccessMdb InnerDataBase
    
    '清空原始表
    For i = 0 To UBound(TableArr)
        SqlStr = "Delete From " & TableArr(i)
        SSProcess.ExecuteSql SqlStr
    Next 'i
    
    '表迁移
    For i = 0 To UBound(TableArr)
        SSProcess.ImportMdbTable OuterDataBase,TableArr(i),TableArr(i),1
    Next 'i
    
    SSProcess.CloseAccessMdb InnerDataBase
    
    
End Function' TableMigration

'增量输入输出
Function Export(ByRef EDBName,ByRef ExportPath)
    
    CurrentPath = SSProcess.GetProjectFileName
    PathArr = Split(CurrentPath,"\", - 1,1)
    NamePos = UBound(PathArr)
    EDBName = PathArr(NamePos)
    
    For i = 0 To NamePos - 1
        If ExportPath = "" Then
            ExportPath = PathArr(i)
        Else
            ExportPath = ExportPath & "\" & PathArr(i)
        End If
    Next 'i
    
    ExportPath = ExportPath & "\过程文件\" & EDBName
    
    SSProcess.ClearDataXParameter
    SSProcess.SetDataXParameter "DataType","21"
    SSProcess.SetDataXParameter "FeatureCodeTBName","FeatureCodeTB_500"
    SSProcess.SetDataXParameter "SymbolScriptTBName","SymbolScriptTB_500"
    SSProcess.SetDataXParameter "NoteTemplateTBName","NoteTemplateTB_500"
    SSProcess.SetDataXParameter "UseUpdateDataStatus","1"
    SSProcess.SetDataXParameter "DataBoundMode","0"
    SSProcess.SetDataXParameter "SymbolExplodeMode","0"
    SSProcess.SetDataXParameter "ExportPathName",ExportPath
    'SSProcess.SetDataXParameter "LayerUseStatus","0"
    
    CurProjectName2 = SSProcess.GetImportFileName() '获取四文件路径
    StrProjectName = SSProcess.GetProjectFileName '获取当前工程文件名
    XGwz2 = InStrRev(CurProjectName,"\")  '获取“\”最后一次出现的位置
    CurProjectName2 = Left(CurProjectName,XGwz2)
    LJlenth = Len(StrProjectName)
    XGwz = InStrRev(StrProjectName,"\")  '获取“\”最后一次出现的位置
    wjm = Mid(StrProjectName,XGwz + 1, LJlenth)
    wjm = Replace(wjm,".edb","")
    CurProjectName1 = CurProjectName2 & wjm & ".edb"
    ZSGWFileName = Replace(CurProjectName1,".edb","_Add.edb")'四文件文件名
    IsFileExist ZSGWFileName,IsFolderExistMark
    If IsFolderExistMark = True Then DeleteFile ZSGWFileName
    ZSGWFileName = Replace(CurProjectName1,".edb","_Delete.edb")'四文件文件名
    IsFileExist ZSGWFileName,IsFolderExistMark
    If IsFolderExistMark = True Then DeleteFile ZSGWFileName
    ZSGWFileName = Replace(CurProjectName1,".edb","_Edit.edb")'四文件文件名
    IsFileExist ZSGWFileName,IsFolderExistMark
    If IsFolderExistMark = True Then DeleteFile ZSGWFileName
    ZSGWFileName = Replace(CurProjectName1,".edb","_None.edb")'四文件文件名
    IsFileExist ZSGWFileName,IsFolderExistMark
    If IsFolderExistMark = True Then DeleteFile ZSGWFileName
    
    SSProcess.ExportData
    
End Function' Export

'获取文件下指定格式edb
Function GetAllFiles(ByRef pathname, ByRef fileExt)
    Dim fso, folder, file, files, subfolder,folder0, fcount
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(pathname)
    Set files = folder.Files
    '查找文件
    For Each file In files
        extname = fso.GetExtensionName(file.name)
        If UCase(extname) = UCase(fileExt)  Then
            filenames(filecount) = pathname & file.name
            filecount = filecount + 1
        End If
    Next
    '查找子目录
    Set subfolder = folder.SubFolders
    For Each folder0 In subfolder
        GetAllFiles pathname & folder0.name & "\", fileExt
    Next
End Function


'删除文件
Function  DeleteFile (fpath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set MyFile = fso.GetFile(fpath)
    MyFile.Delete
    Set MyFile = Nothing
    Set fso = Nothing
End Function

'查找文件夹是否存在
Function IsFolderExist(FolderName,IsFolderExistMark)
    IsFolderExistMark = True
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.folderExists(FolderName) = False  Then
        IsFolderExistMark = False
    End If
    Set fso = Nothing
End Function

'查找文件是否存在
Function IsFileExist(FileName,IsFileExistMark)
    IsFileExistMark = True
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.fileExists(FileName) = False  Then
        IsFileExistMark = False
    End If
    Set fso = Nothing
End Function


Function DeleteStatestree()
    Dim vArray(100000),ALLXGZDDM
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130223"
    SSProcess.SetSelectCondition "[State]", "<>", "3"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount
    ALLXGZDDM = ""
    For i = 0 To geocount - 1
        XGgeoid = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        XGZDDM = SSProcess.GetSelGeoValue(i, "[BDCDYH]")
        XGZDDM = Left(XGZDDM,19)
        If ALLXGZDDM = "" Then
            ALLXGZDDM = XGZDDM
        Else
            ALLXGZDDM = ALLXGZDDM & "," & XGZDDM
        End If
    Next
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130223"
    SSProcess.SetSelectCondition "[State]", "==", "3"
    SSProcess.SelectFilter
    geocount1 = SSProcess.GetSelGeoCount
    For j = 0 To geocount1 - 1
        scgeoid = SSProcess.GetSelGeoValue(j, "SSObj_ID")
        SDZDDM = SSProcess.GetSelGeoValue(j, "[ZDDM]")
        SDZDDM = Left(SDZDDM,19)
        If Replace(ALLXGZDDM,SDZDDM,"") <> ALLXGZDDM  Then
            SSProcess.DeleteObject scgeoid
        End If
    Next
End Function


Function DeleteStatestreezrz()
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9210123"
    SSProcess.SetSelectCondition "[State]", "<>", "3"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount
    ALLBDCDYH = ""
    For i = 0 To geocount - 1
        XGgeoid = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        BDCDYH = SSProcess.GetSelGeoValue(i, "[BDCDYH]")
        If ALLBDCDYH = "" Then
            ALLBDCDYH = BDCDYH
        Else
            ALLBDCDYH = ALLBDCDYH & "," & BDCDYH
        End If
    Next
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9210123"
    SSProcess.SetSelectCondition "[State]", "==", "3"
    SSProcess.SelectFilter
    geocount1 = SSProcess.GetSelGeoCount
    For j = 0 To geocount1 - 1
        scgeoid = SSProcess.GetSelGeoValue(j, "SSObj_ID")
        SDBDCDYH = SSProcess.GetSelGeoValue(j, "[BDCDYH]")
        If InStr(ALLBDCDYH,SDBDCDYH) Then
            SSProcess.DeleteObject scgeoid
        End If
    Next
End Function

