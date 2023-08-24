
'========================================================文件路径操作对象================================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'=============================================================功能入口=======================================================================

Sub OnClick()
    
    GetAllFildValue ZStr
    
    ZStrArr = Split(ZStr,",", - 1,1)
    
    For i = 0 To UBound(ZStrArr)
        If i = 0  Then
            GetFeatureIdStr ZStrArr(i),IdStr
            GetFour IdStr,MinX,MinY,MaxX,MaxY
            RightX = MaxX
            BottomY = MinY
        Else
            GetFeatureIdStr ZStrArr(i),IdStr
            GetFour IdStr,MinX,MinY,MaxX,MaxY
            OffSet IdStr,RightX,BottomY,MinX,MinY,MaxX,NextRigthX
            RightX = NextRigthX
        End If
    Next 'i
    
    BoderMinX = ""
    BoderMinY = ""
    BoderMaxX = ""
    BoderMaxY = ""
    
    For i = 0 To UBound(ZStrArr)
        GetFeatureIdStr ZStrArr(i),IdStr
        GetFour IdStr,MinX,MinY,MaxX,MaxY
        If BoderMinX = "" Then
            BoderMinX = MinX
            BoderMinY = MinY
            BoderMaxX = MaxX
            BoderMaxY = MaxY
        Else
            
            If MinX < BoderMinX Then
                BoderMinX = MinX
            Else
                BoderMinX = BoderMinX
            End If
            
            If MaxX > BoderMaxX Then
                BoderMaxX = MaxX
            Else
                BoderMaxX = BoderMaxX
            End If
            
            If MinY < BoderMaxX Then
                BoderMinY = BoderMinY
            Else
                BoderMinY = BoderMinY
            End If
            
            If MaxY > BoderMaxY Then
                BoderMaxY = MaxY
            Else
                BoderMaxY = BoderMaxY
            End If
            
        End If
    Next 'i
    
    DrawArea BoderMinX - 1,BoderMinY - 1,BoderMaxX + 1,BoderMinY - 1,BoderMaxX + 1,BoderMaxY + 1,BoderMinX - 1,BoderMaxY + 1
    
    Path = SSProcess.GetSysPathName(7) & "Pictures\"
    StrBmpFile = Path & "立面图" & ".wmf"
    Dpi = 300
    
    SSFunc.DrawToImage BoderMinX - 1,BoderMinY - 1,BoderMaxX + 1,BoderMaxY + 1,"297X100",Dpi,StrBmpFile
    
End Sub' OnClick   

'获取某一幢四至
Function GetFour(ByVal IdStr,ByRef MinX,ByRef MinY,ByRef MaxX,ByRef MaxY)
    
    MinX = ""
    MinY = ""
    MaxX = ""
    MaxY = ""
    
    IdArr = Split(IdStr,",", - 1,1)
    
    For i = 0 To UBound(IdArr)
        TypeValue = SSProcess.GetObjectAttr(IdArr(i),"SSObj_LineType")
        If TypeValue = "1" Then
            Pointcount = SSProcess.GetObjectAttr(IdArr(i),"SSObj_PointCount")
            For j = 0 To Pointcount - 1
                SSProcess.GetObjectPoint IdArr(i),j,X,Y,Z,PType,PName
                If MinX = "" Then
                    MinX = X
                    MinY = Y
                    MaxX = X
                    MaxY = Y
                Else
                    If X > MaxX  Then
                        MaxX = X
                    Else
                        MaxX = MaxX
                    End If
                    
                    If X < MinX  Then
                        MinX = X
                    Else
                        MinX = MinX
                    End If
                    
                    If Y > MaxY  Then
                        MaxY = Y
                    Else
                        MaxY = MaxY
                    End If
                    
                    If Y < MinY  Then
                        MinY = Y
                    Else
                        MinY = MinY
                    End If
                End If
            Next 'j
        Else
            PointX = Transform(SSProcess.GetObjectAttr(IdArr(i),"SSObj_X"))
            PointY = Transform(SSProcess.GetObjectAttr(IdArr(i),"SSObj_Y"))
            If MinX = "" Then
                MinX = PointX
                MinY = PointY
                MaxX = PointX
                MaxY = PointY
            Else
                If PointX > MaxX  Then
                    MaxX = PointX
                Else
                    MaxX = MaxX
                End If
                
                If PointX < MinX  Then
                    MinX = PointX
                Else
                    MinX = MinX
                End If
                
                If PointY > MaxY  Then
                    MaxY = PointY
                Else
                    MaxY = MaxY
                End If
                
                If PointY < MinY  Then
                    MinY = PointY
                Else
                    MinY = MinY
                End If
            End If
        End If
    Next 'i
    
End Function' GetFour

'偏移
Function OffSet(ByVal IdStr,ByVal RightX,ByVal BottomY,ByVal MinX,ByVal MinY,ByVal MaxX,ByRef NextRigthX)
    
    RightX = RightX + 10
    
    XLength = Sqr((MinX - RightX) ^ 2)
    YLength = Sqr((MinY - BottomY) ^ 2)
    
    If MinX > RightX Then
        XLength =  - XLength
    Else
        XLength = XLength
    End If
    
    If MinY > BottomY Then
        YLength =  - YLength
    Else
        YLength = YLength
    End If
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "==", IdStr
    SSProcess.SelectFilter
    
    SSProcess.OffsetSelectionObj XLength,YLength,0
    
    NextRigthX = MaxX + XLength
    
End Function' OffSet

'获取所有的字段名称
Function GetAllFildValue(ByRef ZStr)
    SqlStr = "Select DISTINCT JG_立面图线属性表.ID_ZRZ From JG_立面图线属性表 INNER JOIN GeoLineTB ON JG_立面图线属性表.ID = GeoLineTB.ID WHERE ([GeoLineTB].[Mark] Mod 2)<>0"
    GetSQLRecordAll SqlStr,ZStrArr,ValCount
    
    If ValCount > 0 Then
        For i = 0 To ValCount - 1
            If ZStr = "" Then
                If ZStrArr(i) <> "" Then
                    ZStr = ZStrArr(i)
                End If
            Else
                If ZStrArr(i) <> "" Then
                    ZStr = ZStr & "," & ZStrArr(i)
                End If
            End If
        Next 'i
    End If
End Function' GetAllFildValue

'获取某一幢所有要素ID
Function GetFeatureIdStr(ByVal ZStr,ByRef IdStr)
    
    IdStr = ""
    
    If ZStr <> "" Then
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Type", "==", "POINT,LINE,NOTE"
        SSProcess.SetSelectCondition "[ID_ZRZ]", "==", ZStr
        SSProcess.SelectFilter
        GeoCount = SSProcess.GetSelGeoCount()
        
        For i = 0 To GeoCount - 1
            If IdStr = "" Then
                IdStr = SSProcess.GetSelGeoValue(i,"SSObj_ID")
            Else
                IdStr = IdStr & "," & SSProcess.GetSelGeoValue(i,"SSObj_ID")
            End If
        Next 'i
        
    End If
End Function' GetFeatureIdStr


'判断文件是否存在
Function IsFileExist(ByVal FilePath)
    
    IsFileExist = False
    
    If FileSysObj.FileExists(File) Then
        IsFileExist = True
    End If
    
End Function' FileExists

'删除文件
Function DeleteFile(ByVal FilePath)
    
    FileSysObj.DeleteFile FilePath
    
End Function

'数据类型转换
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

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
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

'绘制面要素
Function DrawArea(ByVal X1,ByVal Y1,ByVal X2,ByVal Y2,ByVal X3,ByVal Y3,ByVal X4,ByVal Y4)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code",2
    SSProcess.AddNewObjPoint X1, Y1, 0, 0, ""
    SSProcess.AddNewObjPoint X2, Y2, 0, 0, ""
    SSProcess.AddNewObjPoint X3, Y3, 0, 0, ""
    SSProcess.AddNewObjPoint X4, Y4, 0, 0, ""
    SSProcess.AddNewObjPoint X1, Y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function