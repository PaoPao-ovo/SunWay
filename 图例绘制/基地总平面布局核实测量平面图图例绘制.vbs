'全局变量
Dim vArray(1000)

'图例绘制函数
Function DrawTuLi()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420037 '图廓
    SSProcess.SelectFilter
    GeoCount = SSProcess.GetSelGeoCount()
    If Geocount > 0 Then
        For i = 0 To GeoCount - 1
            ID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            SSProcess.GetObjectPoint ID, 1, x, y, z, pointtype, name '右下角点
        Next
        innerids = SSProcess.SearchInnerObjIDs(ID,10,"9410001,9410011,9310032,9460091,9616201,8202002", 0)
        If innerids <> "" Then
            SSFunc.ScanString innerids, ",", vArray, nCount
            ZDrawCode = ""
            For j = 0 To nCount - 1
                DrawCode = SSProcess.GetObjectAttr(vArray(j), "SSObj_Code")
                DrawColor = SSProcess.GetObjectAttr(vArray(j), "SSObj_Color")
                DrawName = SSProcess.GetFeatureCodeInfo(DrawCode,"ObjectName")
                If ZDrawCode = "" Then
                    ZDrawCode = DrawCode
                    ZDrawColor = DrawColor
                    ZDrawName = DrawName
                Else
                    If Replace(ZDrawCode,DrawCode,"") = ZDrawCode Then
                        ZDrawCode = ZDrawCode & "," & DrawCode
                        ZDrawColor = ZDrawColor & "," & DrawColor
                        ZDrawName = ZDrawName & "," & Draw    Name
                    End If
                End If
            Next
            
            '绘制外框
            arDrawCode = Split(ZDrawCode,",")
            count = UBound(arDrawCode) + 4
            DrawBorder x,y,0,"RGB(255,255,255)",ID,count
            
            CreateWindows ZDrawCode,ZDrawColor,ZDrawName,Code,Color,Name
            
            If Code <> "" Then
                '绘制内部图例
                DrawInner x,y,ID,Code,Color,Name
                
                '绘制固定点注记
                DrawPoint x - 43,y + 11,"9000001",ID
            End If
            
        End If
    End If
End Function

'生成选择弹窗,返回选择的Code
Function CreateWindows(ByVal ZDrawCode,ByVal ZDrawColor,ByVal ZDrawName,ByRef Code,ByRef Color,ByRef Name)
    
    ZDrawCodeArr = Split(ZDrawCode,",", - 1,1)
    ReDim Preserve ZDrawCodeArr(UBound(ZDrawCodeArr) + 1)
    
    For i = 0 To UBound(ZDrawCodeArr)
        If i < UBound(ZDrawCodeArr) Then
            ZDrawCodeArr(i) = ZDrawCodeArr(i) & "【" & SSProcess.GetFeatureCodeInfo(ZDrawCodeArr(i),"ObjectName") & "】"
        Else
            ZDrawCodeArr(i) = ""
        End If
        
    Next 'i
    
    '制空
    Code = ""
    Color = ""
    Name = ""
    
    RecordShortListCount = UBound(ZDrawCodeArr) + 1
    ResVal_Dlg = SSFunc.SelectListAttr("选择列表","待选数据列表","选中数据列表",ZDrawCodeArr,RecordShortListCount)
    If ResVal_Dlg = 1 Then
        If RecordShortListCount > 0 Then
            Size = UBound(ZDrawCodeArr)
            If Size > 0 Then
                For i = 0 To RecordShortListCount - 1
                    If Code = "" Then
                        CodeArr = Split(ZDrawCodeArr(i),"【", - 1,1)
                        Code = CodeArr(0)
                        Color = SSProcess.GetFeatureCodeInfo(Code,"LineColor")
                        Name = SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
                    Else
                        CodeArr = Split(ZDrawCodeArr(i),"【", - 1,1)
                        Code = Code & "," & CodeArr(0)
                        Color = Color & "," & SSProcess.GetFeatureCodeInfo(CodeArr(0),"LineColor")
                        Name = Name & "," & SSProcess.GetFeatureCodeInfo(CodeArr(0),"ObjectName")
                    End If
                Next 'i
            End If
        End If
    End If
End Function' CreateWindows

'绘制图例
Function DrawInner(x,y,polygonID,ZDrawCode,ZDrawColor,ZDrawName)
    FountWith = 200
    FountHight = 200
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    arDrawName = Split(ZDrawName,",")
    count = UBound(arDrawCode) + 3
    
    '按数据绘制图例
    For i = 0 To count - 3
        DrawLine x - 20,y + 3 * (count - 3 - i) + 3,x - 15,y + 3 * (count - 3 - i) + 3,arDrawCode(i),arDrawColor(i),polygonID
        DrawNote x - 13,y + 3 * (count - 3 - i) + 3,arDrawCode(i),arDrawColor(i),FountWith,FountHight,arDrawName(i),polygonID
    Next
End Function

'绘制点要素
Function DrawPoint(x,y,code,polygonID)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工图廓"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'区域框线绘制
Function DrawBorder(x,y,code,color,polygonID,count)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    'SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工图廓"
    
    '内框线
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjPoint x, y + 3 * count , 0, 0, ""
    SSProcess.AddNewObjPoint x - 60, y + 3 * count , 0, 0, ""
    SSProcess.AddNewObjPoint x - 60, y , 0, 0, ""
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    
    '外框线
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjPoint x, y + 3 * count + 1 , 0, 0, ""
    SSProcess.AddNewObjPoint x - 61, y + 3 * count + 1 , 0, 0, ""
    SSProcess.AddNewObjPoint x - 61, y , 0, 0, ""
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
    '绘制标题
    DrawTitle x - 30,y + 3 * count - 2,400,400
End Function

'绘制面要素
Function DrawArea(x1,y1,x2,y2,x3,y3,x4,y4,code,color,polygonID)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工图廓"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'绘制线要素
Function DrawLine(x1,y1,x2,y2,code, color, polygonID)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工图廓"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'绘制注记
Function DrawNote(x, y, code, color, width, height, fontString,polygonID)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工图廓"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'绘制标题
Function DrawTitle(x, y, width, height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", "图 例"
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工图廓"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Sub OnClick()
    '添加代码
    DrawTuLi
End Sub