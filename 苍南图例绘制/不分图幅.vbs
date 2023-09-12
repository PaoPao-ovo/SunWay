
'=========================================================图层名称配置=======================================================

LayStr = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"

MainCode = "54311203,54324004,54323004,54412004,54423004,54452004,54511004,54512114,54534114,54523114,54611114,54612004,54623004,54111003,54112003,54123003,54145003,54134003,54211003,54212003,54223003,54234003,54245003,54256003,54267003,54278003,54289003,54720114,54730114,54030003,54040003,51011203,52011203,53011204,53022204,53033204,53044204"

Table_LineName = "地下管线线属性表"

'===========================================功能入口========================================================

'总入口
Sub OnClick()
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "图廓层"
    SSProcess.SetSelectCondition "SSObj_Code", "==", "59999999"
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount
    If SelCount <= 0 Then
        MsgBox "不存在图廓"
        Exit Sub
    End If
    
    AllVisible
    
    DelFormerTl
    
    GetMapBorderPoision StandX,StandY
    
    LayArr = Split(LayStr,",", - 1,1)
    
    BorderStartX = StandX - 10 - 20 - 4
    BorderStartY = StandY - 10
    BorderEndX = StandX - 14 + 4
    FeatureY = BorderStartY - 2 - 2
    
    For i = 0 To UBound(LayArr)
        SelAllLine LayArr(i),CodeVal,CodeCount
        If CodeCount > 0 Then
            CodeArr = Split(CodeVal,",", - 1,1)
            For j = 0 To CodeCount - 1
                If j = 0 Then
                    DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    FeatureY = FeatureY - 2.25
                Else
                    DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    FeatureY = FeatureY - 2.25
                End If
            Next 'j
        End If
        SelAllPoi LayArr(i),CodeVal,CodeCount
        If CodeCount > 0 Then
            CodeArr = Split(CodeVal,",", - 1,1)
            For j = 0 To CodeCount - 1
                DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                FeatureY = FeatureY - 2.25
            Next 'j
        End If
    Next 'i
    
    DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
    
    Ending
    
End Sub' OnClick

'==============================================图例绘制==========================================================

'获取图廓的右上角坐标值
Function GetMapBorderPoision(ByRef X,ByRef Y)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "图廓层"
    SSProcess.SetSelectCondition "SSObj_Code", "==", "59999999"
    SSProcess.SelectFilter
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
    SSProcess.GetObjectPoint ID, 2, X, Y, Z, PointType, Name '左上角点坐标值
End Function' GetMapBorderPoision

'获取所有的线要素名称
Function SelAllLine(ByVal LayerName,ByRef CodeVal,ByRef CodeCount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SetSelectCondition "SSObj_Type", "==", "LINE"
    SSProcess.SetSelectCondition "SSObj_Code", "<>", "54100004,54200304,54245304,54256304,54267304,54412005,54423005,54452005,54111004,54211304,54400005,54411005,54212304,54223304,54120004,54130004,54140004,54150004,54234304,54278304,54289304,0"
    SSProcess.SelectFilter
    LineCount = SSProcess.GetSelGeoCount
    Dim k
    k = 0
    ReDim Preserve CodeStr(k)
    For i = 0 To LineCount - 1
        If SSProcess.GetSelGeoValue(i,"[FSFS]") = "架空" Or SSProcess.GetSelGeoValue(i,"[FSFS]") = "非开挖" Then
            CodeStr(k) = SSProcess.GetSelGeoValue(i,"SSObj_Code")
            k = k + 1
            ReDim Preserve CodeStr(k)
        ElseIf SSProcess.GetSelGeoValue(i,"[FSFS]") <> "井内连线" And SSProcess.GetSelGeoValue(i,"[YYKS]") <> "0" Then
            CodeStr(k) = SSProcess.GetSelGeoValue(i,"SSObj_Code")
            k = k + 1
            ReDim Preserve CodeStr(k)
        End If
    Next 'i
    DelRepeatLine CodeStr,CodeVal,CodeCount
    MainCodeArr = Split(MainCode,",", - 1,1)
    
    CodeArr = Split(CodeVal,",", - 1,1)
    For i = 0 To CodeCount - 1
        For j = 0 To UBound(MainCodeArr)
            If CodeArr(i) = MainCodeArr(j) Then
                Temp = CodeArr(0)
                CodeArr(0) = CodeArr(i)
                CodeArr(i) = Temp
            End If
        Next 'j
    Next 'i
    
    CodeVal = ""
    For i = 0 To UBound(CodeArr)
        If CodeVal = "" Then
            CodeVal = CodeArr(i)
        Else
            CodeVal = CodeVal & "," & CodeArr(i)
        End If
    Next 'i
End Function' SelAllLine

'获取所有的线要素名称
Function SelAllPoi(ByVal LayerName,ByRef CodeVal,ByRef CodeCount)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SetSelectCondition "SSObj_Type", "==", "Point"
    SSProcess.SetSelectCondition "SSObj_Code", "<>", "0"
    SSProcess.SelectFilter
    PoiCount = SSProcess.GetSelGeoCount
    ReDim CodeStr(PoiCount - 1)
    For i = 0 To PoiCount - 1
        CodeStr(i) = SSProcess.GetSelGeoValue(i,"SSObj_Code")
    Next 'i
    DelRepeat CodeStr,CodeVal,CodeCount
End Function' SelAllPoi

'绘制点注记
Function DrawPoint(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawPointNote X + 7,Y,Code,Width,Height
End Function

'绘制点注记名
Function DrawPointNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X + 2, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'绘制线注记
Function DrawLine(ByVal X1,ByVal X2,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X1, Y, 0, 0, ""
    SSProcess.AddNewObjPoint X2, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawLineNote X2 + 1,Y,Code,Width,Height
End Function

'绘制线注记名
Function DrawLineNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X + 4, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'区域框线绘制
Function DrawBorder(ByVal StartX,ByVal EndX,ByVal StartY,ByVal EndY)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", "51111111"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GroupId
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.AddNewObjPoint StartX,StartY,0,0,""
    SSProcess.AddNewObjPoint EndX, StartY,0,0,""
    SSProcess.AddNewObjPoint EndX,EndY,0, 0,""
    SSProcess.AddNewObjPoint StartX,EndY,0,0,""
    SSProcess.AddNewObjPoint StartX,StartY,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
    '绘制标题
    DrawTitle (StartX + EndX) / 2,StartY - 1,250,250
    
End Function

'绘制标题
Function DrawTitle(ByVal X,ByVal Y,ByVal Width, ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", "图 例"
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    'SSProcess.SetNewObjValue "SSObj_GroupID", GroupId
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'删除图例
Function DelFormerTl()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "管线图例层"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
End Function' DelFormerTl

'=================================================================工具函数=======================================================

Function AllVisible()
    Count = SSProcess.GetLayerCount
    For i = 0 To Count - 1
        SSProcess.SetLayerStatus SSProcess.GetLayerName(i), 1, 1
    Next
    SSProcess.RefreshView
End Function

'去除字符串中重复值(特殊)
Function DelRepeatLine(ByVal StrArr(),ByRef ToTalVal,ByRef LxCount)
    ToTalVal = ""
    For i = 0 To UBound(StrArr) - 1
        If ToTalVal = "" Then
            ToTalVal = "'" & StrArr(i) & "'"
        ElseIf Replace(ToTalVal,StrArr(i),"") = ToTalVal Then
            ToTalVal = ToTalVal & "," & "'" & StrArr(i) & "'"
            'MsgBox i & "----" & StrArr(i)
        End If
    Next 'i
    ToTalVal = Replace(ToTalVal,"'","")
    LxCount = UBound(Split(ToTalVal,",", - 1,1)) + 1
End Function' DelRepeatLine

'去除字符串中重复值
Function DelRepeat(ByVal StrArr(),ByRef ToTalVal,ByRef LxCount)
    ToTalVal = ""
    For i = 0 To UBound(StrArr)
        If ToTalVal = "" Then
            ToTalVal = "'" & StrArr(i) & "'"
        ElseIf Replace(ToTalVal,StrArr(i),"") = ToTalVal Then
            ToTalVal = ToTalVal & "," & "'" & StrArr(i) & "'"
            'MsgBox i & "----" & StrArr(i)
        End If
    Next 'i
    ToTalVal = Replace(ToTalVal,"'","")
    LxCount = UBound(Split(ToTalVal,",", - 1,1)) + 1
End Function' DelRepeat

Function Ending()
    MsgBox "生成完成"
End Function' Ending

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