
'=========================================================图层名称配置=======================================================

LayStr = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"

'===========================================功能入口========================================================

'总入口
Sub OnClick()
    
    DelTk
    
    GxVisible LayStr
    
    CreatMap
    
    Ending
    
    AllVisible
    
    DelTk
    
End Sub

'生成图廓和图例
Function CreatMap()
    SSProcess.CreateMapFrame
    FrameCount = SSProcess.GetMapFrameCount()
    For i = 0 To FrameCount - 1
        SSProcess.GetMapFrameCenterPoint i, CenterX, CenterY
        SSProcess.SetFrameCode("9130225")
        SSProcess.SetCurMapFrame CenterX, CenterY, 0, ""
        CreateNote SSProcess.GetCurMapFrame()
    Next
    SSProcess.SaveBufferObjToDatabase
    SSProcess.MapMethod "LoadData","图廓层"
    SSProcess.FreeMapFrame
End Function

Function CreateNote(ByVal MapId)
    
    SSProcess.GetObjectPoint MapId, 2, StandX, StandY, StandZ, PointType, Name '左上角点坐标值
    
    BorderStartX = StandX - 10 - 20
    BorderStartY = StandY - 10
    BorderEndX = StandX - 14
    FeatureY = BorderStartY - 2 - 2
    
    SelAll MapId,CodeVal,CodeCount
    
    If CodeCount > 0 Then
        CodeArr = Split(CodeVal,",", - 1,1)
        For j = 0 To CodeCount - 1
            If SSProcess.GetFeatureCodeInfo(CodeArr(j),"Type") = 0 Then
                DrawPoint BorderStartX + 3.5,FeatureY,CodeArr(j)
                FeatureY = FeatureY - 2.25
            Else
                DrawLine BorderStartX + 2,BorderStartX + 5,FeatureY,CodeArr(j)
                FeatureY = FeatureY - 2.25
            End If
        Next 'j
    End If
    
    DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
    
End Function' CreateNote

'获取所有的点和线要素名称
Function SelAll(ByVal OuterId,ByRef CodeVal,ByRef CodeCount)
    PoiIds = SSProcess.SearchInPolyObjIDs(OuterId,0,"",0,1,1)
    LinIds = SSProcess.SearchInPolyObjIDs(OuterId,1,"",0,1,1)
    PoiArr = Split(PoiIds,",", - 1,1)
    LinArr = Split(LinIds,",", - 1,1)
    For i = 0 To UBound(PoiArr)
        Select Case SSProcess.GetObjectAttr(PoiArr(i),"SSObj_LayerName")
            
            Case "CD"
            If CDCodeStr = "" Then
                CDCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CDCodeStr = CDCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            
            Case "CT"
            If CTCodeStr = "" Then
                CTCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CTCodeStr = CTCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "CY"
            If CYCodeStr = "" Then
                CYCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CYCodeStr = CYCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "CQ"
            If CQCodeStr = "" Then
                CQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CQCodeStr = CQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "CS"
            If CSCodeStr = "" Then
                CSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                CSCodeStr = CSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "QT"
            If QTCodeStr = "" Then
                QTCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                QTCodeStr = QTCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "BM"
            If BMCodeStr = "" Then
                BMCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                BMCodeStr = BMCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "FQ"
            If FQCodeStr = "" Then
                FQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                FQCodeStr = FQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "DL"
            If DLCodeStr = "" Then
                DLCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                DLCodeStr = DLCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "GD"
            If GDCodeStr = "" Then
                GDCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                GDCodeStr = GDCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "LD"
            If LDCodeStr = "" Then
                LDCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                LDCodeStr = LDCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "DC"
            If DCCodeStr = "" Then
                DCCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                DCCodeStr = DCCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "XH"
            If XHCodeStr = "" Then
                XHCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                XHCodeStr = XHCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "TX"
            If TXCodeStr = "" Then
                TXCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                TXCodeStr = TXCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "DX"
            If DXCodeStr = "" Then
                DXCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                DXCodeStr = DXCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "YD"
            If YDCodeStr = "" Then
                YDCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                YDCodeStr = YDCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "LT"
            If LTCodeStr = "" Then
                LTCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                LTCodeStr = LTCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "JX"
            If JXCodeStr = "" Then
                JXCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                JXCodeStr = JXCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "JK"
            If JKCodeStr = "" Then
                JKCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                JKCodeStr = JKCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "DS"
            If DSCodeStr = "" Then
                DSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                DSCodeStr = DSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "BZ"
            If BZCodeStr = "" Then
                BZCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                BZCodeStr = BZCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "JS"
            If JSCodeStr = "" Then
                JSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                JSCodeStr = JSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "XF"
            If XFCodeStr = "" Then
                XFCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                XFCodeStr = XFCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "PS"
            If PSCodeStr = "" Then
                PSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                PSCodeStr = PSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "YS"
            If YSCodeStr = "" Then
                YSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                YSCodeStr = YSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "WS"
            If WSCodeStr = "" Then
                WSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                WSCodeStr = WSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "FS"
            If FSCodeStr = "" Then
                FSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                FSCodeStr = FSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "RQ"
            If RQCodeStr = "" Then
                RQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                RQCodeStr = RQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "MQ"
            If MQCodeStr = "" Then
                MQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                MQCodeStr = MQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "YH"
            If YHCodeStr = "" Then
                YHCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                YHCodeStr = YHCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "RL"
            If RLCodeStr = "" Then
                RLCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                RLCodeStr = RLCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "RS"
            If RSCodeStr = "" Then
                RSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                RSCodeStr = RSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "ZQ"
            If ZQCodeStr = "" Then
                ZQCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                ZQCodeStr = ZQCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "SY"
            If SYCodeStr = "" Then
                SYCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                SYCodeStr = SYCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "GS"
            If GSCodeStr = "" Then
                GSCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                GSCodeStr = GSCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "EX"
            If EXCodeStr = "" Then
                EXCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                EXCodeStr = EXCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
            Case "TR"
            If TRCodeStr = "" Then
                TRCodeStr = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            Else
                TRCodeStr = TRCodeStr & "," & SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")
            End If
        End Select
    Next 'i
    
    For i = 0 To UBound(LinArr)
        Select Case SSProcess.GetObjectAttr(LinArr(i),"SSObj_LayerName")
            
            Case "CD"
            If CDCodeStr = "" Then
                CDCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                CDCodeStr = CDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            
            Case "CT"
            If CTCodeStr = "" Then
                CTCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                CTCodeStr = CTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "CY"
            If CYCodeStr = "" Then
                CYCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                CYCodeStr = CYCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "CQ"
            If CQCodeStr = "" Then
                CQCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                CQCodeStr = CQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "CS"
            If CSCodeStr = "" Then
                CSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                CSCodeStr = CSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "QT"
            If QTCodeStr = "" Then
                QTCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                QTCodeStr = QTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "BM"
            If BMCodeStr = "" Then
                BMCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                BMCodeStr = BMCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "FQ"
            If FQCodeStr = "" Then
                FQCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                FQCodeStr = FQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "DL"
            If DLCodeStr = "" Then
                DLCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                DLCodeStr = DLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "GD"
            If GDCodeStr = "" Then
                GDCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                GDCodeStr = GDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "LD"
            If LDCodeStr = "" Then
                LDCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                LDCodeStr = LDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "DC"
            If DCCodeStr = "" Then
                DCCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                DCCodeStr = DCCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "XH"
            If XHCodeStr = "" Then
                XHCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                XHCodeStr = XHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "TX"
            If TXCodeStr = "" Then
                TXCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                TXCodeStr = TXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "DX"
            If DXCodeStr = "" Then
                DXCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                DXCodeStr = DXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "YD"
            If YDCodeStr = "" Then
                YDCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                YDCodeStr = YDCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "LT"
            If LTCodeStr = "" Then
                LTCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                LTCodeStr = LTCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "JX"
            If JXCodeStr = "" Then
                JXCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                JXCodeStr = JXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "JK"
            If JKCodeStr = "" Then
                JKCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                JKCodeStr = JKCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "DS"
            If DSCodeStr = "" Then
                DSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                DSCodeStr = DSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "BZ"
            If BZCodeStr = "" Then
                BZCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                BZCodeStr = BZCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "JS"
            If JSCodeStr = "" Then
                JSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                JSCodeStr = JSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "XF"
            If XFCodeStr = "" Then
                XFCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                XFCodeStr = XFCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "PS"
            If PSCodeStr = "" Then
                PSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                PSCodeStr = PSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "YS"
            If YSCodeStr = "" Then
                YSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                YSCodeStr = YSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "WS"
            If WSCodeStr = "" Then
                WSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                WSCodeStr = WSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "FS"
            If FSCodeStr = "" Then
                FSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                FSCodeStr = FSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "RQ"
            If RQCodeStr = "" Then
                RQCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                RQCodeStr = RQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "MQ"
            If MQCodeStr = "" Then
                MQCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                MQCodeStr = MQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "YH"
            If YHCodeStr = "" Then
                YHCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                YHCodeStr = YHCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "RL"
            If RLCodeStr = "" Then
                RLCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                RLCodeStr = RLCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "RS"
            If RSCodeStr = "" Then
                RSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                RSCodeStr = RSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "ZQ"
            If ZQCodeStr = "" Then
                ZQCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                ZQCodeStr = ZQCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "SY"
            If SYCodeStr = "" Then
                SYCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                SYCodeStr = SYCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "GS"
            If GSCodeStr = "" Then
                GSCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                GSCodeStr = GSCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "EX"
            If EXCodeStr = "" Then
                EXCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                EXCodeStr = EXCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
            Case "TR"
            If TRCodeStr = "" Then
                TRCodeStr = SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            Else
                TRCodeStr = TRCodeStr & "," & SSProcess.GetObjectAttr(LinArr(i),"SSObj_Code")
            End If
        End Select
    Next 'i

    ' ReDim CodeStr(UBound(PoiArr) + UBound(LinArr))
    ' For i = 0 To UBound(PoiArr) + UBound(LinArr)
    '     If i <= UBound(PoiArr) Then
    '         CodeStr(i) = SSProcess.GetObjectAttr(PoiArr(i),"SSObj_Code")      
    '     Else
    '         CodeStr(i) = SSProcess.GetObjectAttr(LinArr(i - UBound(PoiArr) ),"SSObj_Code")
    '     End If
    ' Next 'i

    CodeNameVal = CDCodeStr & ";" & CTCodeStr & ";" & CYCodeStr & ";" & CQCodeStr & ";" & CSCodeStr & ";" & QTCodeStr & ";" & BMCodeStr & ";" & FQCodeStr & ";" & DLCodeStr & ";" & GDCodeStr & ";" & LDCodeStr & ";" & DCCodeStr & ";" & XHCodeStr & ";" & TXCodeStr & ";" & DXCodeStr & ";" & YDCodeStr & ";" & LTCodeStr & ";" & JXCodeStr & ";" & JKCodeStr & ";" & DSCodeStr & ";" & BZCodeStr & ";" & JSCodeStr & ";" & XFCodeStr & ";" & PSCodeStr & ";" & YSCodeStr & ";" & WSCodeStr & ";" & FSCodeStr & ";" & RQCodeStr & ";" & MQCodeStr & ";" & YHCodeStr & ";" & RLCodeStr & ";" & RSCodeStr & ";" & ZQCodeStr & ";" & SYCodeStr & ";" & GSCodeStr & ";" & EXCodeStr & ";" & TRCodeStr
    
    CodeNameArr = Split(CodeNameVal,";", - 1,1)
    For i = 0 To UBound(CodeNameArr)
        If CodeNameArr(i) <> "" Then
            If TempCodeStr = "" Then
                TempCodeStr = CodeNameArr(i)
            Else
                TempCodeStr = TempCodeStr & "," & CodeNameArr(i)
            End If
        End If
    Next 'i
    CodeStr = Split(TempCodeStr,",", - 1,1)
    DelRepeat CodeStr,CodeVal,CodeCount
End Function' SelAllPoi

'绘制点注记
Function DrawPoint(ByVal X,ByVal Y,ByVal Code)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawPointNote X + 2.5,Y,Code,150,150
End Function

'绘制点注记名
Function DrawPointNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'绘制线注记
Function DrawLine(ByVal X1,ByVal X2,ByVal Y,ByVal Code)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X1, Y, 0, 0, ""
    SSProcess.AddNewObjPoint X2, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawLineNote X2 + 1,Y,Code,150,150
End Function

'绘制线注记名
Function DrawLineNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "管线图例层"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
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
    DrawTitle (StartX + EndX) / 2,StartY - 1,200,200
    
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

'选择指定地物并返回个数
Function SelFeatures(ByRef Count,ByRef IdArr()) 'EngLayerName 图层名称(英文),Count 个数(返回值),IdArr() Id数组(返回值)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "图廓层"
    SSProcess.SetSelectCondition "SSObj_Type", "==", "Area"
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
    ReDim IdArr(Count)
    For i = 0 To Count - 1
        IdArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' SelFeatures

'去除字符串中重复值
Function DelRepeat(ByVal StrArr(),ByRef ToTalVal,ByRef LxCount)
    ToTalVal = ""
    For i = 0 To UBound(StrArr)
        If ToTalVal = "" Then
            ToTalVal = "'" & StrArr(i) & "'"
        ElseIf Replace(ToTalVal,StrArr(i),"") = ToTalVal Then
            ToTalVal = ToTalVal & "," & "'" & StrArr(i) & "'"
        End If
    Next 'i
    ToTalVal = Replace(ToTalVal,"'","")
    LxCount = UBound(Split(ToTalVal,",", - 1,1)) + 1
End Function' DelRepeat

Function GxVisible(ByVal LayString)
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 0
    Next
    LayArr = Split(LayString,",", - 1,1)
    For i = 0 To UBound(LayArr)
        SSProcess.SetLayerStatus LayArr(i), 1, 1
    Next 'i
    SSProcess.SetLayerStatus "图廓层", 1, 1
    SSProcess.RefreshView
End Function

Function DelTk()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "管线图例层"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "图廓层"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
End Function' DelTk

Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'完成提示
Function Ending()
    MsgBox "输出完成"
End Function' Ending