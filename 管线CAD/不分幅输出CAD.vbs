
MainCode = "54311203,54324004,54323004,54412004,54423004,54452004,54511004,54512114,54534114,54523114,54611114,54612004,54623004,54111003,54112003,54123003,54145003,54134003,54211003,54212003,54223003,54234003,54245003,54256003,54267003,54278003,54289003,54720114,54730114,54030003,54040003,51011203,52011203,53011204,53022204,53033204,53044204"

'=======================================================功能入口=========================================================

'总入口
Sub OnClick()
    ConFirmWay Way,res,GroupStr
    If res = 1 Then
        If Way = "综合管线图" Then
            AllVisible
            ModifyAttr "59999999",Way,TkId,XmMc,Count
            If Count = 0 Then
                MsgBox "不存在图廓"
                Exit Sub
            End If
            ' GetMapBorderPoision StandX,StandY
            ' LayStr = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"
            ' LayArr = Split(LayStr,",", - 1,1)
            ' BorderStartX = StandX - 10 - 20
            ' BorderStartY = StandY - 10
            ' BorderEndX = StandX - 14
            ' FeatureY = BorderStartY - 2 - 2
            ' DelFormerTl
            ' For i = 0 To UBound(LayArr)
            '     SelAllLine LayArr(i),CodeVal,CodeCount
            '     If CodeCount > 0 Then
            '         CodeArr = Split(CodeVal,",", - 1,1)
            '         For j = 0 To CodeCount - 1
            '             If j = 0 Then
            '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
            '                 FeatureY = FeatureY - 2.25
            '             Else
            '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
            '                 FeatureY = FeatureY - 2.25
            '             End If
            '         Next 'j
            '     End If
            '     SelAllPoi LayArr(i),CodeVal,CodeCount
            '     If CodeCount > 0 Then
            '         CodeArr = Split(CodeVal,",", - 1,1)
            '         For j = 0 To CodeCount - 1
            '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
            '             FeatureY = FeatureY - 2.25
            '         Next 'j
            '     End If
            ' Next 'i
            ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
            FilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & "综合管线图.dwg"
            DelFormerTl
            SZDWT TkId,FilePath
            AllVisible
            
        ElseIf Way = "分层输出" Then
            GetAllLayerName SmallArr,BigArr
            For k = 0 To UBound(BigArr)
                Select Case BigArr(k)
                    Case "长输输电"
                    AllDisVisible
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "CD"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "CD", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "CD注记", 1, 1
                    SSProcess.SetLayerStatus "CD点注记", 1, 1
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "长输通信"
                    AllDisVisible
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "CT"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "CT", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "CT注记", 1, 1
                    
                    SSProcess.SetLayerStatus "CT点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "长输油气水"
                    AllDisVisible
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "CY,CQ,CS,QT"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "CY", 1, 1
                    SSProcess.SetLayerStatus "CQ", 1, 1
                    SSProcess.SetLayerStatus "CS", 1, 1
                    SSProcess.SetLayerStatus "QT", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "CY注记", 1, 1
                    SSProcess.SetLayerStatus "CQ注记", 1, 1
                    SSProcess.SetLayerStatus "CS注记", 1, 1
                    SSProcess.SetLayerStatus "QT注记", 1, 1
                    
                    SSProcess.SetLayerStatus "CY点注记", 1, 1
                    SSProcess.SetLayerStatus "CQ点注记", 1, 1
                    SSProcess.SetLayerStatus "CS点注记", 1, 1
                    SSProcess.SetLayerStatus "QT点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    
                    Case "城市管线"
                    AllDisVisible
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "BM,FQ"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "BM", 1, 1
                    SSProcess.SetLayerStatus "FQ", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "BM注记", 1, 1
                    SSProcess.SetLayerStatus "FQ注记", 1, 1
                    
                    SSProcess.SetLayerStatus "BM点注记", 1, 1
                    SSProcess.SetLayerStatus "FQ点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "电力"
                    AllDisVisible
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "DL,GD,LD,DC,XH"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "XH", 1, 1
                    SSProcess.SetLayerStatus "DC", 1, 1
                    SSProcess.SetLayerStatus "LD", 1, 1
                    SSProcess.SetLayerStatus "GD", 1, 1
                    SSProcess.SetLayerStatus "DL", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "XH注记", 1, 1
                    SSProcess.SetLayerStatus "DC注记", 1, 1
                    SSProcess.SetLayerStatus "LD注记", 1, 1
                    SSProcess.SetLayerStatus "GD注记", 1, 1
                    SSProcess.SetLayerStatus "DL注记", 1, 1
                    
                    SSProcess.SetLayerStatus "XH点注记", 1, 1
                    SSProcess.SetLayerStatus "DC点注记", 1, 1
                    SSProcess.SetLayerStatus "LD点注记", 1, 1
                    SSProcess.SetLayerStatus "GD点注记", 1, 1
                    SSProcess.SetLayerStatus "DL点点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "通信"
                    AllDisVisible
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "DX,YD,LT,JX,JK,EX,DS,BZ,TX"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "BZ", 1, 1
                    SSProcess.SetLayerStatus "DX", 1, 1
                    SSProcess.SetLayerStatus "YD", 1, 1
                    SSProcess.SetLayerStatus "LT", 1, 1
                    SSProcess.SetLayerStatus "JX", 1, 1
                    SSProcess.SetLayerStatus "JK", 1, 1
                    SSProcess.SetLayerStatus "EX", 1, 1
                    SSProcess.SetLayerStatus "DS", 1, 1
                    SSProcess.SetLayerStatus "TX", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "BZ注记", 1, 1
                    SSProcess.SetLayerStatus "DX注记", 1, 1
                    SSProcess.SetLayerStatus "YD注记", 1, 1
                    SSProcess.SetLayerStatus "LT注记", 1, 1
                    SSProcess.SetLayerStatus "JX注记", 1, 1
                    SSProcess.SetLayerStatus "JK注记", 1, 1
                    SSProcess.SetLayerStatus "EX注记", 1, 1
                    SSProcess.SetLayerStatus "DS注记", 1, 1
                    SSProcess.SetLayerStatus "TX注记", 1, 1
                    
                    SSProcess.SetLayerStatus "BZ点注记", 1, 1
                    SSProcess.SetLayerStatus "DX点注记", 1, 1
                    SSProcess.SetLayerStatus "YD点注记", 1, 1
                    SSProcess.SetLayerStatus "LT点注记", 1, 1
                    SSProcess.SetLayerStatus "JX点注记", 1, 1
                    SSProcess.SetLayerStatus "JK点注记", 1, 1
                    SSProcess.SetLayerStatus "EX点注记", 1, 1
                    SSProcess.SetLayerStatus "DS点注记", 1, 1
                    SSProcess.SetLayerStatus "TX点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "给水"
                    AllDisVisible
                    'GetMapBorderPoision StandX,StandY
                    ' LayStr = "JS,XF"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "JS", 1, 1
                    SSProcess.SetLayerStatus "XF", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "JS注记", 1, 1
                    SSProcess.SetLayerStatus "XF注记", 1, 1
                    
                    SSProcess.SetLayerStatus "JS点注记", 1, 1
                    SSProcess.SetLayerStatus "XF点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "排水"
                    AllDisVisible
                    'GetMapBorderPoision StandX,StandY
                    ' LayStr = "FS,WS,YS,PS"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "FS", 1, 1
                    SSProcess.SetLayerStatus "WS", 1, 1
                    SSProcess.SetLayerStatus "YS", 1, 1
                    SSProcess.SetLayerStatus "PS", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "FS注记", 1, 1
                    SSProcess.SetLayerStatus "WS注记", 1, 1
                    SSProcess.SetLayerStatus "YS注记", 1, 1
                    SSProcess.SetLayerStatus "PS注记", 1, 1
                    
                    SSProcess.SetLayerStatus "FS点注记", 1, 1
                    SSProcess.SetLayerStatus "WS点注记", 1, 1
                    SSProcess.SetLayerStatus "YS点注记", 1, 1
                    SSProcess.SetLayerStatus "PS点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "燃气"
                    AllDisVisible
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "RQ,MQ,TR,YH"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "YH", 1, 1
                    SSProcess.SetLayerStatus "MQ", 1, 1
                    SSProcess.SetLayerStatus "TR", 1, 1
                    SSProcess.SetLayerStatus "RQ", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "YH注记", 1, 1
                    SSProcess.SetLayerStatus "MQ注记", 1, 1
                    SSProcess.SetLayerStatus "TR注记", 1, 1
                    SSProcess.SetLayerStatus "RQ注记", 1, 1
                    
                    SSProcess.SetLayerStatus "YH点注记", 1, 1
                    SSProcess.SetLayerStatus "MQ点注记", 1, 1
                    SSProcess.SetLayerStatus "TR点注记", 1, 1
                    SSProcess.SetLayerStatus "RQ点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "热力"
                    AllDisVisible
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "RL,RS,ZQ"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "ZQ", 1, 1
                    SSProcess.SetLayerStatus "RL", 1, 1
                    SSProcess.SetLayerStatus "RS", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "ZQ注记", 1, 1
                    SSProcess.SetLayerStatus "RL注记", 1, 1
                    SSProcess.SetLayerStatus "RS注记", 1, 1
                    
                    SSProcess.SetLayerStatus "ZQ点注记", 1, 1
                    SSProcess.SetLayerStatus "RL点注记", 1, 1
                    SSProcess.SetLayerStatus "RS点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "工业"
                    AllDisVisible
                    ' GetMapBorderPoision StandX,StandY
                    ' LayStr = "SY,GS"
                    ' LayArr = Split(LayStr,",", - 1,1)
                    ' BorderStartX = StandX - 10 - 20
                    ' BorderStartY = StandY - 10
                    ' BorderEndX = StandX - 14
                    ' FeatureY = BorderStartY - 2 - 2
                    SSProcess.SetLayerStatus "GS", 1, 1
                    SSProcess.SetLayerStatus "SY", 1, 1
                    SSProcess.SetLayerStatus "TK", 1, 1
                    SSProcess.SetLayerStatus "GS注记", 1, 1
                    SSProcess.SetLayerStatus "SY注记", 1, 1
                    
                    SSProcess.SetLayerStatus "GS点注记", 1, 1
                    SSProcess.SetLayerStatus "SY点注记", 1, 1
                    
                    SSProcess.SetLayerStatus "管线图例层", 1, 1
                    SSProcess.SetLayerStatus "图廓层", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "不存在图廓"
                        Exit Sub
                    End If
                    ' For i = 0 To UBound(LayArr)
                    '     SelAllLine LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             If j = 0 Then
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),200,200
                    '                 FeatureY = FeatureY - 2.25
                    '             Else
                    '                 DrawLine BorderStartX + 2 ,BorderStartX + 5 + 4,FeatureY,CodeArr(j),150,150
                    '                 FeatureY = FeatureY - 2.25
                    '             End If
                    '         Next 'j
                    '     End If
                    '     SelAllPoi LayArr(i),CodeVal,CodeCount
                    '     If CodeCount > 0 Then
                    '         CodeArr = Split(CodeVal,",", - 1,1)
                    '         For j = 0 To CodeCount - 1
                    '             DrawPoint BorderStartX + 5,FeatureY,CodeArr(j),150,150
                    '             FeatureY = FeatureY - 2.25
                    '         Next 'j
                    '     End If
                    ' Next 'i
                    ' DrawBorder BorderStartX,BorderEndX,BorderStartY,FeatureY
                    FFilePath = SSProcess.GetSysPathName(5) & "专业管线图\" & XmMc & BigArr(k) & "地下管线图.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                End Select
            Next 'z
        End If
        FYNOTE GroupStr
        Ending
    Else
        MsgBox "已退出"
    End If
    AllVisible
    
End Sub' OnClick

Function FYNOTE(STR)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "CD点注记,CT点注记,CY点注记,CQ点注记,CS点注记,QT点注记,BM点注记,FQ点注记,DL点注记,GD点注记,LD点注记,DC点注记,XH点注记,TX点注记,DX点注记,YD点注记,LT点注记,JX点注记,JK点注记,EX点注记,DS点注记,BZ点注记,JS点注记,XF点注记,PS点注记,YS点注记,WS点注记,FS点注记,RQ点注记,MQ点注记,TR点注记,YH点注记,RL点注记,RS点注记,ZQ点注记,SY点注记,GS点注记"
    SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
    SSProcess.SelectFilter
    Count = SSProcess.GetSelNoteCount
    
    For j = 0 To Count - 1
        FormerVal = SSProcess.GetSelNoteValue(j,"SSObj_FontString")
        IDStr = SSProcess.GetSelNoteValue(j,"SSObj_ID")
        ws = Len(str) + 2
        qbwz = Len(FormerVal)
        rwz = qbwz - ws
        hmzte = Right(FormerVal,rwz)
        q2 = Left(FormerVal,2)
        fystr = q2 & hmzte
        'if j=0 then msgbox  fystr
        SSProcess.SetObjectAttr IDStr, "SSObj_FontString", fystr
    Next
End Function
'======================================================选择输出方式=========================================================

Function ConFirmWay(ByRef Way,ByRef res,ByRef GroupStr)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "选择输出方式","综合管线图",0,"综合管线图,分层输出",""
    'SSProcess.AddInputParameter "选择组号","",0,"",""
    res = SSProcess.ShowInputParameterDlg ("管线图输出方式")
    SSProcess.RefreshView
    If res = 1  Then
        Way = SSProcess.GetInputParameter("选择输出方式")
    End If
    GroupStr = ""
    'GroupStr = SSProcess.GetInputParameter("选择组号")
    If GroupStr <> "" Then
        SetPoiNote GroupStr
    End If
End Function' ConFirmWay

'设置注记名
Function SetPoiNote(ByVal GroupStr)
    LayArr = Split("CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS",",", - 1,1)
    For i = 0 To UBound(LayArr)
        SelNote LayArr(i) & "点注记",GroupStr
    Next 'i
End Function' SetPoiNote

'===================================================扩展属性修改========================================================

' [管线CAD输出]
' 附注=
' 图幅名称=
' 作业单位=苍南县测绘院
' 委托单位=
' 测量日期=2023年7月计算机成图
' 平面坐标体系=苍南城市坐标系
' 高程体系=1985国家高程基准，等高距0.5米。
' 图式=2017年版图式
' 探测员=张三
' 测量员=张三
' 绘图员=张三
' 检查员=张三

AttrStr = "所有权单位,委托单位,测量日期,平面坐标体系,高程体系,图式,探测员,测量员,绘图员,检查员"
KeyStr = "作业单位,委托单位,测量日期,平面坐标体系,高程体系,图式,探测员,测量员,绘图员,检查员"

Function ModifyAttr(ByVal Code,ByVal Way,ByRef TkId,ByRef XmMc,ByRef Count)
    SelFeatures Code,TkId,Count
    If Count = 0 Then Exit Function
    AttrArr = Split(AttrStr,",", - 1,1)
    KeyArr = Split(KeyStr,",", - 1,1)
    For i = 0 To UBound(AttrArr)
        SSProcess.SetObjectAttr TkId,"[" & AttrArr(i) & "]",SSProcess.ReadEpsIni("管线CAD输出", KeyArr(i) ,"")
    Next 'i
    XmMc = SSProcess.ReadEpsIni("管线报告信息", "项目名称" ,"")
    SqlStr = "Select XMMC From 管线项目信息表 Where 管线项目信息表.ID = 1"
    GetSQLRecordAll SqlStr,XmmcArr,Count
    If Count > 0 Then
        XmMc = XmmcArr(0)
    End If
    SSProcess.SetObjectAttr TkId,"[图幅名称]",XmMc
    If Way = "综合管线图" Then SSProcess.SetObjectAttr TkId,"[附注]",Way
    SSProcess.ObjectDeal TkId, "FreeDisplayList", Parameters, Result
    SSProcess.RefreshView
End Function' ModifyAttr

Function SetFcAttr(ByVal Code,ByRef TkId,ByRef XmMc,ByRef Count,ByVal BigName)
    SelFeatures Code,TkId,Count
    If Count = 0 Then Exit Function
    AttrArr = Split(AttrStr,",", - 1,1)
    KeyArr = Split(KeyStr,",", - 1,1)
    For i = 0 To UBound(AttrArr)
        SSProcess.SetObjectAttr TkId,"[" & AttrArr(i) & "]",SSProcess.ReadEpsIni("管线CAD输出", KeyArr(i) ,"")
    Next 'i
    SqlStr = "Select XMMC From 管线项目信息表 Where 管线项目信息表.ID = 1"
    GetSQLRecordAll SqlStr,XmmcArr,Count
    If Count > 0 Then
        XmMc = XmmcArr(0)
    End If
    SSProcess.SetObjectAttr TkId,"[图幅名称]",XmMc
    SSProcess.SetObjectAttr TkId,"[附注]",BigName & "地下管线图"
    SSProcess.ObjectDeal TkId, "FreeDisplayList", Parameters, Result
    SSProcess.RefreshView
End Function' SetFcAttr

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

'选择当前图廓并返回图廓ID
Function SelFeatures(ByVal Code,ByRef ID,ByRef Count)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
End Function' SelFeatures

'搜索所有的注记
Function SelNote(ByVal LayerName,ByVal GroupStr)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
    SSProcess.SelectFilter
    Count = SSProcess.GetSelNoteCount
    For j = 0 To Count - 1
        FormerVal = SSProcess.GetSelNoteValue(j,"SSObj_FontString")
        Prefix = Left(FormerVal,2)
        Suffix = Right(FormerVal,Len(FormerVal) - 2)
        CurrentVal = Prefix & GroupStr & Suffix
        SSProcess.SetSelNoteValue j,"SSObj_FontString",CurrentVal
    Next 'i
End Function' SelNote

'选择当前图层数据并返回个数
Function SelData(ByVal LayerName)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SelectFilter
    SelData = SSProcess.GetSelGeoCount
End Function' SelData

'获取当前图廓中所有的管线图层名称(大类)
Function GetAllLayerName(ByRef SmallArr(),ByRef LayArr())
    LayArr = Split("CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS",",", - 1,1)
    For j = 0 To UBound(LayArr)
        If SelData(LayArr(j)) > 0 Then
            If LayerStr = "" Then
                LayerStr = LayArr(j)
            Else
                LayerStr = LayerStr & "," & LayArr(j)
            End If
        End If
    Next 'j
    SmallArr = Split(LayerStr,",", - 1,1)
    Count = 0
    ReDim BigArr(Count)
    For i = 0 To UBound(SmallArr)
        If SmallArr(i) = "CD" Then
            BigArr(Count) = "CDD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "CT" Then
            BigArr(Count) = "CXD"
            Count = Count + 1
            ReDim  Preserve BigArr(Count)
        ElseIf SmallArr(i) = "CY" Or SmallArr(i) = "CQ" Or SmallArr(i) = "CS" Or SmallArr(i) = "QT" Then
            BigArr(Count) = "CYD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "BM" Or SmallArr(i) = "FQ" Then
            BigArr(Count) = "CSD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "DL" Or SmallArr(i) = "GD" Or SmallArr(i) = "LD" Or SmallArr(i) = "DC" Or SmallArr(i) = "XH" Then
            BigArr(Count) = "DLD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "TX" Or SmallArr(i) = "DX" Or SmallArr(i) = "YD" Or SmallArr(i) = "LT" Or SmallArr(i) = "JX" Or SmallArr(i) = "EX" Or SmallArr(i) = "DS" Or SmallArr(i) = "BZ" Then
            BigArr(Count) = "TXD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "JS" Or SmallArr(i) = "XF" Then
            BigArr(Count) = "JSD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "PS" Or SmallArr(i) = "YS" Or SmallArr(i) = "WS" Or SmallArr(i) = "FS" Then
            BigArr(Count) = "PSD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "RQ" Or SmallArr(i) = "MQ" Or SmallArr(i) = "TR" Or SmallArr(i) = "YH" Then
            BigArr(Count) = "RQD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "RL" Or SmallArr(i) = "RS" Or SmallArr(i) = "ZQ" Then
            BigArr(Count) = "RLD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        ElseIf SmallArr(i) = "SY" Or SmallArr(i) = "GS" Then
            BigArr(Count) = "GYD"
            Count = Count + 1
            ReDim Preserve BigArr(Count)
        End If
    Next 'i
    DelRepeat BigArr,LayerStr,LayerCount
    LayArr = Split(LayerStr,",", - 1,1)
    For i = 0 To UBound(LayArr)
        LayArr(i) = ToChinese(LayArr(i))
    Next 'i
End Function' GetAllLayerName

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

'==================================================CAD输出======================================================================

Function SZDWT(ByVal TkId,ByVal FilePath)
    SSProcess.SetFeatureCodeTB "FeatureCodeTB_kcad", "SymbolScriptTB_cad"
    SSProcess.SetNotetemplateTB "NoteTemplateTB_cad"
    SSProcess.ClearDataXParameter
    SSProcess.SetDataXParameter "DataType", "1"      '数据格式格式。0(ArcGIS SDE)、 1(DWG)、2(DXF)、 3(E00)、 4(Coverage)、 5(Shp)
    SSProcess.SetDataXParameter "Version", "2008"    'AutoCad数据版本号。2000,2004,2006
    SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_kcad"
    SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_cad"
    SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_cad"
    SSProcess.SetDataXParameter "ExportPathName", FilePath                    '输出文件名(或者路径名),如果为空时,则自动弹出对话框选择
    SSProcess.SetDataXParameter "DataBoundMode", "0"                    '数据输出范围方式， 0(所有数据)， 1(选择集数据)， 2(当前图幅)。
    SSProcess.SetDataXParameter "ZeroLineWidth", "10"
    SSProcess.SetDataXParameter "AcadColorMethod", "0"
    SSProcess.SetDataXParameter "ColorUseStatus", "1"       '颜色使用状态。0（按编码表设定颜色输出）、1（按地物设定颜色输出）
    SSProcess.SetDataXParameter "ExplodeObjColorStatus", "1"
    SSProcess.SetDataXParameter "FontWidthScale", "0.7"            '输出注记字宽缩放比
    SSProcess.SetDataXParameter "FontHeightScale", "0.7"        '输出注记字高缩放比  
    SSProcess.SetDataXParameter "FontSizeUseStatus","1"               '字体大小使用状态 0 （按注记分类表设置字高宽输出）、 1 （按注记设置字高宽输出）
    SSProcess.SetDataXParameter "OthersExportMode", "3"'输出AutoCAD数据时，厚度输出方式。 0（地物编码）、 1（编码表中的厚度）、 2（编码表中的别名）、3（置成0）
    SSProcess.SetDataXParameter "OthersExportToZFactor", "1"       '输出AutoCAD数据时，厚度输出到块Z比例方式。 0（不输出）、 1（输出）
    SSProcess.SetDataXParameter "ExplodeNoteStatus","0"
    SSProcess.SetDataXParameter "SymbolExplodeMode", "1"   '符号打散方式。 0（自动打散）、 1（根据编码表设定打散）、 2（全部不打散）
    SSProcess.SetDataXParameter "LayerUseStatus", "1"     '数据输出层名使用状态。0（按编码表设定层名输出）、1（按地物设定层名输出）
    SSProcess.SetDataXParameter "ExplodeObjLayerStatus", "0"  '内嵌符号图层输出方式。0（按符号描述设定输出）、 1（与主地物同层输出）
    SSProcess.SetDataXParameter "LineExportMode", "1" '输出AutoCAD数据时，多义线输出方式， 0 （缺省方式，带不同高程时按3DPolyline输出，其余按2DPolyline输出）、 1（强制按2DPolyline输出）、 2（ 强制按3DPolyline输出） 3（ 强制按Polyline输出）
    SSProcess.SetDataXParameter "LineWidthUseStatus", "0"
    SSProcess.SetDataXParameter "GotoPointsMode", "1"                     '输出图形折线化方式。 0 （不折线化）、 1 （只折线化曲线）、 2 （所有图形折线化）
    SSProcess.SetDataXParameter "AcadLineWidthMode", "3"
    SSProcess.SetDataXParameter "AcadLineScaleMode", "0"                'Acad线型比例输出方式。0 与比例尺成正比输出 1 总是按1输出
    SSProcess.SetDataXParameter "AcadLineWeightMode","0"               'Acad线重输出方式。0 地物线宽 1 随层 2 随块 3 随线定义
    SSProcess.SetDataXParameter "AcadBlockUseColorMode", "1"        'Acad图块输出颜色使用方式。0 随层 1 随块 2 随块内实体
    SSProcess.SetDataXParameter "AcadLinetypeGenerateMode", "1"
    SSProcess.SetDataXParameter "ExplodeObjMakeGroup ", "0"       'AutoCAD数据时，打散对象编组输出方式。 0（不编组）、 1（ 编组，同时要求FeatureCodeTB表中的ExtraInfo=1 
    SSProcess.SetDataXParameter "AcadUsePersonalBlockScaleCodes ", "1=7601023"       'AcadUsePersonalBlockScaleCodes 指定使用特殊块比例的编码。格式1： 比例1=编码1,编码2;比例2=编码1,编码2格式2： 比例 (该方式指定所有编码均使用指定的块比例) 
    SSProcess.SetDataXParameter "AcadDwtFileName", SSProcess.GetSysPathName (0) & "\Acadlin\acad.dwt"
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayerCount", "0"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DEFAULT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"叠加分析过渡面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"标注层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"三维测图"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"图廓层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"征地"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"乡镇属性点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"村属性点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"征地界址点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"境界"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地类图斑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"勘测图廓层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"测量控制点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"求算控制点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"控制点检查线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"数学基础"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"图幅范围面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"水系网线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"水系面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"水系线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"水系点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"水系附属设施线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"水系附属设施点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"水系附属设施面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"海洋线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"海洋面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"海洋点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"门址"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"楼址"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"POI"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"居民地点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"居民地附属设施线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"居民地面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"居民地线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"房屋外轮廓面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"房产辅助线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"居民地附属设施点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"居民地附属设施面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"部件"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"铁路线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"交通附属设施面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"交通附属设施点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"交通附属设施线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"道路面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"道路线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"道路网线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"交通附属设施"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"道路交叉口面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"管线线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"管线点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"管线面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"境界线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"境界点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"境界面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"其他境界面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"其他境界线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"其他境界点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"等高线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"高程点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地貌点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地貌面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地貌线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"植被与土质线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"植被与土质点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"植被与土质面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"勘测标注信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"三调"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"勘测村界图例"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"勘测图例"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"出图范围"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"征地注记"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地类图斑图例"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"征地村注记"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"征地界址线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"行政区"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地籍区"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地籍子区"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"理论控制点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"所有权宗地"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GPS检测点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"理论测站点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"实测测站点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"支点线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"使用权宗地"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"放验线用地范围"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"宗地界址点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"检查线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"宗地界址线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"方向线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"实测控制点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"宗海"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"宗海界址点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"宗海界址线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"理论放样点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"实测放样点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"正负零标高"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"宗地图廓层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"勘测图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"面积注记"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"构筑物"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"自然幢"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"点状定着物"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"线状定着物"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"面状定着物"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"楼层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"面积块"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"户"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"房间"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"外墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"内墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"分户中墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"分间中墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"门线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"窗户线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"房产部件点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"房产部件线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"房产部件面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"房产图廓层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"构件信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"规划建筑轴线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"偏差方向"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"规划控制线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"规划建筑物范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"建筑物范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"建筑物基底范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"建筑白膜"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"规划围墙线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"竣工标高信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"竣工标注信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"出图范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"规划测量成果图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"位移图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"建筑占地面积计算略图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"建筑高度及层高测量略图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"消防登高面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"机动车停车位"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"非机动车停车位"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"停车位分布图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"绿地范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"绿地竣工图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"问题信息"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"待更新区域"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"工作区域"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"更新区域"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"院落街区面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"TERP"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GTFA"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GTFL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"测量控制线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地方坐标原点"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地类图斑辅助线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"国家不一致图斑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"不一致图斑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"地方不一致图斑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"权属区域"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"征地项目面"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"宗地"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"放样建筑"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"建筑物外轮廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"立面图辅助线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"竣工平面图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"竣工对比图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"土地核验图图廓"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"停车位范围线"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"管线图例层"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"KZ_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"KZ_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"SX_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"SX_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"SX_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"SX_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JMD_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JMD_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JMD_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JMD_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JT_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JT_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JT_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JT_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GX_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GX_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GX_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JJ_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JJ_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JJ_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"JJ_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DM_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DM_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DM_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DM_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ZB_G"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ZB_L"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ZB_P"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ZB_A"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"QT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"TK"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"PSANNEXE"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"FSANNEXE"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"YSANNEXE"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"WSANNEXE"
    
    
    
    LayStr = "CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS"
    LayArr = Split(LayStr,",", - 1,1)
    For i = 0 To UBound(LayArr)
        SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),LayArr(i)
    Next 'i
    
    For i = 0 To UBound(LayArr)
        SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),LayArr(i) & "点注记"
        SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),LayArr(i) & "注记"
    Next 'i
    startIndex = 0
    SSProcess.SetDataXParameter "LayerRelationCount", "100"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CD:CDPOINT:CDLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CT:CTPOINT:CTLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CY:CYPOINT:CYLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQ:CQPOINT:CQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CS:CSPOINT:CSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QT:QTPOINT:QTLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BM:BMPOINT:BMLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQ:FQPOINT:FQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DL:DLPOINT:DLLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GD:GDPOINT:GDLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LD:LDPOINT:LDLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DC:DCPOINT:DCLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XH:XHPOINT:XHLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TX:TXPOINT:TXLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DX:DXPOINT:DXLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YD:DYPOINT:YDLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LT:LTPOINT:LTLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JX:JXPOINT:JXLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JK:JKPOINT:JKLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EX:EXPOINT:EXLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DS:DSPOINT:DSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZ:BZPOINT:BZLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JS:JSPOINT:JSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XF:XFPOINT:XFLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PS:PSPOINT:PSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YS:YSPOINT:YSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WS:WSPOINT:WSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FS:FSPOINT:FSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQ:RQPOINT:RQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQ:MQPOINT:MQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TR:TRPOINT:TRLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YH:YHPOINT:YHLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RL:RLPOINT:RLLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RS:RSPOINT:RSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQ:ZQPOINT:ZQLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SY:SYPOINT:SYLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GS:GSPOINT:GSLINE:::"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "管线图例层:TK:TK:TK:TK:TK"
    
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CD点注记::::CDTEXT:CDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CT点注记::::CTTEXT:CTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CY点注记::::CYTEXT:CYTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQ点注记::::CQTEXT:CQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CS点注记::::CSTEXT:CSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QT点注记::::QTTEXT:QTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BM点注记::::BMTEXT:BMTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQ点注记::::FQTEXT:FQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DL点注记::::DLTEXT:DLTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GD点注记::::GDTEXT:GDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LD点注记::::LDTEXT:LDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DC点注记::::DCTEXT:DCTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XH点注记::::XHTEXT:XHTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TX点注记::::TXTEXT:TXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DX点注记::::DXTEXT:DXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YD点注记::::YDTEXT:YDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LT点注记::::LTTEXT:LTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JX点注记::::JXTEXT:JXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JK点注记::::JKTEXT:JKTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EX点注记::::EXTEXT:EXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DS点注记::::DSTEXT:DSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZ点注记::::BZTEXT:BZTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JS点注记::::JSTEXT:JSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XF点注记::::XFTEXT:XFTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PS点注记::::PSTEXT:PSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YS点注记::::YSTEXT:YSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WS点注记::::WSTEXT:WSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FS点注记::::FSTEXT:FSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQ点注记::::RQTEXT:RQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQ点注记::::MQTEXT:MQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TR点注记::::TRTEXT:TRTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YH点注记::::YHTEXT:YHTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RL点注记::::RLTEXT:RLTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RS点注记::::RSTEXT:RSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQ点注记::::ZQTEXT:ZQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SY点注记::::SYTEXT:SYTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GS点注记::::GSTEXT:GSTEXT"
    
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CD注记::::CDMARK:CDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CT注记::::CTMARK:CTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CY注记::::CYMARK:CYMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQ注记::::CQMARK:CQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CS注记::::CSMARK:CSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QT注记::::QTMARK:QTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BM注记::::BMMARK:BMMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQ注记::::FQMARK:FQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DL注记::::DLMARK:DLMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GD注记::::GDMARK:GDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LD注记::::LDMARK:LDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DC注记::::DCMARK:DCMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XH注记::::XHMARK:XHMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TX注记::::TXMARK:TXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DX注记::::DXMARK:DXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YD注记::::YDMARK:YDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LT注记::::LTMARK:LTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JX注记::::JXMARK:JXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JK注记::::JKMARK:JKMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EX注记::::EXMARK:EXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DS注记::::DSMARK:DSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZ注记::::BZMARK:BZMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JS注记::::JSMARK:JSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XF注记::::XFMARK:XFMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PS注记::::PSMARK:PSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YS注记::::YSMARK:YSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WS注记::::WSMARK:WSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FS注记::::FSMARK:FSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQ注记::::RQMARK:RQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQ注记::::MQMARK:MQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TR注记::::TRMARK:TRMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YH注记::::YHMARK:YHMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RL注记::::RLMARK:RLMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RS注记::::RSMARK:RSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQ注记::::ZQMARK:ZQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SY注记::::SYMARK:SYMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GS注记::::GSMARK:GSMARK"
    startIndex = 0
    SSProcess.SetDataXParameter "TableFieldDefCount","3000"
    
    SSProcess.ExportData
    ' SSProcess.SetFeatureCodeTB "FeatureCodeTB_500", "SymbolScriptTB_500"
    ' SSProcess.SetNotetemplateTB "NoteTemplateTB_500"
    
End Function

'索引自动增加
Function AddOne(ByRef StartIndex)
    StartIndex = StartIndex + 1
    AddOne = StartIndex
End Function

'层名转化为中文
Function ToChinese(ByVal EngLayerName) 'EngLayerName 图层名称(英文)
    EngStr = "CDD,CXD,CYD,CSD,DLD,TXD,JSD,PSD,RQD,RLD,GYD"
    CheStr = "长输输电,长输通信,长输油气水,城市管线,电力,通信,给水,排水,燃气,热力,工业"
    EngArr = Split(EngStr,",", - 1,1)
    CheArr = Split(CheStr,",", - 1,1)
    ToChinese = ""
    For j = 0 To UBound(EngArr)
        If EngArr(j) = EngLayerName Then
            ToChinese = CheArr(j)
        End If
    Next 'j
End Function' ToChinese

'打开所有图层
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'关闭所有图层
Function AllDisVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 0, 1
    Next
    SSProcess.SetLayerStatus "图廓层", 1, 1
    SSProcess.RefreshView
End Function

Function Ending()
    MsgBox "输出完成"
End Function' Ending

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
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
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

'删除图廓
Function DelFormerTk()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "图廓层"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
End Function' DelFormerTk

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
