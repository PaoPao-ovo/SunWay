
MainCode = "54311203,54324004,54323004,54412004,54423004,54452004,54511004,54512114,54534114,54523114,54611114,54612004,54623004,54111003,54112003,54123003,54145003,54134003,54211003,54212003,54223003,54234003,54245003,54256003,54267003,54278003,54289003,54720114,54730114,54030003,54040003,51011203,52011203,53011204,53022204,53033204,53044204"

'=======================================================�������=========================================================

'�����
Sub OnClick()
    ConFirmWay Way,res,GroupStr
    If res = 1 Then
        If Way = "�ۺϹ���ͼ" Then
            AllVisible
            ModifyAttr "59999999",Way,TkId,XmMc,Count
            If Count = 0 Then
                MsgBox "������ͼ��"
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
            FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & "�ۺϹ���ͼ.dwg"
            DelFormerTl
            SZDWT TkId,FilePath
            AllVisible
            
        ElseIf Way = "�ֲ����" Then
            GetAllLayerName SmallArr,BigArr
            For k = 0 To UBound(BigArr)
                Select Case BigArr(k)
                    Case "�������"
                    AllDisVisible
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    SSProcess.SetLayerStatus "CDע��", 1, 1
                    SSProcess.SetLayerStatus "CD��ע��", 1, 1
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
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
                    FilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "����ͨ��"
                    AllDisVisible
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    SSProcess.SetLayerStatus "CTע��", 1, 1
                    
                    SSProcess.SetLayerStatus "CT��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "��������ˮ"
                    AllDisVisible
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    SSProcess.SetLayerStatus "CYע��", 1, 1
                    SSProcess.SetLayerStatus "CQע��", 1, 1
                    SSProcess.SetLayerStatus "CSע��", 1, 1
                    SSProcess.SetLayerStatus "QTע��", 1, 1
                    
                    SSProcess.SetLayerStatus "CY��ע��", 1, 1
                    SSProcess.SetLayerStatus "CQ��ע��", 1, 1
                    SSProcess.SetLayerStatus "CS��ע��", 1, 1
                    SSProcess.SetLayerStatus "QT��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    
                    Case "���й���"
                    AllDisVisible
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    SSProcess.SetLayerStatus "BMע��", 1, 1
                    SSProcess.SetLayerStatus "FQע��", 1, 1
                    
                    SSProcess.SetLayerStatus "BM��ע��", 1, 1
                    SSProcess.SetLayerStatus "FQ��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "����"
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
                    SSProcess.SetLayerStatus "XHע��", 1, 1
                    SSProcess.SetLayerStatus "DCע��", 1, 1
                    SSProcess.SetLayerStatus "LDע��", 1, 1
                    SSProcess.SetLayerStatus "GDע��", 1, 1
                    SSProcess.SetLayerStatus "DLע��", 1, 1
                    
                    SSProcess.SetLayerStatus "XH��ע��", 1, 1
                    SSProcess.SetLayerStatus "DC��ע��", 1, 1
                    SSProcess.SetLayerStatus "LD��ע��", 1, 1
                    SSProcess.SetLayerStatus "GD��ע��", 1, 1
                    SSProcess.SetLayerStatus "DL���ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "ͨ��"
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
                    SSProcess.SetLayerStatus "BZע��", 1, 1
                    SSProcess.SetLayerStatus "DXע��", 1, 1
                    SSProcess.SetLayerStatus "YDע��", 1, 1
                    SSProcess.SetLayerStatus "LTע��", 1, 1
                    SSProcess.SetLayerStatus "JXע��", 1, 1
                    SSProcess.SetLayerStatus "JKע��", 1, 1
                    SSProcess.SetLayerStatus "EXע��", 1, 1
                    SSProcess.SetLayerStatus "DSע��", 1, 1
                    SSProcess.SetLayerStatus "TXע��", 1, 1
                    
                    SSProcess.SetLayerStatus "BZ��ע��", 1, 1
                    SSProcess.SetLayerStatus "DX��ע��", 1, 1
                    SSProcess.SetLayerStatus "YD��ע��", 1, 1
                    SSProcess.SetLayerStatus "LT��ע��", 1, 1
                    SSProcess.SetLayerStatus "JX��ע��", 1, 1
                    SSProcess.SetLayerStatus "JK��ע��", 1, 1
                    SSProcess.SetLayerStatus "EX��ע��", 1, 1
                    SSProcess.SetLayerStatus "DS��ע��", 1, 1
                    SSProcess.SetLayerStatus "TX��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "��ˮ"
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
                    SSProcess.SetLayerStatus "JSע��", 1, 1
                    SSProcess.SetLayerStatus "XFע��", 1, 1
                    
                    SSProcess.SetLayerStatus "JS��ע��", 1, 1
                    SSProcess.SetLayerStatus "XF��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "��ˮ"
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
                    SSProcess.SetLayerStatus "FSע��", 1, 1
                    SSProcess.SetLayerStatus "WSע��", 1, 1
                    SSProcess.SetLayerStatus "YSע��", 1, 1
                    SSProcess.SetLayerStatus "PSע��", 1, 1
                    
                    SSProcess.SetLayerStatus "FS��ע��", 1, 1
                    SSProcess.SetLayerStatus "WS��ע��", 1, 1
                    SSProcess.SetLayerStatus "YS��ע��", 1, 1
                    SSProcess.SetLayerStatus "PS��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "ȼ��"
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
                    SSProcess.SetLayerStatus "YHע��", 1, 1
                    SSProcess.SetLayerStatus "MQע��", 1, 1
                    SSProcess.SetLayerStatus "TRע��", 1, 1
                    SSProcess.SetLayerStatus "RQע��", 1, 1
                    
                    SSProcess.SetLayerStatus "YH��ע��", 1, 1
                    SSProcess.SetLayerStatus "MQ��ע��", 1, 1
                    SSProcess.SetLayerStatus "TR��ע��", 1, 1
                    SSProcess.SetLayerStatus "RQ��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "����"
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
                    SSProcess.SetLayerStatus "ZQע��", 1, 1
                    SSProcess.SetLayerStatus "RLע��", 1, 1
                    SSProcess.SetLayerStatus "RSע��", 1, 1
                    
                    SSProcess.SetLayerStatus "ZQ��ע��", 1, 1
                    SSProcess.SetLayerStatus "RL��ע��", 1, 1
                    SSProcess.SetLayerStatus "RS��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                    Case "��ҵ"
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
                    SSProcess.SetLayerStatus "GSע��", 1, 1
                    SSProcess.SetLayerStatus "SYע��", 1, 1
                    
                    SSProcess.SetLayerStatus "GS��ע��", 1, 1
                    SSProcess.SetLayerStatus "SY��ע��", 1, 1
                    
                    SSProcess.SetLayerStatus "����ͼ����", 1, 1
                    SSProcess.SetLayerStatus "ͼ����", 1, 1
                    DelFormerTl
                    SetFcAttr "59999999",TkId,XmMc,Count,BigArr(k)
                    If Count = 0 Then
                        MsgBox "������ͼ��"
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
                    FFilePath = SSProcess.GetSysPathName(5) & "רҵ����ͼ\" & XmMc & BigArr(k) & "���¹���ͼ.dwg"
                    DelFormerTl
                    SZDWT TkId,FilePath
                End Select
            Next 'z
        End If
        FYNOTE GroupStr
        Ending
    Else
        MsgBox "���˳�"
    End If
    AllVisible
    
End Sub' OnClick

Function FYNOTE(STR)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "CD��ע��,CT��ע��,CY��ע��,CQ��ע��,CS��ע��,QT��ע��,BM��ע��,FQ��ע��,DL��ע��,GD��ע��,LD��ע��,DC��ע��,XH��ע��,TX��ע��,DX��ע��,YD��ע��,LT��ע��,JX��ע��,JK��ע��,EX��ע��,DS��ע��,BZ��ע��,JS��ע��,XF��ע��,PS��ע��,YS��ע��,WS��ע��,FS��ע��,RQ��ע��,MQ��ע��,TR��ע��,YH��ע��,RL��ע��,RS��ע��,ZQ��ע��,SY��ע��,GS��ע��"
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
'======================================================ѡ�������ʽ=========================================================

Function ConFirmWay(ByRef Way,ByRef res,ByRef GroupStr)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "ѡ�������ʽ","�ۺϹ���ͼ",0,"�ۺϹ���ͼ,�ֲ����",""
    'SSProcess.AddInputParameter "ѡ�����","",0,"",""
    res = SSProcess.ShowInputParameterDlg ("����ͼ�����ʽ")
    SSProcess.RefreshView
    If res = 1  Then
        Way = SSProcess.GetInputParameter("ѡ�������ʽ")
    End If
    GroupStr = ""
    'GroupStr = SSProcess.GetInputParameter("ѡ�����")
    If GroupStr <> "" Then
        SetPoiNote GroupStr
    End If
End Function' ConFirmWay

'����ע����
Function SetPoiNote(ByVal GroupStr)
    LayArr = Split("CD,CT,CY,CQ,CS,QT,BM,FQ,DL,GD,LD,DC,XH,TX,DX,YD,LT,JX,JK,EX,DS,BZ,JS,XF,PS,YS,WS,FS,RQ,MQ,TR,YH,RL,RS,ZQ,SY,GS",",", - 1,1)
    For i = 0 To UBound(LayArr)
        SelNote LayArr(i) & "��ע��",GroupStr
    Next 'i
End Function' SetPoiNote

'===================================================��չ�����޸�========================================================

' [����CAD���]
' ��ע=
' ͼ������=
' ��ҵ��λ=�����ز��Ժ
' ί�е�λ=
' ��������=2023��7�¼������ͼ
' ƽ��������ϵ=���ϳ�������ϵ
' �߳���ϵ=1985���Ҹ̻߳�׼���ȸ߾�0.5�ס�
' ͼʽ=2017���ͼʽ
' ̽��Ա=����
' ����Ա=����
' ��ͼԱ=����
' ���Ա=����

AttrStr = "����Ȩ��λ,ί�е�λ,��������,ƽ��������ϵ,�߳���ϵ,ͼʽ,̽��Ա,����Ա,��ͼԱ,���Ա"
KeyStr = "��ҵ��λ,ί�е�λ,��������,ƽ��������ϵ,�߳���ϵ,ͼʽ,̽��Ա,����Ա,��ͼԱ,���Ա"

Function ModifyAttr(ByVal Code,ByVal Way,ByRef TkId,ByRef XmMc,ByRef Count)
    SelFeatures Code,TkId,Count
    If Count = 0 Then Exit Function
    AttrArr = Split(AttrStr,",", - 1,1)
    KeyArr = Split(KeyStr,",", - 1,1)
    For i = 0 To UBound(AttrArr)
        SSProcess.SetObjectAttr TkId,"[" & AttrArr(i) & "]",SSProcess.ReadEpsIni("����CAD���", KeyArr(i) ,"")
    Next 'i
    XmMc = SSProcess.ReadEpsIni("���߱�����Ϣ", "��Ŀ����" ,"")
    SqlStr = "Select XMMC From ������Ŀ��Ϣ�� Where ������Ŀ��Ϣ��.ID = 1"
    GetSQLRecordAll SqlStr,XmmcArr,Count
    If Count > 0 Then
        XmMc = XmmcArr(0)
    End If
    SSProcess.SetObjectAttr TkId,"[ͼ������]",XmMc
    If Way = "�ۺϹ���ͼ" Then SSProcess.SetObjectAttr TkId,"[��ע]",Way
    SSProcess.ObjectDeal TkId, "FreeDisplayList", Parameters, Result
    SSProcess.RefreshView
End Function' ModifyAttr

Function SetFcAttr(ByVal Code,ByRef TkId,ByRef XmMc,ByRef Count,ByVal BigName)
    SelFeatures Code,TkId,Count
    If Count = 0 Then Exit Function
    AttrArr = Split(AttrStr,",", - 1,1)
    KeyArr = Split(KeyStr,",", - 1,1)
    For i = 0 To UBound(AttrArr)
        SSProcess.SetObjectAttr TkId,"[" & AttrArr(i) & "]",SSProcess.ReadEpsIni("����CAD���", KeyArr(i) ,"")
    Next 'i
    SqlStr = "Select XMMC From ������Ŀ��Ϣ�� Where ������Ŀ��Ϣ��.ID = 1"
    GetSQLRecordAll SqlStr,XmmcArr,Count
    If Count > 0 Then
        XmMc = XmmcArr(0)
    End If
    SSProcess.SetObjectAttr TkId,"[ͼ������]",XmMc
    SSProcess.SetObjectAttr TkId,"[��ע]",BigName & "���¹���ͼ"
    SSProcess.ObjectDeal TkId, "FreeDisplayList", Parameters, Result
    SSProcess.RefreshView
End Function' SetFcAttr

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

'ѡ��ǰͼ��������ͼ��ID
Function SelFeatures(ByVal Code,ByRef ID,ByRef Count)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
End Function' SelFeatures

'�������е�ע��
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

'ѡ��ǰͼ�����ݲ����ظ���
Function SelData(ByVal LayerName)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", LayerName
    SSProcess.SelectFilter
    SelData = SSProcess.GetSelGeoCount
End Function' SelData

'��ȡ��ǰͼ�������еĹ���ͼ������(����)
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

'ȥ���ַ������ظ�ֵ
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

'==================================================CAD���======================================================================

Function SZDWT(ByVal TkId,ByVal FilePath)
    SSProcess.SetFeatureCodeTB "FeatureCodeTB_kcad", "SymbolScriptTB_cad"
    SSProcess.SetNotetemplateTB "NoteTemplateTB_cad"
    SSProcess.ClearDataXParameter
    SSProcess.SetDataXParameter "DataType", "1"      '���ݸ�ʽ��ʽ��0(ArcGIS SDE)�� 1(DWG)��2(DXF)�� 3(E00)�� 4(Coverage)�� 5(Shp)
    SSProcess.SetDataXParameter "Version", "2008"    'AutoCad���ݰ汾�š�2000,2004,2006
    SSProcess.SetDataXParameter "FeatureCodeTBName", "FeatureCodeTB_kcad"
    SSProcess.SetDataXParameter "SymbolScriptTBName", "SymbolScriptTB_cad"
    SSProcess.SetDataXParameter "NoteTemplateTBName", "NoteTemplateTB_cad"
    SSProcess.SetDataXParameter "ExportPathName", FilePath                    '����ļ���(����·����),���Ϊ��ʱ,���Զ������Ի���ѡ��
    SSProcess.SetDataXParameter "DataBoundMode", "0"                    '���������Χ��ʽ�� 0(��������)�� 1(ѡ������)�� 2(��ǰͼ��)��
    SSProcess.SetDataXParameter "ZeroLineWidth", "10"
    SSProcess.SetDataXParameter "AcadColorMethod", "0"
    SSProcess.SetDataXParameter "ColorUseStatus", "1"       '��ɫʹ��״̬��0����������趨��ɫ�������1���������趨��ɫ�����
    SSProcess.SetDataXParameter "ExplodeObjColorStatus", "1"
    SSProcess.SetDataXParameter "FontWidthScale", "0.7"            '���ע���ֿ����ű�
    SSProcess.SetDataXParameter "FontHeightScale", "0.7"        '���ע���ָ����ű�  
    SSProcess.SetDataXParameter "FontSizeUseStatus","1"               '�����Сʹ��״̬ 0 ����ע�Ƿ���������ָ߿�������� 1 ����ע�������ָ߿������
    SSProcess.SetDataXParameter "OthersExportMode", "3"'���AutoCAD����ʱ����������ʽ�� 0��������룩�� 1��������еĺ�ȣ��� 2��������еı�������3���ó�0��
    SSProcess.SetDataXParameter "OthersExportToZFactor", "1"       '���AutoCAD����ʱ������������Z������ʽ�� 0����������� 1�������
    SSProcess.SetDataXParameter "ExplodeNoteStatus","0"
    SSProcess.SetDataXParameter "SymbolExplodeMode", "1"   '���Ŵ�ɢ��ʽ�� 0���Զ���ɢ���� 1�����ݱ�����趨��ɢ���� 2��ȫ������ɢ��
    SSProcess.SetDataXParameter "LayerUseStatus", "1"     '�����������ʹ��״̬��0����������趨�����������1���������趨���������
    SSProcess.SetDataXParameter "ExplodeObjLayerStatus", "0"  '��Ƕ����ͼ�������ʽ��0�������������趨������� 1����������ͬ�������
    SSProcess.SetDataXParameter "LineExportMode", "1" '���AutoCAD����ʱ�������������ʽ�� 0 ��ȱʡ��ʽ������ͬ�߳�ʱ��3DPolyline��������ఴ2DPolyline������� 1��ǿ�ư�2DPolyline������� 2�� ǿ�ư�3DPolyline����� 3�� ǿ�ư�Polyline�����
    SSProcess.SetDataXParameter "LineWidthUseStatus", "0"
    SSProcess.SetDataXParameter "GotoPointsMode", "1"                     '���ͼ�����߻���ʽ�� 0 �������߻����� 1 ��ֻ���߻����ߣ��� 2 ������ͼ�����߻���
    SSProcess.SetDataXParameter "AcadLineWidthMode", "3"
    SSProcess.SetDataXParameter "AcadLineScaleMode", "0"                'Acad���ͱ��������ʽ��0 ������߳�������� 1 ���ǰ�1���
    SSProcess.SetDataXParameter "AcadLineWeightMode","0"               'Acad���������ʽ��0 �����߿� 1 ��� 2 ��� 3 ���߶���
    SSProcess.SetDataXParameter "AcadBlockUseColorMode", "1"        'Acadͼ�������ɫʹ�÷�ʽ��0 ��� 1 ��� 2 �����ʵ��
    SSProcess.SetDataXParameter "AcadLinetypeGenerateMode", "1"
    SSProcess.SetDataXParameter "ExplodeObjMakeGroup ", "0"       'AutoCAD����ʱ����ɢ������������ʽ�� 0�������飩�� 1�� ���飬ͬʱҪ��FeatureCodeTB���е�ExtraInfo=1 
    SSProcess.SetDataXParameter "AcadUsePersonalBlockScaleCodes ", "1=7601023"       'AcadUsePersonalBlockScaleCodes ָ��ʹ�����������ı��롣��ʽ1�� ����1=����1,����2;����2=����1,����2��ʽ2�� ���� (�÷�ʽָ�����б����ʹ��ָ���Ŀ����) 
    SSProcess.SetDataXParameter "AcadDwtFileName", SSProcess.GetSysPathName (0) & "\Acadlin\acad.dwt"
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayerCount", "0"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"DEFAULT"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ӷ���������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ά��ͼ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������Ե�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����Ե�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���Ƶ�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ѧ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͼ����Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ˮϵ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ַ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"¥ַ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"POI"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ص�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ظ�����ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͨ������ʩ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͨ������ʩ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��·�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ߵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ȸ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�̵߳�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ò��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ֲ����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ֲ�������ʵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ֲ����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����ע��Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͼ��Χ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ش�ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ؼ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ؼ�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ۿ��Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����Ȩ�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GPS����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���۲�վ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ʵ���վ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"֧����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ʹ��Ȩ�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�������õط�Χ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ڵؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ڵؽ�ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ʵ����Ƶ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ں�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ں���ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ں���ַ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���۷�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ʵ�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ڵ�ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���ע��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��Ȼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��״������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"¥��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ֻ���ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ּ���ǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ƫ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮�����ﷶΧ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����ﷶΧ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������׷�Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ĥ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮Χǽ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������ע��Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��ͼ��Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�滮�����ɹ�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"λ��ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ռ�����������ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����߶ȼ���߲�����ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����Ǹ���"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������ͣ��λ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ǻ�����ͣ��λ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͣ��λ�ֲ�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�̵ط�Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�̵ؿ���ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ϣ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"Ժ�������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"TERP"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GTFA"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"GTFL"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ط�����ԭ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ�߸�����"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���Ҳ�һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ط���һ��ͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"Ȩ������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������Ŀ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�ڵ�"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"��������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"������������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ������"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ƽ��ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"�����Ա�ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"���غ���ͼͼ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"ͣ��λ��Χ��"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),"����ͼ����"
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
        SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),LayArr(i) & "��ע��"
        SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(StartIndex)),LayArr(i) & "ע��"
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
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "����ͼ����:TK:TK:TK:TK:TK"
    
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CD��ע��::::CDTEXT:CDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CT��ע��::::CTTEXT:CTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CY��ע��::::CYTEXT:CYTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQ��ע��::::CQTEXT:CQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CS��ע��::::CSTEXT:CSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QT��ע��::::QTTEXT:QTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BM��ע��::::BMTEXT:BMTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQ��ע��::::FQTEXT:FQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DL��ע��::::DLTEXT:DLTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GD��ע��::::GDTEXT:GDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LD��ע��::::LDTEXT:LDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DC��ע��::::DCTEXT:DCTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XH��ע��::::XHTEXT:XHTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TX��ע��::::TXTEXT:TXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DX��ע��::::DXTEXT:DXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YD��ע��::::YDTEXT:YDTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LT��ע��::::LTTEXT:LTTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JX��ע��::::JXTEXT:JXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JK��ע��::::JKTEXT:JKTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EX��ע��::::EXTEXT:EXTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DS��ע��::::DSTEXT:DSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZ��ע��::::BZTEXT:BZTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JS��ע��::::JSTEXT:JSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XF��ע��::::XFTEXT:XFTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PS��ע��::::PSTEXT:PSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YS��ע��::::YSTEXT:YSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WS��ע��::::WSTEXT:WSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FS��ע��::::FSTEXT:FSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQ��ע��::::RQTEXT:RQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQ��ע��::::MQTEXT:MQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TR��ע��::::TRTEXT:TRTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YH��ע��::::YHTEXT:YHTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RL��ע��::::RLTEXT:RLTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RS��ע��::::RSTEXT:RSTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQ��ע��::::ZQTEXT:ZQTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SY��ע��::::SYTEXT:SYTEXT"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GS��ע��::::GSTEXT:GSTEXT"
    
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CDע��::::CDMARK:CDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CTע��::::CTMARK:CTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CYע��::::CYMARK:CYMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CQע��::::CQMARK:CQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "CSע��::::CSMARK:CSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "QTע��::::QTMARK:QTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BMע��::::BMMARK:BMMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FQע��::::FQMARK:FQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DLע��::::DLMARK:DLMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GDע��::::GDMARK:GDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LDע��::::LDMARK:LDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DCע��::::DCMARK:DCMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XHע��::::XHMARK:XHMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TXע��::::TXMARK:TXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DXע��::::DXMARK:DXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YDע��::::YDMARK:YDMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "LTע��::::LTMARK:LTMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JXע��::::JXMARK:JXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JKע��::::JKMARK:JKMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "EXע��::::EXMARK:EXMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "DSע��::::DSMARK:DSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "BZע��::::BZMARK:BZMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "JSע��::::JSMARK:JSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "XFע��::::XFMARK:XFMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "PSע��::::PSMARK:PSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YSע��::::YSMARK:YSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "WSע��::::WSMARK:WSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "FSע��::::FSMARK:FSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RQע��::::RQMARK:RQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "MQע��::::MQMARK:MQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "TRע��::::TRMARK:TRMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "YHע��::::YHMARK:YHMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RLע��::::RLMARK:RLMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "RSע��::::RSMARK:RSMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "ZQע��::::ZQMARK:ZQMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "SYע��::::SYMARK:SYMARK"
    SSProcess.SetDataXParameter "LayerRelation" & CStr(AddOne(startIndex)), "GSע��::::GSMARK:GSMARK"
    startIndex = 0
    SSProcess.SetDataXParameter "TableFieldDefCount","3000"
    
    SSProcess.ExportData
    ' SSProcess.SetFeatureCodeTB "FeatureCodeTB_500", "SymbolScriptTB_500"
    ' SSProcess.SetNotetemplateTB "NoteTemplateTB_500"
    
End Function

'�����Զ�����
Function AddOne(ByRef StartIndex)
    StartIndex = StartIndex + 1
    AddOne = StartIndex
End Function

'����ת��Ϊ����
Function ToChinese(ByVal EngLayerName) 'EngLayerName ͼ������(Ӣ��)
    EngStr = "CDD,CXD,CYD,CSD,DLD,TXD,JSD,PSD,RQD,RLD,GYD"
    CheStr = "�������,����ͨ��,��������ˮ,���й���,����,ͨ��,��ˮ,��ˮ,ȼ��,����,��ҵ"
    EngArr = Split(EngStr,",", - 1,1)
    CheArr = Split(CheStr,",", - 1,1)
    ToChinese = ""
    For j = 0 To UBound(EngArr)
        If EngArr(j) = EngLayerName Then
            ToChinese = CheArr(j)
        End If
    Next 'j
End Function' ToChinese

'������ͼ��
Function AllVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'�ر�����ͼ��
Function AllDisVisible()
    count = SSProcess.GetLayerCount
    For i = 0 To count - 1
        layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus layername, 0, 1
    Next
    SSProcess.SetLayerStatus "ͼ����", 1, 1
    SSProcess.RefreshView
End Function

Function Ending()
    MsgBox "������"
End Function' Ending

'��ȡͼ�������Ͻ�����ֵ
Function GetMapBorderPoision(ByRef X,ByRef Y)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "ͼ����"
    SSProcess.SetSelectCondition "SSObj_Code", "==", "59999999"
    SSProcess.SelectFilter
    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
    SSProcess.GetObjectPoint ID, 2, X, Y, Z, PointType, Name '���Ͻǵ�����ֵ
End Function' GetMapBorderPoision

'���Ƶ�ע��
Function DrawPoint(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawPointNote X + 7,Y,Code,Width,Height
End Function

'���Ƶ�ע����
Function DrawPointNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GropuId
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'������ע��
Function DrawLine(ByVal X1,ByVal X2,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.AddNewObjPoint X1, Y, 0, 0, ""
    SSProcess.AddNewObjPoint X2, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    DrawLineNote X2 + 1,Y,Code,Width,Height
End Function

'������ע����
Function DrawLineNote(ByVal X,ByVal Y,ByVal Code,ByVal Width,ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", SSProcess.GetFeatureCodeInfo(Code,"ObjectName")
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    SSProcess.SetNewObjValue "SSObj_Color", SSProcess.GetFeatureCodeInfo(Code,"LineColor")
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X + 4, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'������߻���
Function DrawBorder(ByVal StartX,ByVal EndX,ByVal StartY,ByVal EndY)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", "51111111"
    'SSProcess.SetNewObjValue "SSObj_GroupID", GroupId
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.AddNewObjPoint StartX,StartY,0,0,""
    SSProcess.AddNewObjPoint EndX, StartY,0,0,""
    SSProcess.AddNewObjPoint EndX,EndY,0, 0,""
    SSProcess.AddNewObjPoint StartX,EndY,0,0,""
    SSProcess.AddNewObjPoint StartX,StartY,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
    '���Ʊ���
    DrawTitle (StartX + EndX) / 2,StartY - 1,250,250
    
End Function

'���Ʊ���
Function DrawTitle(ByVal X,ByVal Y,ByVal Width, ByVal Height)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", "ͼ ��"
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_LayerName", "����ͼ����"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    'SSProcess.SetNewObjValue "SSObj_GroupID", GroupId
    SSProcess.AddNewObjPoint X, Y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'ɾ��ͼ��
Function DelFormerTl()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "����ͼ����"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
End Function' DelFormerTl

'ɾ��ͼ��
Function DelFormerTk()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "ͼ����"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj()
End Function' DelFormerTk

'��ȡ���е���Ҫ������
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
        If SSProcess.GetSelGeoValue(i,"[FSFS]") = "�ܿ�" Or SSProcess.GetSelGeoValue(i,"[FSFS]") = "�ǿ���" Then
            CodeStr(k) = SSProcess.GetSelGeoValue(i,"SSObj_Code")
            k = k + 1
            ReDim Preserve CodeStr(k)
        ElseIf SSProcess.GetSelGeoValue(i,"[FSFS]") <> "��������" And SSProcess.GetSelGeoValue(i,"[YYKS]") <> "0" Then
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

'��ȡ���е���Ҫ������
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

'ȥ���ַ������ظ�ֵ(����)
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
