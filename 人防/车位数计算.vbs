
'车位编码
Dim CodeStr
CodeStr = "9461053,9461013,9461023,9461033,9461043"

Sub OnClick()
    
    CodeArr = Split(CodeStr,",", - 1,1)
    
    CarCount = 0
    
    For i = 0 To UBound(CodeArr)
        j = CodeArr(i)
        'MsgBox TypeName(j) 
        Select Case j
            Case "9461013"
            CarCount = CarCount + JDC(j)
            Case "9461033"
            CarCount = CarCount + JDC(j)
            Case "9461053"
            CarCount = CarCount + WXC(j)
            Case "9461023"
            CarCount = CarCount + FJDC(j)
            Case "9461043"
            CarCount = CarCount + FJDC(j)
        End Select
    Next 'i
    
    RFCount = FH_FJDC()
    TotalWXCount = WXC("9461053")
    
    NoteStr = "地下车位合计" & CarCount & "个，其中人防区域内车位" & RFCount & "个。"
    TipStr = "注：地下室微型车位" & TotalWXCount & "个，按0.7折算。"
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", "9464923"
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    If SelCount > 0 Then
        ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
        SSProcess.GetObjectPoint ID, 2, x, y, z, pointtype, name
        DrawNote x - 3,y - 3,NoteStr
        DrawNote x - 3,y - 9,TipStr
    End If
End Sub' OnClick

'================================================业务函数======================================================================

'计算机动车
Function JDC(ByVal Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", Code
    SSProcess.SelectFilter
    TotalCount = SSProcess.GetSelGeoCount()
    JDC = TotalCount
End Function' JDC

'计算微型车
Function WXC(ByVal Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", Code
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    BLXS = 0.7
    TotalCount = 0
    If SelCount > 0 Then
        For i = 0 To SelCount - 1
            TotalCount = TotalCount + Int(Transform(SSProcess.GetSelGeoValue(i,"[CheWMJ]")) / BLXS)
        Next 'i
    End If
    WXC = TotalCount
End Function' WxC

'计算非机动车
Function FJDC(ByVal Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", Code
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    TotalCount = 0
    If SelCount > 0 Then
        For i = 0 To SelCount - 1
            CWLB = SSProcess.GetSelGeoValue(i,"[CheWLB]")
            Select Case CWLB
                Case "室内"
                BLXS = 1.8
                TotalCount = TotalCount + Int(Transform(SSProcess.GetSelGeoValue(i,"[CheWMJ]")) / BLXS)
                Case "露天"
                BLXS = 1.5
                TotalCount = TotalCount + Int(Transform(SSProcess.GetSelGeoValue(i,"[CheWMJ]")) / BLXS)
                Case  "路边"
                BLXS = 1.2
                TotalCount = TotalCount + Int(Transform(SSProcess.GetSelGeoValue(i,"[CheWMJ]")) / BLXS)
            End Select
        Next 'i
    End If
    FJDC = TotalCount
End Function' FJDC

'计算防护单元内车位数
Function FH_FJDC()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", "9530226"
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    TotalCount = 0
    If SelCount > 0 Then
        For i = 0 To SelCount - 1
            TotalCount = TotalCount + Transform((SSProcess.GetSelGeoValue(i,"[FJDCS]"))) + Transform((SSProcess.GetSelGeoValue(i,"[TCWS]")))
        Next 'i
    End If
    FH_FJDC = TotalCount
End Function' FH_FJDC

'设置注释
Function DrawNote(ByVal x,ByVal y, ByVal fontString)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_FontName", "黑体"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", 4
    SSProcess.SetNewObjValue "SSObj_Color", RGB(255, 0, 0)
    SSProcess.SetNewObjValue "SSObj_FontWidth", 437
    SSProcess.SetNewObjValue "SSObj_FontHeight", 437
    SSProcess.SetNewObjValue "SSObj_FontWeight", 400
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function' DrawNote


'=================================================工具类函数==================================================================

'数据类型转换
Function Transform(ByVal Content)
    If Content <> "" Then
        Content = CDbl(Content)
    Else
        Content = 0
    End If
    Transform = Content
End Function' Transform
