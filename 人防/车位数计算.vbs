
'��λ����
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
    
    NoteStr = "���³�λ�ϼ�" & CarCount & "���������˷������ڳ�λ" & RFCount & "����"
    TipStr = "ע��������΢�ͳ�λ" & TotalWXCount & "������0.7���㡣"
    
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

'================================================ҵ����======================================================================

'���������
Function JDC(ByVal Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "==", Code
    SSProcess.SelectFilter
    TotalCount = SSProcess.GetSelGeoCount()
    JDC = TotalCount
End Function' JDC

'����΢�ͳ�
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

'����ǻ�����
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
                Case "����"
                BLXS = 1.8
                TotalCount = TotalCount + Int(Transform(SSProcess.GetSelGeoValue(i,"[CheWMJ]")) / BLXS)
                Case "¶��"
                BLXS = 1.5
                TotalCount = TotalCount + Int(Transform(SSProcess.GetSelGeoValue(i,"[CheWMJ]")) / BLXS)
                Case  "·��"
                BLXS = 1.2
                TotalCount = TotalCount + Int(Transform(SSProcess.GetSelGeoValue(i,"[CheWMJ]")) / BLXS)
            End Select
        Next 'i
    End If
    FJDC = TotalCount
End Function' FJDC

'���������Ԫ�ڳ�λ��
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

'����ע��
Function DrawNote(ByVal x,ByVal y, ByVal fontString)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_FontName", "����"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", 4
    SSProcess.SetNewObjValue "SSObj_Color", RGB(255, 0, 0)
    SSProcess.SetNewObjValue "SSObj_FontWidth", 437
    SSProcess.SetNewObjValue "SSObj_FontHeight", 437
    SSProcess.SetNewObjValue "SSObj_FontWeight", 400
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function' DrawNote


'=================================================�����ຯ��==================================================================

'��������ת��
Function Transform(ByVal Content)
    If Content <> "" Then
        Content = CDbl(Content)
    Else
        Content = 0
    End If
    Transform = Content
End Function' Transform
