
Sub OnClick()
    Const JZDBM = "1234"
    Const QSMBM = "504"
    Dim jzdx(50000)
    Dim jzdy(50000)
    Dim jzdcd(50000)
    Dim cdjh(50000)
    jzdcd(0) = 0
    jzdh = 1
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_FontClass", "==", "9135035"
    SSProcess.SelectFilter
    geoecount = SSProcess.GetSelNoteCount
    For i = 0 To geoecount - 1
        SSProcess.DelSelNote i
    Next
    SetZRZingNoteOffset
    AddNote2
    DelInner
    'MsgBox  "界址点标注完成"
End Sub

Function SetZRZingNoteOffset
    MapScale = SSProcess.GetMapScale
    xs = MapScale / 1000
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    Dim JZDIDSZ(50000)
    If geoCount > 0 Then
        For J = 0 To geoCount - 1
            DKID = SSProcess.GetSelGeoValue( J, "SSObj_ID" )
            JZDID = SSProcess.SearchInnerObjIDs(DKID,0, "1234", 0)
            SSFunc.ScanString JZDID,",",JZDIDSZ,JZDSL
            For m = 0 To JZDSL - 1
                SSProcess.GetObjectPoint JZDIDSZ(m), 0, x0,  y0,  z0, pt0,  pn0
                Select Case m
                    Case 0
                    SSProcess.GetObjectPoint JZDIDSZ(m), 0, x0,  y0,  z0, pt0,  pn0
                    SSProcess.GetObjectPoint JZDIDSZ(JZDSL - 1), 0, x1,  y1,  z1, pt1,  pn1
                    SSProcess.GetObjectPoint JZDIDSZ(m + 1), 0, x2,  y2,  z2, pt2,  pn2
                    Case JZDSL - 1
                    SSProcess.GetObjectPoint JZDIDSZ(m), 0, x0,  y0,  z0, pt0,  pn0
                    SSProcess.GetObjectPoint JZDIDSZ(m - 1), 0, x1,  y1,  z1, pt1,  pn1
                    SSProcess.GetObjectPoint JZDIDSZ(0), 0, x2,  y2,  z2, pt2,  pn2
                    Case Else
                    SSProcess.GetObjectPoint JZDIDSZ(m), 0, x0,  y0,  z0, pt0,  pn0
                    SSProcess.GetObjectPoint JZDIDSZ(m - 1), 0, x1,  y1,  z1, pt1,  pn1
                    SSProcess.GetObjectPoint JZDIDSZ(m + 1), 0, x2,  y2,  z2, pt2,  pn2
                End Select
                jzdh = SSProcess.GetObjectAttr(JZDIDSZ(m), "[jzdh]")
                strJZXSM = "J" & jzdh
                flag = 0
                SSProcess.XYSA x1,y1,x0,y0,dist1,angle1,flag
                SSProcess.XYSA x2,y2,x0,y0,dist2,angle2,flag
                If angle1 > angle2 Then
                    minang = angle2
                Else
                    minang = angle1
                End If
                
                anglepj = (Abs(angle2 - angle1)) * 0.5 + minang
                PI = 3.1415926535898
                If xs <> 1 Then
                    dist1 = 200 * xs * 0.01 * xs + 3'实际距离
                Else
                    dist1 = 2.5
                End If
                For n = 1 To 2
                    If n = 1  Then
                        angle = anglepj
                    Else
                        angle = anglepj + pi
                    End If
                    SSProcess.XYSA x0,y0,x4,y4,dist1,angle,1
                    makenote  x4,y4, "9135035",350 * xs,350 * xs,strJZXSM'图上距离
                Next
            Next
        Next
    End If
End Function

'圈号大于1的界址点的注记
Function AddNote2()
    MapScale = SSProcess.GetMapScale
    xs = MapScale / 1000
    GetAllDKH DKHArr
    For i = 0 To UBound(DKHArr)
        GetMaxQH DKHArr(i),MaxQH
        For j = 0 To MaxQH - 2
            SelJZPoi j + 2,DKHArr(i),JZDArr
            PoiCount = UBound(JZDArr) + 1
            For m = 0 To PoiCount - 1
                SSProcess.GetObjectPoint JZDArr(m), 0, x0,  y0,  z0, pt0,  pn0
                Select Case m
                    Case 0
                    SSProcess.GetObjectPoint JZDArr(m), 0, x0,  y0,  z0, pt0,  pn0
                    SSProcess.GetObjectPoint JZDArr(PoiCount - 1), 0, x1,  y1,  z1, pt1,  pn1
                    SSProcess.GetObjectPoint JZDArr(m + 1), 0, x2,  y2,  z2, pt2,  pn2
                    Case PoiCount - 1
                    SSProcess.GetObjectPoint JZDArr(m), 0, x0,  y0,  z0, pt0,  pn0
                    SSProcess.GetObjectPoint JZDArr(m - 1), 0, x1,  y1,  z1, pt1,  pn1
                    SSProcess.GetObjectPoint JZDArr(0), 0, x2,  y2,  z2, pt2,  pn2
                    Case Else
                    SSProcess.GetObjectPoint JZDArr(m), 0, x0,  y0,  z0, pt0,  pn0
                    SSProcess.GetObjectPoint JZDArr(m - 1), 0, x1,  y1,  z1, pt1,  pn1
                    SSProcess.GetObjectPoint JZDArr(m + 1), 0, x2,  y2,  z2, pt2,  pn2
                End Select
                jzdh = SSProcess.GetObjectAttr(JZDArr(m), "[jzdh]")
                strJZXSM = "J" & jzdh
                flag = 0
                SSProcess.XYSA x1,y1,x0,y0,dist1,angle1,flag
                SSProcess.XYSA x2,y2,x0,y0,dist2,angle2,flag
                If angle1 > angle2 Then
                    minang = angle2
                Else
                    minang = angle1
                End If
                
                anglepj = (Abs(angle2 - angle1)) * 0.5 + minang
                PI = 3.1415926535898
                If xs <> 1 Then
                    dist1 = 200 * xs * 0.01 * xs + 3'实际距离
                Else
                    dist1 = 2.5
                End If
                
                For n = 1 To 2
                    If n = 1  Then
                        angle = anglepj
                    Else
                        angle = anglepj + pi
                    End If
                    SSProcess.XYSA x0,y0,x4,y4,dist1,angle,1
                    makenote  x4,y4, "9135035",350 * xs,350 * xs,strJZXSM'图上距离
                Next
            Next 'm
        Next 'j
    Next 'i
End Function' AddNote2

'获取所有的地块号
Function GetAllDKH(ByRef DKHArr())
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    DKCount = SSProcess.GetSelGeoCount()
    If DKCount > 0 Then
        ReDim DKHArr(DKCount - 1)
        For i = 0 To DKCount - 1
            DKHArr(i) = SSProcess.GetSelGeoValue(i,"[DKH]")
        Next 'i
    End If
End Function' GetAllDKH

'选择对应的界址点
Function SelJZPoi(ByVal QH,ByVal DKH,ByRef JZDArr())
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 1234
    SSProcess.SetSelectCondition "[DKH]", "==", DKH
    SSProcess.SetSelectCondition "[QH]", "==", QH
    SSProcess.SelectFilter
    PoiCount = SSProcess.GetSelGeoCount()
    ReDim JZDArr(PoiCount - 1)
    For i = 0 To PoiCount - 1
        JZDArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
    Next 'i
End Function' SelJZPoi

'选择对应地块的圈号大于1的界址点
Function GetMaxQH(ByVal DKH,ByRef MaxQH)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 1234
    SSProcess.SetSelectCondition "[DKH]", "==", DKH
    SSProcess.SetSelectCondition "[QH]", ">", 1
    SSProcess.SelectFilter
    PoiCount = SSProcess.GetSelGeoCount()
    MaxQH = 1
    For i = 0 To PoiCount - 1
        CurrentQH = Transform(SSProcess.GetSelGeoValue(i,"[QH]"))
        If  CurrentQH > MaxQH Then
            MaxQH = CurrentQH
        End If
    Next 'i
End Function' GetMaxQH

'删除内部要素
Function DelInner()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        DKID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        znjzd = SSProcess.SearchInnerObjIDs(DKID,3,"9135035",0)
        If znjzd <> ""Then
            'MSGBOX ZNJZD
            arr1 = Split(znjzd,",", - 1,1)
            For j = 0 To UBound(arr1)
                SSProcess.DeleteObject arr1(j)
            Next 'j
        Else
            Exit For
        End If
    Next 'i
End Function' DelInner

Function makeNote(x, y, code, width, Height, fontString)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", code
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color",  "RGB(255,0,0)"
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'数据类型转换
Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform