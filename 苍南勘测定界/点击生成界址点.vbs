
Sub OnInitScript()
    mode = 0 '=0 无参数对话框 =1 有参数对话框
    title = "点击生成界址点"
    SSProcess.ShowScriptDlg mode,title
    '自定义对话框
    'SSProcess.ShowScriptUserDefDlg title, dlgTemplateName, dlgWidth, dlgHeight, colCount, titleWidth, valueWidth
    '添加代码
End Sub

Sub OnExitScript()
    '添加代码
End Sub

Sub OnOK()
    '添加代码
End Sub

Sub OnCancel()
    '添加代码
End Sub

Function OnLButtonDown(x, y, spx, spy, flags)
    MapScale = SSProcess.GetMapScale
    xs = MapScale / 1000
    Dim jzdid(100)
    ids = SSProcess.SearchNearObjIDs(spx,spy, 0.005, 0, "1234", 0 )
    If ids = "" Then
        MsgBox "附近无界址点"
    Else
        SSFunc.ScanString ids,",",jzdid,JZDSL
        If    JZDSL <> 1 Then
            MsgBox "附近有多个界址点，点击的位置离，所选界址点更近一点"
        Else
            jzdh = SSProcess.GetObjectAttr(JZDID(0), "[jzdh]")
            dkh = SSProcess.GetObjectAttr(JZDID(0), "[dkh]")
            strJZXSM = "J" & jzdh
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_Code", "==", 504
            SSProcess.SetSelectCondition "[dkh]", "==", dkh
            SSProcess.SelectFilter
            Count = SSProcess.GetSelGeoCount()
            DKID = SSProcess.GetSelGeoValue( 0, "SSObj_ID" )
            'jzdzs=SSProcess.GetSelGeoValue( 0, "SSObj_PointCount" )-1
            'msgbox jzdh
            'msgbox typename(jzdzs)
            'if jzdh=jzdzs then 
            
            'else 
            'jzdzs=cstr(jzdzs)
            jzdh = Int(jzdh)
            Dim strs_qb(500)
            minjzdh = jzdh
            maxjzdh = jzdh
            'msgbox typename (maxjzdh)
            qbids = SSProcess.SearchInPolyObjIDs(DKID, 0, "1234", 0,1,0)
            SSFunc.ScanString qbids,",",strs_qb,scount_qb
            
            'msgbox scount_qb
            For m = 0 To scount_qb - 1
                czjzdh = Int(SSProcess.GetObjectAttr (strs_qb(m),"[jzdh]"))
                'msgbox    czjzdh
                If czjzdh > maxjzdh Then  maxjzdh = czjzdh
                If czjzdh < minjzdh Then  minjzdh = czjzdh
                
            Next
            'msgbox "最大"& maxjzdh
            'msgbox  "最小"&minjzdh
            Select Case jzdh
                Case minjzdh
                getid maxjzdh,dkh,qydid
                getid  minjzdh + 1,dkh,hydid
                
                SSProcess.GetObjectPoint JZDID(0), 0, x0,  y0,  z0, pt0,  pn0
                SSProcess.GetObjectPoint qydid, 0, x1,  y1,  z1, pt1,  pn1
                SSProcess.GetObjectPoint hydid, 0, x2,  y2,  z2, pt2,  pn2
                Case maxjzdh
                'msgbox "1"
                getid maxjzdh - 1,dkh,qydid
                getid  minjzdh,dkh,hydid
                SSProcess.GetObjectPoint JZDID(0), 0, x0,  y0,  z0, pt0,  pn0
                SSProcess.GetObjectPoint qydid, 0, x1,  y1,  z1, pt1,  pn1
                SSProcess.GetObjectPoint hydid, 0, x2,  y2,  z2, pt2,  pn2
                Case Else
                getid jzdh - 1,dkh,qydid
                getid  jzdh + 1,dkh,hydid
                'msgbox qydid
                SSProcess.GetObjectPoint JZDID(0), 0, x0,  y0,  z0, pt0,  pn0
                SSProcess.GetObjectPoint qydid, 0, x1,  y1,  z1, pt1,  pn1
                SSProcess.GetObjectPoint hydid, 0, x2,  y2,  z2, pt2,  pn2
            End Select
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
                makenote  x4,y4, "9135035",200 * xs,200 * xs,strJZXSM
                znjzd = SSProcess.SearchInnerObjIDs(DKID,3, "9135035", 0)
                If znjzd <> ""Then
                    'MSGBOX ZNJZD
                    SSProcess.DeleteObject znjzd
                Else
                    'MSGBOX "1"
                    Exit For
                End If
            Next
        End If
    End If
End Function

Function getid(jzdh,dkh,id)
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 1234
    SSProcess.SetSelectCondition "[dkh]", "==", dkh
    SSProcess.SetSelectCondition "[jzdh]", "==", jzdh
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount()
    ID = SSProcess.GetSelGeoValue( 0, "SSObj_ID" )
    
End Function

Function makeNote(x, y, code, width, height, fontString)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", code
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color",  "RGB(255,0,0)"
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    'SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    'SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    'SSProcess.SetNewObjValue "SSObj_FontStringAngle", angle0
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function