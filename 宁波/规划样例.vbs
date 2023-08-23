
Sub OnClick()

    theID = SSProcess.GetSelGeoValue (0,"SSObj_ID")
    pnt = SSProcess.GetSelGeoValue(0,"SSObj_PointCount")
    zg = 200
    zk = 160

    XG = 2
    bilichi = SSProcess.GetMapScale
    xs = 6 * zg / 200
    xs = xs * bilichi / 500
    If IsNumeric(xs) = False Then
        xs = 1
    Else
        If xs = 0 Then xs = 1
    End If
    
    xiaoshu = 2
    If IsNumeric(XG) = False Then XG = 2
    If XG = 2 Then
        XG = True
    Else
        XG = False
    End If
    SSProcess.GetSelGeoPoint 0, 0, x0,  y0,  z0,  ptype0,  name
    SSProcess.GetSelGeoPoint 0,1, x1,  y1,  z1,  ptype1,  name
    SSProcess.GetSelGeoPoint 0,pnt - 1, x2,  y2,  z2,  ptype2,  name
    hchang = Sqr(CDbl(x0 - x1) * CDbl(x0 - x1) + CDbl(y0 - y1) * CDbl(y0 - y1))
    SSProcess.XYSA x0,y0,x1,y1,dist,angle,0
    SSProcess.RadianToDeg radian
    xx = CDbl((x0 + x1) / 2)
    yy = CDbl((y0 + y1) / 2)
    angle00 = CDbl(angle - (3.141592654 / 2))
    tt = 0
    SSProcess.XYSA xx,yy,x,y,tt,angle00,1
    angle = SSProcess.RadianToDeg(angle)

    NR = FormatNumber(CDbl(hchang),xiaoshu, - 1,0,0)
    
    If pnt = 2 Then
        xxz = (x0 + x1) / 2
        yyz = ((y0 + y1 + 1.4) / 2)
    ElseIf pnt = 3 Then
        xxz = x2
        yyz = y2
    End If
    addnote  NR,RGB(255,255,255),angle,"0",zg,zk,pnt,xxz,yyz,z,"5121"
    zjangle = angle
    
    'SSProcess.SetObjectAttr theID, "[TJBC]", NR
    'SSProcess.SetSelGeoValue 0 ,"[JL]", NR  
    gd = Abs(CDbl(x0 - x2))
    angle = angle + 90
    angle = SSProcess.DegToRadian(angle)
    juli = SSProcess.DistPerpend(resx, resy, pResRelation, x2, y2, x1, y1, x0, y0 )
    If XG <> True Then
        juli1 = juli + 0.18 * xs
    Else
        juli1 = juli
    End If
    SSProcess.XYSA resx,resy,x2,y2,juli1,angle,0
    SSProcess.XYSA x0,y0,xx0,yy0,juli1,angle,1
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code","931009201"
    SSProcess.SetNewObjValue "SSObj_Color",RGB(255,255,255)
    'SSProcess.AddNewObjPoint x0, y0,z, 0, ""
    'SSProcess.AddNewObjPoint xx0, yy0,z, 0, ""
    SSProcess.AddNewObjToSelObjList
    SSProcess.XYSA x1,y1,xx1,yy1,juli1,angle,1
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code","1"
    SSProcess.SetNewObjValue "SSObj_Color",RGB(255,255,255)
    'SSProcess.AddNewObjPoint x1, y1,z, 0, ""
    'SSProcess.AddNewObjPoint xx1, yy1,z, 0, ""
    SSProcess.AddNewObjToSelObjList
    SSProcess.XYSA x0,y0,xx0,yy0,juli,angle,1
    SSProcess.XYSA x1,y1,xx1,yy1,juli,angle,1
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code","931009201"
    SSProcess.SetNewObjValue "SSObj_Color",RGB(255,255,255)
    SSProcess.AddNewObjPoint CDbl(xx0), CDbl(yy0),z, 0, ""
    SSProcess.AddNewObjPoint CDbl(xx1), CDbl(yy1),z, 0, ""
    SSProcess.AddNewObjToSelObjList
    
    
End Sub

Function addnote(zjnr0,color,angle0,FontAlignment0,FontWidth0,FontHeight0,pnt0,xx0,yy0,z0,Status)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "FX001"
    SSProcess.SetNewObjValue "SSObj_Color",color
    SSProcess.SetNewObjValue "SSObj_FontString",zjnr0
    If angle0 > 90 And angle0 < 270  Then
        angle0 = angle0 + 180
        If angle0 > 360 Then angle0 = angle0 - 360
    End If
    If angle0 < 90 Then xx0 = xx0 - 0.7
    SSProcess.SetNewObjValue "SSObj_FontAlignment",FontAlignment0
    SSProcess.SetNewObjValue "SSObj_FontStringAngle",angle0
    SSProcess.SetNewObjValue "SSObj_FontWordAngle",angle0
    If Status <> "" Then
        SSProcess.SetNewObjValue "SSObj_Status",Status
    End If
    SSProcess.SetNewObjValue "SSObj_FontWidth",FontHeight0
    SSProcess.SetNewObjValue "SSObj_FontHeight",FontWidth0
    SSProcess.AddNewObjPoint CDbl(xx0), CDbl(yy0),z0, 0, ""
    SSProcess.AddNewObjToSelObjList
End Function
