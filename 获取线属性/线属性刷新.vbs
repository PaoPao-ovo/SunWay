'支点线线上点
Dim ZDXArr(100000,3)

'检查线线上点
Dim JCXArr(100000,4)

'方向线线上点
Dim FXXArr(100000,4)

'控制点检查线线上点
Dim KZDJCXArr(100000,4)

'入口函数
Sub OnClick()
    SetZDX(9130221)
    SetJCX(9130241)
    SetFXX(9130251)
    SetKZDJCX(1130212)
    MsgBox "属性提取完成"
End Sub' OnClick

'设置支点线
Function SetZDX(Code)
    
    strGroupName = "绘线检查"
    strCheckName = "支点线检查"
    strPromptMessage = "请手动填写测站点号和支站点号"
    
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
    selcount = SSProcess.GetSelGeoCount
    For i = 0 To selcount - 1
        id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
        pointcount = transform(pointcount)
        For j = 0 To pointcount - 1
            SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name
            x = transform(x)
            y = transform(y)
            z = transform(z)
            ZDXArr(j,0) = x
            ZDXArr(j,1) = y
            ZDXArr(j,2) = z
        Next
        
        x1 = ZDXArr(0,0)
        y1 = ZDXArr(0,1)
        x2 = ZDXArr(1,0)
        y2 = ZDXArr(1,1)
        
        If y1 < y2 And x1 > x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 270 + SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 < y2 And x1 < x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 90 - SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 < y2 And x1 < x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 90 + SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 > y2 And x1 > x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 180 + SSProcess.RadianToDms(Atn(Abs(y / x)))
        End If
        angarr = Split(angles,".", - 1,1)
        If UBound(angarr) > 0 Then
            str = angarr(1)
            dd = ""
            ss = ""
            If Len(str) > 4 Then
                dd = Mid(str,1,2)
                ss = Mid(str,3,2)
            End If
            If Len(str) = 3 Then
                dd = Mid(str,1,2)
                ss = Mid(str,3,1) & "0"
            End If
            If Len(str) = 2 Then
                dd = Mid(str,1,2)
                ss = "00"
            End If
            If Len(str) = 1 Then
                dd = Mid(str,1,1) & "0"
                ss = "00"
            End If
            If Len(str) = 0 Then
                dd = "00"
                ss = "00"
            End If
        ElseIf UBound(angarr) = 0 Then
            dd = "00"
            ss = "00"
        End If
        longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
        longtitude = transform(longtitude)
        longtitude = FormatNumber(longtitude,3)
        SSProcess.SetObjectAttr id,"[ShuiPJL]",longtitude
        SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "°" & dd & "′" & ss & "″"
        
        idstring = SSProcess.SearchNearObjIDs(x1,y1,0.0001,0,"",0)
        idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
        IdCount = UBound(idarr) + 1
        'MsgBox IdCount
        If IdCount = 0 Then ExportInfo x1,y1,0,id,strGroupName,strCheckName,strPromptMessage
        If IdCount = 1 Then
            pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            'MsgBox id
            SSProcess.SetObjectAttr id,"[CeZDH]",pointname
        ElseIf IdCount = 2 Then
            Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
            If Firstname <> Secondname Then
                'MsgBox id
                ExportInfo x1,y1,0,id,strGroupName,strCheckName,strPromptMessage
                'Exit Function
            End If
            If Firstname = Secondname Then
                SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
            End If
        End If
        
        idstring = SSProcess.SearchNearObjIDs(x2,y2,0.0001,0,"",0)
        idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
        IdCount = UBound(idarr) + 1
        'MsgBox IdCount
        If IdCount = 0 Then ExportInfo x2,y2,0,id,strGroupName,strCheckName,strPromptMessage
        If IdCount = 1 Then
            pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            SSProcess.SetObjectAttr id,"[ZhiZDH]",pointname
        ElseIf IdCount = 2 Then
            Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
            If Firstname <> Secondname Then
                ExportInfo x2,y2,0,id,strGroupName,strCheckName,strPromptMessage
                'Exit Function
            End If
            If Firstname = Secondname Then
                SSProcess.SetObjectAttr id,"[ZhiZDH]",Firstname
            End If
        End If
    Next
End Function' SetZDX

'设置检查线
Function SetJCX(Code)
    Dim idnum(10000)
    strGroupName = "绘线检查"
    strCheckName = "检查线检查"
    strPromptMessage = "请手动填写测站点号和检查点号"
    
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
    selcount = SSProcess.GetSelGeoCount
    For i = 0 To selcount - 1
        id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        idnum(i) = id
    Next
    
    For k = 0 To selcount - 1
        pointcount = SSProcess.GetObjectAttr(idnum(k),"SSObj_PointCount")
        pointcount = transform(pointcount)
        For j = 0 To pointcount - 1
            SSProcess.GetObjectPoint idnum(k),j,x,y,z,pointtype,name
            x = transform(x)
            y = transform(y)
            z = transform(z)
            
            JCXArr(j,0) = x
            JCXArr(j,1) = y
            JCXArr(j,2) = z
            JCXArr(j,3) = name
            
        Next
        
        x1 = JCXArr(0,0)
        y1 = JCXArr(0,1)
        x2 = JCXArr(1,0)
        y2 = JCXArr(1,1)
        
        longtitude = SSProcess.GetObjectAttr(idnum(k),"SSObj_Length")
        longtitude = transform(longtitude)
        longtitude = FormatNumber(longtitude,3)
        If y1 < y2 And x1 > x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 270 + SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 < y2 And x1 < x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 90 - SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 > y2 And x1 < x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 90 + SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 > y2 And x1 >x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 180 + SSProcess.RadianToDms(Atn(Abs(y / x)))
        End If
        angarr = Split(angles,".", - 1,1)
        If UBound(angarr) > 0 Then
            str = angarr(1)
            dd = ""
            ss = ""
            If Len(str) > 4 Then
                dd = Mid(str,1,2)
                ss = Mid(str,3,2)
            End If
            If Len(str) = 3 Then
                dd = Mid(str,1,2)
                ss = Mid(str,3,1) & "0"
            End If
            If Len(str) = 2 Then
                dd = Mid(str,1,2)
                ss = "00"
            End If
            If Len(str) = 1 Then
                dd = Mid(str,1,1) & "0"
                ss = "00"
            End If
            If Len(str) = 0 Then
                dd = "00"
                ss = "00"
            End If
        ElseIf UBound(angarr) = 0 Then
            dd = "00"
            ss = "00"
        End If
        SSProcess.SetObjectAttr idnum(k),"[FangXZ]",angarr(0) & "°" & dd & "′" & ss & "″"
        SSProcess.SetObjectAttr idnum(k),"[ShuiPJL]",longtitude
        
        idstring = SSProcess.SearchNearObjIDs(x1,y1,0.0001,0,"",0)
        idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
        IdCount = UBound(idarr) + 1
        'MsgBox IdCount
        If IdCount = 0 Then ExportInfo x1,y1,0,idnum(k),strGroupName,strCheckName,strPromptMessage
        If IdCount = 1 Then
            pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            'MsgBox id
            SSProcess.SetObjectAttr idnum(k),"[CeZDH]",pointname
        ElseIf IdCount = 2 Then
            Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
            If Firstname <> Secondname Then
                'MsgBox id
                ExportInfo x1,y1,0,idnum(k),strGroupName,strCheckName,strPromptMessage
                'Exit Function
            End If
            If Firstname = Secondname Then
                SSProcess.SetObjectAttr idnum(k),"[CeZDH]",Firstname
            End If
        End If
        
        idstring = SSProcess.SearchNearObjIDs(x2,y2,0.0001,0,"9130311,9130312,9130217,9130512,9130412",0)
        idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
        IdCount = UBound(idarr) + 1
        'MsgBox IdCount
        
        If IdCount = 0 Then ExportInfo x2,y2,0,idnum(k),strGroupName,strCheckName,strPromptMessage
        If IdCount = 1 Then
            pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            SSProcess.SetObjectAttr idnum(k),"[JianCDH]",pointname
            code = SSProcess.GetObjectAttr(idarr(0),"SSObj_Code")
            If code = "9130217" Then
                DiffXY idnum(k),"9130216"
            ElseIf code = "9130311" Then
                DiffXY idnum(k),"9130211"
            ElseIf code = "9130312" Then
                DiffXY idnum(k),"9130212"
            ElseIf code = "9130512" Then
                DiffXY idnum(k),"1103021"
            ElseIf code = "9130412" Then
                DiffXY idnum(k),"1102021"
            End If
        ElseIf IdCount = 2 Then
            Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
            If Firstname <> Secondname Then
                ExportInfo x2,y2,0,idnum(k),strGroupName,strCheckName,strPromptMessage
                'Exit Function
            End If
            If Firstname = Secondname Then
                SSProcess.SetObjectAttr idnum(k),"[JianCDH]",Firstname
            End If
        End If
        
    Next
End Function' SetJCX

'设置方向线
Function SetFXX(Code)
    
    
    strGroupName = "绘线检查"
    strCheckName = "方向线检查"
    strPromptMessage = "请手动填写测站点号和方向点号"
    
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
    selcount = SSProcess.GetSelGeoCount
    For i = 0 To selcount - 1
        id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
        pointcount = transform(pointcount)
        For j = 0 To pointcount - 1
            SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name
            x = transform(x)
            y = transform(y)
            z = transform(z)
            
            FXXArr(j,0) = x
            FXXArr(j,1) = y
            FXXArr(j,2) = z
            FXXArr(j,3) = name
            
            
        Next
        x1 = FXXArr(0,0)
        y1 = FXXArr(0,1)
        x2 = FXXArr(1,0)
        y2 = FXXArr(1,1)
        
        longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
        longtitude = transform(longtitude)
        longtitude = FormatNumber(longtitude,3)
        SSProcess.SetObjectAttr id,"[ShuiPJL]",longtitude
        If y1 < y2 And x1 > x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 270 + SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 < y2 And x1 < x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 90 - SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 > y2 And x1 < x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 90 + SSProcess.RadianToDms(Atn(Abs(x / y)))
        End If
        If y1 > y2 And x1 > x2 Then
            x = y2 - y1
            y = x2 - x1
            angles = 180 + SSProcess.RadianToDms(Atn(Abs(y / x)))
        End If
        angarr = Split(angles,".", - 1,1)
        If UBound(angarr) > 0 Then
            str = angarr(1)
            dd = ""
            ss = ""
            If Len(str) > 4 Then
                dd = Mid(str,1,2)
                ss = Mid(str,3,2)
            End If
            If Len(str) = 3 Then
                dd = Mid(str,1,2)
                ss = Mid(str,3,1) & "0"
            End If
            If Len(str) = 2 Then
                dd = Mid(str,1,2)
                ss = "00"
            End If
            If Len(str) = 1 Then
                dd = Mid(str,1,1) & "0"
                ss = "00"
            End If
            If Len(str) = 0 Then
                dd = "00"
                ss = "00"
            End If
        ElseIf UBound(angarr) = 0 Then
            dd = "00"
            ss = "00"
        End If
        SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "°" & dd & "′" & ss & "″"
        
        idstring = SSProcess.SearchNearObjIDs(x1,y1,0.0001,0,"",0)
        idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
        IdCount = UBound(idarr) + 1
        'MsgBox IdCount
        If IdCount = 0 Then ExportInfo x1,y1,0,id,strGroupName,strCheckName,strPromptMessage
        If IdCount = 1 Then
            pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            'MsgBox id
            SSProcess.SetObjectAttr id,"[CeZDH]",pointname
        ElseIf IdCount = 2 Then
            Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
            If Firstname <> Secondname Then
                'MsgBox id
                ExportInfo x1,y1,0,id,strGroupName,strCheckName,strPromptMessage
                'Exit Function
            End If
            If Firstname = Secondname Then
                SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
            End If
        End If
        
        idstring = SSProcess.SearchNearObjIDs(x2,y2,0.0001,0,"",0)
        idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
        IdCount = UBound(idarr) + 1
        'MsgBox IdCount
        If IdCount = 0 Then ExportInfo x2,y2,0,id,strGroupName,strCheckName,strPromptMessage
        If IdCount = 1 Then
            pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            SSProcess.SetObjectAttr id,"[FangXDH]",pointname
        ElseIf IdCount = 2 Then
            Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
            If Firstname <> Secondname Then
                ExportInfo x2,y2,0,id,strGroupName,strCheckName,strPromptMessage
                'Exit Function
            End If
            If Firstname = Secondname Then
                SSProcess.SetObjectAttr id,"[FangXDH]",Firstname
            End If
        End If
    Next
End Function' SetFXX

'设置控制点检查线
Function SetKZDJCX(Code)
    Dim idnum(10000)
    strGroupName = "绘线检查"
    strCheckName = "控制点检查线线检查"
    strPromptMessage = "请手动填写测站点号和检查点号"
    
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SelectFilter
    selcount = SSProcess.GetSelGeoCount
    'MsgBox selcount
    For i = 0 To selcount - 1
        id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        'MsgBox id
        idnum(i) = id
    Next
    
    For k = 0 To selcount - 1
        pointcount = SSProcess.GetObjectAttr(idnum(k),"SSObj_PointCount")
        pointcount = transform(pointcount)
        
        For j = 0 To pointcount - 1
            SSProcess.GetObjectPoint idnum(k),j,x,y,z,pointtype,name
            x = transform(x)
            y = transform(y)
            z = transform(z)
            
            KZDJCXArr(j,0) = x
            KZDJCXArr(j,1) = y
            KZDJCXArr(j,2) = z
            KZDJCXArr(j,3) = name
            
            longtitude = SSProcess.GetObjectAttr(idnum(k),"SSObj_Length")
            longtitude = transform(longtitude)
            longtitude = FormatNumber(longtitude,3)
            SSProcess.SetObjectAttr idnum(k),"[JCBC]",longtitude
            
        Next
        x1 = KZDJCXArr(0,0)
        y1 = KZDJCXArr(0,1)
        x2 = KZDJCXArr(1,0)
        y2 = KZDJCXArr(1,1)
        
        idstring = SSProcess.SearchNearObjIDs(x1,y1,0.0001,0,"",0)
        idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
        IdCount = UBound(idarr) + 1
        'MsgBox IdCount
        If IdCount = 0 Then ExportInfo x1,y1,0,idnum(k),strGroupName,strCheckName,strPromptMessage
        If IdCount = 1 Then
            pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            'MsgBox id
            SSProcess.SetObjectAttr idnum(k),"[CeZDH]",pointname
        ElseIf IdCount = 2 Then
            Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
            If Firstname <> Secondname Then
                'MsgBox id
                ExportInfo x1,y1,0,idnum(k),strGroupName,strCheckName,strPromptMessage
                'Exit Function
            End If
            If Firstname = Secondname Then
                SSProcess.SetObjectAttr idnum(k),"[CeZDH]",Firstname
            End If
        End If
        
        idstring = SSProcess.SearchNearObjIDs(x2,y2,0.0001,0,"",0)
        idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
        IdCount = UBound(idarr) + 1
        'MsgBox IdCount
        If IdCount = 0 Then ExportInfo x2,y2,0,idnum(k),strGroupName,strCheckName,strPromptMessage
        If IdCount = 1 Then
            pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            SSProcess.SetObjectAttr idnum(k),"[JianCDH]",pointname
        ElseIf IdCount = 2 Then
            Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
            Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
            If Firstname <> Secondname Then
                ExportInfo x2,y2,0,idnum(k),strGroupName,strCheckName,strPromptMessage
                'Exit Function
            End If
            If Firstname = Secondname Then
                SSProcess.SetObjectAttr idnum(k),"[JianCDH]",Firstname
            End If
        End If
        SetYZBC(idnum(k))
        comparelong(idnum(k))
    Next
End Function' SetKZDJCX

'设置X,Y差值
Function DiffXY(id,code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", code
    SSProcess.SetSelectCondition "SSObj_PointName", "==",JCXArr(1,3)
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    'MsgBox PointArr(1,3)
    If SelCount > 0 Then
        X = SSProcess.GetSelGeoValue(0, "SSObj_X")
        X = transform(X)
        Y = SSProcess.GetSelGeoValue(0, "SSObj_Y")
        Y = transform(Y)
        diffx = Abs(X - JCXArr(1,0))
        diffy = Abs(Y - JCXArr(1,1))
        diffx = FormatNumber(diffx,3)
        diffy = FormatNumber(diffy,3)
        SSProcess.SetObjectAttr id,"[XZuoBCZ]",diffy
        SSProcess.SetObjectAttr id,"[YZuoBCZ]",diffx
    Else
        'MsgBox "不存在同名点"
        Exit Function
    End If
End Function' DiffXY

'计算边长较差
Function comparelong(id)
    yzbc = SSProcess.GetObjectAttr(id,"[YZBC]")
    jcbc = SSProcess.GetObjectAttr(id,"[JCBC]")
    yzbc = transform(yzbc)
    jcbc = transform(jcbc)
    bcjc = Abs(yzbc - jcbc)
    SSProcess.SetObjectAttr id,"[BCJC]",bcjc
End Function' comparelong

'设置已知边长
Function SetYZBC(id)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130211"
    SSProcess.SetSelectCondition "SSObj_PointName", "==",KZDJCXArr(1,3)
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    If SelCount > 0 Then
        X = SSProcess.GetSelGeoValue(0, "SSObj_X")
        X = transform(X)
        Y = SSProcess.GetSelGeoValue(0, "SSObj_Y")
        Y = transform(Y)
        yzbc = Sqr((KZDJCXArr(0,0) - X) ^ 2 + (KZDJCXArr(0,1) - Y) ^ 2)
        yzbc = FormatNumber(yzbc,3)
        SSProcess.SetObjectAttr id,"[YZBC]",yzbc
    End If
End Function' SetYZBC

'数据类型转换
Function transform(content)
    If content <> "" Then
        content = CDbl(content)
    Else
        MsgBox "数据有误"
    End If
    transform = content
End Function

Function ExportInfo(x,y,z,id,strGroupName,strCheckName,strPromptMessage)
    SSProcess.AddCheckRecord strGroupName, strCheckName, "自定义脚本检查类->" & strCheckName, strPromptMessage, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function' ExportInfo