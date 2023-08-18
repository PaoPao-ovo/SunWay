'线上点的数组
Dim PointArr(2,4)
'检查集组名
Dim strGroupName:strGroupName = "重复点检查"
'检查集检查名
Dim strCheckName:strCheckName = "控制点检查线线检查"
'检查日志
Dim strPromptMessage:strPromptMessage = "请手动填写测站点号和检查点号"

'入口函数
Function kzdjcx(id)
    GetOnlinePoint(id)
    SearchNear(id)
End Function ' kzdjcx

'传值函数
Function SearchNear(id)
    x1 = PointArr(0,0)
    y1 = PointArr(0,1)
    x2 = PointArr(1,0)
    y2 = PointArr(1,1)
    SetLinepoiname x1,y1,x2,y2,id
    SetProp x1,y1,x2,y2,id
End Function ' SearchNear

'获取线上的空间点信息
Function GetOnlinePoint(id)  
    Dim x, y, z, pointtype, name
            pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
            'MsgBox pointcount
            pointcount = transform(pointcount)
            For j = 0 To pointcount -1
                SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name 
                x = transform(x)
                y = transform(y)
                z = transform(z)
                PointArr(j,0) = x
                PointArr(j,1) = y
                PointArr(j,2) = z
                PointArr(j,3) = name
            Next
    'MsgBox PointArr(1,0)
End Function ' GetOnlinePoint

'设置线的方向值和水平距离
Function SetProp(x1,y1,x2,y2,id)
    longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
    longtitude = transform(longtitude)
    longtitude = formatnumber(longtitude,3)
    SSProcess.SetObjectAttr id,"[JCBC]",longtitude
    'SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "°" & dd & "′" & ss & "″"
End Function ' SetProp

'设置已知边长
Function SetYZBC(id)
    SSProcess.ClearSelection 
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130211"
    SSProcess.SetSelectCondition "SSObj_PointName", "==",PointArr(1,3)
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    If SelCount > 0 Then
        X = SSProcess.GetSelGeoValue(0, "SSObj_X")
        X = transform(X)
        Y = SSProcess.GetSelGeoValue(0, "SSObj_Y")
        Y = transform(Y) 
        yzbc = Sqr((PointArr(0,0) - X)^2 + (PointArr(0,1) - Y)^2)
        yzbc = FormatNumber(yzbc,3)
        SSProcess.SetObjectAttr id,"[YZBC]",yzbc
    End If 
End Function ' SetYZBC

'计算边长较差
Function comparelong(id)
    yzbc = SSProcess.GetObjectAttr(id,"[YZBC]")
    jcbc = SSProcess.GetObjectAttr(id,"[JCBC]")
    yzbc = transform(yzbc)
    jcbc = transform(jcbc)
    bcjc = Abs(yzbc-jcbc)
    SSProcess.SetObjectAttr id,"[BCJC]",bcjc
End Function ' comparelong

'设置测站检查点号名称
Function SetLinepoiname(x1,y1,x2,y2,id)
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
    idstring = SSProcess.SearchNearObjIDs(x1,y1,0,0,"",0) 
    idarr = Split(idstring,",",-1,1) '与线上点相近的点的ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName") 
        'MsgBox id
        SSProcess.SetObjectAttr id,"[CeZDH]",pointname
    ElseIf IdCount = 2 Then  
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then 
            'MsgBox id
            ExportInfo x1,y1,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
        End If 
    End If
    
    idstring = SSProcess.SearchNearObjIDs(x2,y2,0,0,"",0) 
    idarr = Split(idstring,",",-1,1) '与线上点相近的点的ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        SSProcess.SetObjectAttr id,"[JianCDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then 
            ExportInfo x2,y2,0,id 
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[JianCDH]",Firstname
        End If
    End If
End Function ' SetLinepoiname

'输出检查集函数
Function ExportInfo(x,y,z,id)
    SSProcess.AddCheckRecord strGroupName, strCheckName, "自定义脚本检查类->" & strCheckName, strPromptMessage, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function ' ExportInfo

'数据类型转换
Function transform(content)
	If content <> "" Then
		content = CDbl(content)
	Else 
		MsgBox "数据有误"
	End If
		transform = content
End Function