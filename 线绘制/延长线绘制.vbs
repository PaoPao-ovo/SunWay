'??????
Dim PointArr(10000,3)
'x??????????
Dim xdiff
'y??????????
Dim ydiff
Function hz(id) '??????
    GetOnlinePoint(id)
    proportions()
    DrawLine(id)
End Function ' hz

'?????????
Function DrawLine(id)
	llong = 4
    'llong = transform(llong)
    llong = llong /2
    lines = GetLineLength(id)
    xpro = Abs(xdiff/lines)
    xplus = xpro * llong
    ypro = Abs(ydiff/lines)
    yplus = ypro * llong 
    '????
    If xdiff = 0 Then
        If PointArr(0,1) > PointArr(1,1) Then 
            PointArr(0,1) = PointArr(0,1) + llong
            PointArr(1,1) = PointArr(1,1) - llong
        ElseIf PointArr(0,1) < PointArr(1,1) Then
            PointArr(1,1) = PointArr(1,1) + llong
            PointArr(0,1) = PointArr(0,1) - llong
        End If 
    End If 
    '????
    If ydiff = 0 Then
        If PointArr(0,0) > PointArr(1,0) Then 
            PointArr(0,0) = PointArr(0,0) + llong
            PointArr(1,0) = PointArr(1,0) - llong
        ElseIf PointArr(0,0) < PointArr(1,0) Then
            PointArr(1,0) = PointArr(1,0) + llong
            PointArr(0,0) = PointArr(0,0) - llong
        End If 
    End If 
    'งา??
    If xdiff>0 And ydiff<0 Then 
        PointArr(0,0) = PointArr(0,0) + xplus
        PointArr(0,1) = PointArr(0,1) - yplus
        PointArr(1,0) = PointArr(1,0) - xplus
        PointArr(1,1) = PointArr(1,1) + yplus
    End If 

    If xdiff<0 And ydiff>0 Then
        PointArr(0,0) = PointArr(0,0) - xplus
        PointArr(0,1) = PointArr(0,1) + yplus
        PointArr(1,0) = PointArr(1,0) + xplus
        PointArr(1,1) = PointArr(1,1) - yplus
    End If 
    
    If xdiff>0 And ydiff>0 Then
        PointArr(0,0) = PointArr(0,0) + xplus
        PointArr(0,1) = PointArr(0,1) + yplus
        'MsgBox PointArr(1,0)
        PointArr(1,0) = PointArr(1,0) - xplus
        PointArr(1,1) = PointArr(1,1) - yplus
        'MsgBox yplus
    End If 

    If xdiff<0 And ydiff<0 Then
        PointArr(0,0) = PointArr(0,0) - xplus
        PointArr(0,1) = PointArr(0,1) - yplus
        PointArr(1,0) = PointArr(1,0) + xplus
        PointArr(1,1) = PointArr(1,1) + yplus
    End If  
    MakeLine PointArr(0,0),PointArr(0,1),PointArr(1,0),PointArr(1,1)
    SSProcess.DeleteObject (id)
End Function ' DrawLine

'????????
Function GetLineLength(id)
    linelength = SSProcess.GetObjectAttr(id, "SSObj_Length")
    GetLineLength = linelength
    GetLineLength = transform(GetLineLength)
    'MsgBox GetLineLength
End Function ' GetLineLength

'??????????????
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
            Next
    'MsgBox PointArr(1,0)
End Function ' GetOnlinePoint

'???????
Function proportions()
    xdiff = PointArr(0,0) - PointArr(1,0)
    ydiff = PointArr(0,1) - PointArr(1,1)
End Function ' proportions

'???????????
Function transform(content)
	If content <> "" Then
		content = CDbl(content)
	Else 
		MsgBox "????????"
	End If
		transform = content
End Function

'?????????
Function MakeLine(x1,y1,x2,y2)
		SSProcess.CreateNewObj 1
		SSProcess.SetNewObjValue "SSObj_Code", "1"
        'MsgBox x1 & "," & y1 & ";" & x2 & "," & y2
		SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
		SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
		SSProcess.AddNewObjToSaveObjList
		SSProcess.SaveBufferObjToDatabase
End Function 