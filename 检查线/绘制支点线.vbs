'���ϵ������
Dim PointArr(2,3)
'��鼯����
Dim strGroupName:strGroupName = "�ظ�����"
'��鼯�����
Dim strCheckName:strCheckName = "֧���߼��"
'�����־
Dim strPromptMessage:strPromptMessage = "���ֶ���д��վ��ź�֧վ���"

'��ں���
Function zd(id)
    GetOnlinePoint(id)
    SearchNear(id)
End Function ' zd

'��ֵ����
Function SearchNear(id)
    x1 = PointArr(0,0)
    y1 = PointArr(0,1)
    x2 = PointArr(1,0)
    y2 = PointArr(1,1)
    SetLinepoiname x1,y1,x2,y2,id
    SetProp x1,y1,x2,y2,id
End Function ' SearchNear

'��ȡ���ϵĿռ����Ϣ
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

'�����ߵķ���ֵ��ˮƽ����
Function SetProp(x1,y1,x2,y2,id)
    longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
    longtitude = transform(longtitude)
    longtitude = formatnumber(longtitude,3)
    'If x1 < x2 And y1 < y2 Then
    '    x = x2 - x1
    '    y = y2 - y1
    '    angles = 270 + SSProcess.RadianToDms(atn(abs(x/y)))
    'End If 
    'If x1 > x2 And y1 < y2 Then
    '    x = x2 - x1
    '    y = y2 - y1
    '    angles = 90 - SSProcess.RadianToDms(atn(abs(x/y)))
    'End If 
    'If x1 < x2 And y1 > y2 Then
    '    x = x2 - x1
    '    y = y2 - y1
    '    angles = 90 + SSProcess.RadianToDms(atn(abs(x/y)))
    'End If 
    'If x1 > x2 And y1 > y2 Then
    '    x = x2 - x1
    '    y = y2 - y1
    '    angles = 180 + SSProcess.RadianToDms(atn(abs(y/x)))
    'End If
    'angarr = Split(angles,".",-1,1)
    'If UBound(angarr) > 0 Then
    'str = angarr(1)
    'dd = ""
    'ss = ""
    'If Len(str) > 4 Then 
    'dd = Mid(str,1,2)
    'ss = Mid(str,3,2)
    'End If 
    'If Len(str) = 3 Then
    'dd = Mid(str,1,2)
    'ss = Mid(str,3,1) & "0"
    'End If 
    ' If Len(str) = 2 Then
    'dd = Mid(str,1,2)
    'ss = "00"
    'End If 
    ' If Len(str) = 1 Then
    'dd = Mid(str,1,1) & "0"
    'ss ="00"
    'End If 
    'If Len(str) = 0 Then
    'dd = "00"
    'ss = "00"
    'End If 
    'ElseIf UBound(angarr) = 0 Then
    '    dd = "00"
    '    ss = "00"
    'End IF
    SSProcess.SetObjectAttr id,"[ShuiPJL]",longtitude
    'SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "��" & dd & "��" & ss & "��"
End Function ' SetProp

'���ò�վ֧վ�������
Function SetLinepoiname(x1,y1,x2,y2,id)
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
    idstring = SSProcess.SearchNearObjIDs(x1,y1,0,0,"",0) 
    idarr = Split(idstring,",",-1,1) '�����ϵ�����ĵ��ids
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
    idarr = Split(idstring,",",-1,1) '�����ϵ�����ĵ��ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        SSProcess.SetObjectAttr id,"[ZhiZDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then 
            ExportInfo x2,y2,0,id 
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[ZhiZDH]",Firstname
        End If
    End If
End Function ' SetLinepoiname

'�����鼯����
Function ExportInfo(x,y,z,id)
    SSProcess.AddCheckRecord strGroupName, strCheckName, "�Զ���ű������->" & strCheckName, strPromptMessage, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function ' ExportInfo

'��������ת��
Function transform(content)
	If content <> "" Then
		content = CDbl(content)
	Else 
		MsgBox "��������"
	End If
		transform = content
End Function