Dim arID(1000),idCount
Redim arr(1000,2) 'id�������ά����
Sub OnClick()
'��Ӵ���
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", "9510001" 
	SSProcess.SelectFilter
	fwcount = SSProcess.GetSelGeoCount
	If fwcount > 0 Then
		For i = 0 To fwcount-1
			fwid = SSProcess.GetSelGeoValue(i,"SSObj_ID") '��Χ��id
			zdkmj = DKMJ(fwid)
			ztbmj = TBMJ(fwid)
			MsgBox zdkmj & "," & ztbmj
			diff = formatnumber((zdkmj - ztbmj),2)
			sortmj(fwid) '�ڲ��ؿ�Ӵ�С����
			gs = TBGS(fwid)	'�ڲ��ؿ�ĸ���ֵ
			If gs < 10  And diff > 0 Then 
				elsearea = 0
				For j = gs -1 To 1 Step -1
					warea = Weight(arr(j,1),ztbmj,diff)
					dtb = arr(j,1) 
					newarea = dtb + warea
					SSProcess.SetObjectAttr arr(j,0),"[TBMJ]",newarea
					If elsearea = 0 Then 
							elsearea = newarea
					Else 
							elsearea = elsearea + newarea
					End If
				Next
				finaltb = zdkmj - elsearea
				SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
				MsgBox "���ƽ��1"
			End If
			If gs < 10 And diff < 0 Then 
				elsearea = 0
				For k = gs-1 To 1 Step -1 
					warea = Weight(arr(k,1),ztbmj,diff)
					dtb = arr(k,1) 
					newarea = dtb - warea
					SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
					If elsearea = 0 Then 
						elsearea = newarea
					Else	
						elsearea  = elsearea + newarea
					End if
				Next
				finaltb = zdkmj - elsearea
				SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
				MsgBox "���ƽ��2"
			End If
			If gs > 10 And diff < 0 Then
				elsearea = 0
				For k = 9 To 1 Step -1
					warea = Weight(arr(k,1),ztbmj,diff)
					'MsgBox warea
					dtb = arr(k,1) 
					MsgBox dtb
					newarea = dtb - warea
					SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
					If elsearea = 0 Then 
						elsearea = newarea
					Else	
						elsearea  = elsearea + newarea
					End if
				Next
				temp = 0
				For f = 10 To gs -1
					If temp = 0 Then
						temp = arr(f,1)
					Else
						temp = temp + arr(f,1)
					End If 
				Next
				finaltb = zdkmj - elsearea -temp
				'MsgBox zdkmj 
				SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
				MsgBox "���ƽ��3"
			End If
			If gs > 10 And diff > 0 Then
				elsearea = 0
				For k = 9 To 1 Step -1
					warea = Weight(arr(k,1),ztbmj,diff)
					dtb = arr(k,1) 
					newarea = dtb + warea
					SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
					If elsearea = 0 Then 
						elsearea = newarea
					Else	
						elsearea  = elsearea + newarea
					End if
				Next
				temp = 0
				For d = 10 To gs -1
					If temp = 0 Then
						temp = arr(d,1)
					Else
						temp = temp + arr(d,1)
					End If 
				Next
				finaltb = zdkmj - elsearea -temp
				SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
				MsgBox "���ƽ��4"
			End If
			If gs = 10 And diff < 0 Then
				elsearea = 0
				For k = 9 To 1 Step -1
					warea = Weight(arr(k,1),ztbmj,diff)
					'MsgBox warea
					dtb = arr(k,1) 
					'MsgBox dtb
					newarea = dtb - warea
					SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
					If elsearea = 0 Then 
						elsearea = newarea
					Else	
						elsearea  = elsearea + newarea
					End if
				Next
				temp = 0
				For f = 10 To gs -1
					If temp = 0 Then
						temp = arr(f,1)
					Else
						temp = temp + arr(f,1)
					End If 
				Next
				finaltb = zdkmj - elsearea -temp
				'MsgBox zdkmj 
				SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
				MsgBox "���ƽ��5"
			End If
			If gs = 10 And diff > 0 Then
				elsearea = 0
				For k = 9 To 1 Step -1
					warea = Weight(arr(k,1),ztbmj,diff)
					dtb = arr(k,1) 
					newarea = dtb + warea
					SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
					If elsearea = 0 Then 
						elsearea = newarea
					Else	
						elsearea  = elsearea + newarea
					End if
				Next
				temp = 0
				For d = 10 To gs -1
					If temp = 0 Then
						temp = arr(d,1)
					Else
						temp = temp + arr(d,1)
					End If 
				Next
				finaltb = zdkmj - elsearea -temp
				SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
				MsgBox "���ƽ��6"
			End If
		Next
	End If
End Sub

'��ȡ��ǰ��Χ�����
Function DKMJ(id)
	DKMJ = SSProcess.GetObjectAttr(id,"[DKMJ]")
	If DKMJ<>"" Then DKMJ = CDbl(DKMJ)
End Function

'��ȡ����ͼ�ߵ������(��Χ�ߵ�id)
Function TBMJ(id)
	ids = SSProcess.SearchInnerObjIDs(id,2,"9510021",1)
	ztbmj = 0
	If ids <> "" Then
		SSFunc.ScanString ids, ",", arID, idCount
			For j = 0 To idCount-1
				If ztbmj = 0 Then
					temp = SSProcess.GetObjectAttr(arID(j),"[TBMJ]")
					If temp <> "" Then mj = CDbl(temp)
					ztbmj = mj
				 Else 
				 temp =  SSProcess.GetObjectAttr(arID(j),"[TBMJ]")
				 If temp <> "" Then mj = CDbl(temp)
				 ztbmj = ztbmj + mj
				 End If
			Next
	End If
	TBMJ = ztbmj
End Function

'��ȡ��ǰ��Χ����ͼ�߸���
Function TBGS(id)
	ids = SSProcess.SearchInnerObjIDs(id,2,"9510021",1)
	If ids <> "" Then SSFunc.ScanString ids, ",", arID, idCount
	TBGS = idCount
End Function

'�������ȫ����
Function sortmj(id) '��Χ�ߵ�id
	ids = SSProcess.SearchInnerObjIDs(id,2,"9510021",1)
	If ids <> "" Then SSFunc.ScanString ids, ",", arID, idCount
	For i = 0 To idCount -1 
		arr(i,0) = arID(i)
		mj = SSProcess.GetObjectAttr(arID(i),"[TBMJ]")
		If mj <> "" Then mj = CDbl(mj)
		arr(i,1) = mj 	
	Next
	For i = 0 To idCount-1
		For j = i+1 To idCount-1
			If  arr(i,1) < arr(j,1) Then 
				maxmj = arr(i,1)
				maxid = arr(i,0)
				arr(i,1) = arr(j,1)
				arr(i,0) = arr(j,0)
				arr(j,1) = maxmj
				arr(j,0) = maxid
			End If 
		Next
	Next
End Function

'��ǰ�ؿ�ƽ��ֵ
Function Weight(tb,ztb,diff)
	temp = (tb/ztb)*Abs(diff)
	Weight = Round(temp,2)
End Function