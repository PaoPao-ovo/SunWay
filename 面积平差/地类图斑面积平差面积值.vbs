	'�����ƽ����ֵ
	Dim PCYZ,arr(10000,2)

	Dim arID(10000),idCount

	Sub OnClick()
	'��Ӵ���
		'��������
		SSProcess.ClearInputParameter 
		SSProcess.AddInputParameter "���ƽ����ֵ","", 0,"50","���ƽ����ֵ����λƽ���ף�"
		ret =SSProcess.ShowInputParameterDlg ("���ƽ����ֵ��Ϣ¼����")
		SSProcess.UpdateScriptDlgParameter 1
		PCYZ = SSProcess.GetInputParameter("���ƽ����ֵ")
		
		'��ȡѡ��
		SSProcess.ClearSelection 
		SSProcess.ClearSelectCondition
		SSProcess.SetSelectCondition "SSObj_Code", "==", "9510001" 
		SSProcess.SelectFilter
		fwcount = SSProcess.GetSelGeoCount
		If fwcount > 0 Then
			For i = 0 To fwcount-1
				fwid = SSProcess.GetSelGeoValue(i,"SSObj_ID")
				zdkmj = DKMJ(fwid) '�ܵؿ����
				ztbmj = ITBMJ(fwid) '��ͼ�����
				MsgBox zdkmj & "," & ztbmj
				diff = formatnumber((zdkmj - ztbmj),2)
				'MsgBox diff
				count = GetAvaliableCount(PCYZ,fwid) '����ͼ�߸���
				'MsgBox count
				GetAvaliableArr PCYZ,fwid '����ͼ��˳������
				If count < 10  And diff > 0 Then 
					elsearea = 0
					For j = count -1 To 1 Step -1
						'MsgBox arr(j,1)
						warea = Weight(arr(j,1),ztbmj,diff)
						dtb = arr(j,1) 
						newarea = dtb + warea
						'MsgBox newarea
						SSProcess.SetObjectAttr arr(j,0),"[TBMJ]",newarea
						If elsearea = 0 Then 
							elsearea = newarea
						Else 
							elsearea = elsearea + newarea
						End If
					Next
					'MsgBox elsearea
					finaltb = zdkmj - elsearea
					'MsgBox finaltb
					SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
					MsgBox "���ƽ��һ"
				End If
				If count < 10 And diff < 0 Then 
					elsearea = 0
					For k = count-1 To 1 Step -1 
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
					MsgBox "���ƽ���"
				End If
				If count > 10  And diff < 0 Then
					elsearea = 0
					For k = 9 To 1 Step -1
						are = QSarea()
						warea = Weight(arr(k,1),are,diff)
						dtb = arr(k,1) 
						newarea = dtb - warea
						SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
						If elsearea = 0 Then 
							elsearea = newarea
						Else	
							elsearea  = elsearea + newarea
						End if
					Next
					temp = ztbmj - are
					'MsgBox temp
					finaltb = zdkmj - elsearea - temp
					SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
					MsgBox "���ƽ����"
				End If
				If count > 10  And diff > 0 Then
					elsearea = 0
					For k = 9 To 1 Step -1
						are = QSarea()
						warea = Weight(arr(k,1),are,diff)
						'MsgBox warea
						dtb = arr(k,1) 
						newarea = dtb + warea
						SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
						If elsearea = 0 Then 
							elsearea = newarea
						Else	
							elsearea  = elsearea + newarea
						End if
					Next
					'MsgBox elsearea
					temp = ztbmj - are
					finaltb = zdkmj - elsearea - temp
					SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
					MsgBox "���ƽ����"
				End If
				If count = 10  And diff < 0 Then
					elsearea = 0
					For k = 9 To 1 Step -1
						are = QSarea()
						warea = Weight(arr(k,1),are,diff)
						dtb = arr(k,1) 
						newarea = dtb - warea
						SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
						If elsearea = 0 Then 
							elsearea = newarea
						Else	
							elsearea  = elsearea + newarea
						End if
					Next
					temp = ztbmj - are
					'MsgBox temp
					finaltb = zdkmj - elsearea - temp
					SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
					MsgBox "���ƽ����"
				End If
				If count = 10  And diff > 0 Then
					elsearea = 0
					For k = 9 To 1 Step -1
						are = QSarea()
						warea = Weight(arr(k,1),are,diff)
						'MsgBox warea
						dtb = arr(k,1) 
						newarea = dtb + warea
						SSProcess.SetObjectAttr arr(k,0),"[TBMJ]",newarea
						If elsearea = 0 Then 
							elsearea = newarea
						Else	
							elsearea  = elsearea + newarea
						End if
					Next
					'MsgBox elsearea
					temp = ztbmj - are
					finaltb = zdkmj - elsearea - temp
					SSProcess.SetObjectAttr arr(0,0),"[TBMJ]",finaltb
					MsgBox "���ƽ����"
				End If
				If count = 0 Then 
					MsgBox "�������������" & PCYZ & "����"
					Exit For
				End If
			Next
		End If
	End Sub

	'��ȡ��ǰ��Χ�����
	Function DKMJ(id)
		DKMJ = SSProcess.GetObjectAttr(id,"[DKMJ]")
		If DKMJ<>"" Then DKMJ = CDbl(DKMJ)
	End Function

	'��ȡ��ǰ�ؿ��ڷ���Ҫ���ͼ�߸���
	Function GetAvaliableCount(yz,id)
		ids = SSProcess.SearchInnerObjIDs(id,2,"9510021",1)
		If ids <> "" Then 
			icount = 0
			SSFunc.ScanString ids, ",", arID, idCount
			For i = 0 To idCount - 1
				tbmj = SSProcess.GetObjectAttr(arID(i),"[TBMJ]")
				tbid = arID(i) 
				tbmj = transform(tbmj) 'ת��Ϊ��������
				yz = transform(yz)
				If tbmj > yz Then icount = icount +1
			Next
		End If
		GetAvaliableCount = icount
	End Function

	'��������ת��
	Function transform(content)
		If content <> "" Then
			content = CDbl(content)
		Else 
			MsgBox "�������Ϊ�յ���"
		End If
		transform = content
	End Function

	'��ȡ����Ҫ��ĵؿ������ID������
	Function GetAvaliableArr(yz,id)
		ids = SSProcess.SearchInnerObjIDs(id,2,"9510021",1)
		yz = transform(yz)
		i = 0
		If ids <> "" Then 
			SSFunc.ScanString ids, ",", arID, idCount
			For n = 0 To idCount - 1
				tbmj = SSProcess.GetObjectAttr(arID(n),"[TBMJ]")
				tbid = arID(n) 
				tbmj = transform(tbmj) 'ת��Ϊ��������
				If tbmj > yz Then 
					arr(i,0) = arID(n)
					arr(i,1) = tbmj
					'MsgBox arr(i,1)
					i = i + 1
				End If
			Next
			tempcount = GetAvaliableCount(yz,id)
			'MsgBox tempcount
			'MsgBox arr(0,1)
		For i = 0 To tempcount-1
			For j = i+1 To tempcount-1
				'MsgBox arr(i,1) & "," & arr(j,1)
				If  arr(i,1) < arr(j,1) Then 
					minmj = arr(i,1)
					minid = arr(i,0)
					arr(i,1) = arr(j,1)
					arr(i,0) = arr(j,0)
					arr(j,1) = minmj
					arr(j,0) = minid
					'MsgBox ""
				End If 
			Next
		Next
		End If
	End Function

	'��ȡ����Ҫ��ĵؿ����֮��
	Function GetTotalArea(yz,id)
		iarea = 0
		yz = transform(yz)
		ids = SSProcess.SearchInnerObjIDs(id,2,"9510021",1)
		If ids <> "" Then 
			SSFunc.ScanString ids, ",", arID, idCount
			'MsgBox idCount
			For i = 0 To idCount - 1
				tbmj = SSProcess.GetObjectAttr(arID(i),"[TBMJ]")
				tbid = arID(i) 
				tbmj = transform(tbmj) 'ת��Ϊ��������
				If tbmj > yz Then 
					If iarea = 0 Then
						iarea = tbmj
					Else 
						iarea = iarea + tbmj
					End If
				End If
			Next
		End If
		GetTotalArea = transform(iarea)
		'MsgBox GetTotalArea
	End Function

	'��ǰ�ؿ�ƽ��ֵ
	Function Weight(tb,ztb,diff)
		temp = (tb/ztb)*Abs(diff)
		Weight = Round(temp,2)
	End Function

	'��ȡ����ͼ�ߵ������(��Χ�ߵ�id)
	Function ITBMJ(id)
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
		ITBMJ = ztbmj
	End Function

	'��ȡǰʮ������ͼ�����
	Function QSarea()
		are = 0
		For i = 0 To 9
			If are = 0 Then
				are = arr(0,1)
			Else 
				are = are + arr(i,1)
			End If
		QSarea = are
		Next
		'MsgBox QSarea
	End Function