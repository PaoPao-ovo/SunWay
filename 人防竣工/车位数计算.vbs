
'����ǻ�����λ����
SWCodes = "9460033"

'���ڷǻ���������
SNCode = "9460003"	

Sub OnClick()
	SW()
	SN()
End Sub

'���ڳ�λ
Function SN()
	SSProcess.PushUndoMark
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_code", "==", 9450013
	SSProcess.SelectFilter
	geocount = SSProcess.GetSelGeoCount()
	
	If geocount > 0 Then
		For j=0 To geocount-1 
			ID = SSProcess.GetSelGeoValue(j, "SSObj_ID")
			InnerIds = SSProcess.SearchInnerObjIDs(ID, 2,SNCode, 1)
			SNArr = Split(InnerIds,",",-1,0)
			For i = 0 To UBound(SNArr)
				TsMj = SSProcess.GetObjectAttr( SNArr(i), "SSObj_Area")
				TsMj = transform(TsMj)
				CWMJ = formatnumber(TsMj,2)
				SSProcess.SetObjectAttr SNArr(i), "[CWMJ]", CWMJ
				ZSMJ = CWMJ * 1.8
				FJDCCWGS = Int(ZSMJ)
				SSProcess.SetObjectAttr SNArr(i), "[FJDCCWGS]", FJDCCWGS
			Next
		Next
	End If
End Function

'���⳵λ
Function SW()
	SSProcess.PushUndoMark
	SSProcess.ClearSelection
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_code", "==", 9450013
	SSProcess.SelectFilter
	geocount = SSProcess.GetSelGeoCount()
	
	If geocount > 0 Then
		For j=0 To geocount-1 
			ID = SSProcess.GetSelGeoValue(j, "SSObj_ID")
			InnerIds = SSProcess.SearchInnerObjIDs(ID, 2,SWCodes, 1)
			SWArr = Split(InnerIds,",",-1,0)
			For i = 0 To UBound(SWArr)
				Lx = SSProcess.GetObjectAttr( SWArr(i), "[FJDCLB]")
				TsMj = SSProcess.GetObjectAttr( SWArr(i), "SSObj_Area")
				TsMj = transform(TsMj)
				CWMJ = formatnumber(TsMj,2)
				SSProcess.SetObjectAttr SWArr(i), "[CWMJ]", CWMJ
				If Lx = "¶��" Then 
					ZSMJ = CWMJ * 1.5
					FJDCCWGS = Int(ZSMJ)
					SSProcess.SetObjectAttr SWArr(i), "[FJDCCWGS]", FJDCCWGS
				ElseIf Lx = "·��" Then
					ZSMJ = CWMJ * 1.2
					FJDCCWGS = Int(ZSMJ)
					SSProcess.SetObjectAttr SWArr(i), "[FJDCCWGS]", FJDCCWGS
				End If 
			Next
		Next
	End If
End Function

'��������ת��
Function transform(content)
	If content <> "" Then
		content = CDbl(content)
	Else 
		MsgBox "CWMJ�ֶ�Ϊ�գ�����"
	End If
	transform = content
End Function
				