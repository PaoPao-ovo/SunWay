Dim LvArr(10000),Varr(10000),Neararr(10000),Stack(10000)
Dim idCount
Dim scount:scount=1
Dim dk:dk=1
Sub OnClick()
'---------获取所有的绿地范围面要素-------------
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9470103
	SSProcess.SelectFilter
	'geocount = SSProcess.GetSelGeoCount()
	'MsgBox geocount
	LvDicount = SSProcess.GetSelGeoCount()
	For i = 0 To LvDicount-1
		id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
		LvArr(i) = id
	Next
	For i = 0 To LvDicount-1
		ids = SSProcess.SearchInPolyObjIDs(LvArr(i),2,9470103,0,1,1)
		If ids <> "" Then Neararr(i) = ids
	Next
	Stack(0) = LvArr(0)
		For i = 0 To LvDicount-1
			str = Neararr(i)
			SSFunc.ScanString str, ",", Varr, idCount
			If i = 0 Then
				For j = 0 To idCount-1
					Stack(scount) = Varr(j)
					scount = scount + 1
				Next
			Else
				For j = 0 To idCount-1
					For k = 0 To scount-1
						If Stack(k) <> Varr(j)  Then
							Stack(scount) = Varr(j)
							scount = scount + 1	
						End If 
					Next
				Next
			End If 
		Next	
		msgbox scount
End Sub