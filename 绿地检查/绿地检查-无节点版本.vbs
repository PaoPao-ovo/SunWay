Dim LvArr(10000),Varr(10000),Diffarr(10000),arID(10000)
Dim idCount
Dim dk:dk = 1
Sub OnClick() 
'---------获取所有的绿地范围面要素-------------
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9470103
	SSProcess.SelectFilter
	'geocount = SSProcess.GetSelGeoCount()
	'MsgBox geocount
	LvDicount = SSProcess.GetSelGeoCount()
	ids = ""
	For i = 0 To LvDicount-1
		id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
		SSProcess.SetObjectAttr CInt(id), "[BH]", ""
		LvArr(i) = id
		If ids = "" Then 
			ids = id
		Else 
			ids = ids & "," & id
		End If 
	Next	
	
	SSProcess.MergePolygon ids,0.01,0,0 '构造轮廓面
	
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 9470103
	SSProcess.SelectFilter
	LvDicount = SSProcess.GetSelGeoCount()
	
	Dim k:k=0
	For i = 0 To LvDicount-1
		id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
		If id <> LvArr(i) Then 
			Diffarr(k) = id
			'SSProcess.SetObjectAttr CInt(id), "SSObj_Code", "2"
			k = k + 1
		End If 
	Next
	For i = 0 To k-1
		ids = SSProcess.SearchInnerObjIDs(Diffarr(i),2,"9470103",1)
		If ids <> "" Then
			SSFunc.ScanString ids, ",", arID, idCount
			For l=0 To idCount-1
				SSProcess.SetObjectAttr CInt(arID(l)), "[BH]", "DK"&i+1
			Next
		End If
	Next
	For i = 0 To k-1
		SSProcess.DeleteObject(Diffarr(i))
	Next
End Sub
