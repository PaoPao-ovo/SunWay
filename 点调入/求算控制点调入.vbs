Sub OnClick()
'添加代码
	FileName=SSProcess.SelectFileName(1,"",0,"TXT Files(*.txt)|*.txt|DAT Files (*.dat)|*.dat|All Files (*.*)|*.*||")
	If FileName="" Then
		msgbox "路径不能为空"
		Exit Sub
	End if
	Dim  fso,ts,chLine,strs(10000)
	Set fso=CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(FileName , 1)
	dim n:n=0
	Do While Not ts.AtEndOfStream
	chLine=ts.ReadLine
	n=n+1
	'chLine=Trim(chLine)
	'msgbox chLine
	SSFunc.ScanString chLine," ",strs,count
	If count>0 then
		If strs(0)<>"QS" OR  strs(0)<>"JC"then
			 msgbox"第"&n&"行前缀有误"
	End If 
		If strs(0) ="QS" then
				SSProcess.PushUndoMark
				SSProcess.CreateNewObj 0
				SSProcess.SetNewObjValue "SSObj_Code", 1130211
				SSProcess.AddNewObjPoint strs(3), strs(2), 0, 0, strs(1)
				'SSProcess.SetNewObjValue "SSObj_PointName", strs(1)
				SSProcess.AddNewObjToSaveObjList
				msgbox "第"&n&"行执行成功"
		End if
				If strs(0) ="JC" then
					SSProcess.PushUndoMark
					SSProcess.CreateNewObj 0
					SSProcess.SetNewObjValue "SSObj_Code", 9130311
					SSProcess.AddNewObjPoint strs(3), strs(2), 0, 0, strs(1)
					'SSProcess.SetNewObjValue "SSObj_PointName", strs(1)
					SSProcess.AddNewObjToSaveObjList
					msgbox "第"&n&"行执行成功"
				End if
	End if
	SSProcess.SaveBufferObjToDatabase
	Loop
	ts.Close
	msgbox "执行成功"
End Sub
