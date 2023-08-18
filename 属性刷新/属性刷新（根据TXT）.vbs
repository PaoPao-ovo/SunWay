Sub OnClick()
	'获取地类名称
	SSProcess.ClearSelection 
	SSProcess.ClearSelectCondition
	SSProcess.SetSelectCondition "SSObj_Code", "==", 7320
	SSProcess.SelectFilter
	geoCount = SSProcess.GetSelGeoCount()
	
	'获取文件名称
	Path = SSProcess.GetSysPathName (7)
	Name= "地类图斑.txt"
	FileName = Path + Name
	'判断路径
	If FileName="" Then
		Msgbox "路径不能为空"
		Exit Sub
	End If
	
	'读取TXT文件
	Dim fso,ts,chLine,strs(10000)
	Set fso=CreateObject("Scripting.FileSystemObject")
	
	'只读方式打开
	Set ts = fso.OpenTextFile(FileName , 1)
	
	'按行读取
	Dim n:n=0
	
	'判断是否相等
	SSProcess.RemoveCheckRecord "地类检查", "地类名称检查"
	If geoCount>0 Then 
		For i = 0 To geoCount-1
			Dim bs:bs=0
			Dlmc = SSProcess.GetSelGeoValue( i, "[DLMC]" )
				Do While Not ts.AtEndOfStream
					chLine=ts.ReadLine
					
					'分解字符串
					SSFunc.ScanString chLine,",",strs,count	
				If Dlmc = strs(1) Then 
					SSProcess.SetSelGeoValue i, "[DLBM]", strs(0)
					SSProcess.AddSelGeoToSaveGeoList i 
				End If 	
				If Dlmc = strs(1) Then 
					bs=1
				End If 
				SSProcess.SaveBufferObjToDatabase
				n=n+1
				Loop
				If bs =0 Then
				geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
				geoType = SSProcess.GetSelGeoValue(i,"SSObj_Type")
				SSProcess.GetSelGeoPoint i, 0, x, y, z, pointtype, name 
				SSProcess.AddCheckRecord "地类检查", "地类名称检查", "自定义脚本检查类->地类名称检查", "地类名称不存在", x, y, z, pointtype, geoID, ""
				End If
		Next
	End If
	ts.Close
	SSProcess.ShowCheckOutput
End Sub
	
	