' 表名称
	Table_RF = "PROJECTINFO人防"

'字段名称
	FieldStr = "建筑结构,住宅户数,项目名称,地上建筑面积(㎡),地上住宅建筑面积(㎡),地上其他建筑面积(㎡),地上层数,地下平时功能,地下建筑面积(㎡),地下层数,互联互通面积,防空警报控制室面积,外墙最薄掩体厚度（小于10米时填写）,板坪高差（顶板底面高出室外时填写）,编制人"

'信息值
	RFValues = ""
	RFZD = ""

Sub OnClick()
	JiLuShu = PanKong(Table_RF)
	
	If JiLuShu = 0 Then 
	
	MsgBox "记录数为空请重新输入"
	
	Exit Sub
	
	Else 
	GetInfo(Table_RF)
	
	Dim arrval(1000)
	SSFunc.ScanString RFValues, "," , arrval, vCount 
	
	Dim arrkey(1000)
	SSFunc.ScanString RFZD, "," , arrkey, kCount
	EmptyPROJECTInfo()
	For i = 2 To 16
		SqlString ="SELECT " & "KEY,VALUE" & " From " & " PROJECTINFO"
		Val = arrval(i)
		InsertInfo SqlString,"KEY,VALUE",arrkey(i) & "," & Val
	Next
	End If
End Sub

Function GetInfo(Tablename)
	
	MdbName=SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb MdbName 
	SqlString ="SELECT * FROM " & Tablename & " WHERE " & Tablename & "." & "ID > 0"
	SSProcess.OpenAccessRecordset MdbName, SqlString
	SSProcess.GetAccessRecord MdbName,SqlString,Fields,IdValues
	SSProcess.CloseAccessRecordset MdbName, SqlString
    SSProcess.CloseAccessMdb MdbName
	
	RFValues = IdValues
	RFZD = Fields

End Function 

Function InsertInfo(sql,Infile,Values)
	
	mdbName = SSProcess.GetProjectFileName 
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql
	recordc=SSProcess.GetAccessRecordCount(mdbName, sql)
	SSProcess.AddAccessRecord mdbName,sql,Infile,Values
	SSProcess.CloseAccessRecordset mdbName, sql
	SSProcess.CloseAccessMdb mdbName
End Function

Function EmptyPROJECTInfo()
	sql = "SELECT * FROM PROJECTINFO " & ";"
	mdbName = SSProcess.GetProjectFileName
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql 
	
	while  SSProcess.AccessIsEOF (mdbName, sql)=false
			SSProcess.DelAccessRecord mdbName, sql 
	wend 
		SSProcess.CloseAccessRecordset mdbName, sql
		SSProcess.CloseAccessMdb mdbName
End Function

Function PanKong(Tablename)
	mdbName = SSProcess.GetProjectFileName 	
	sql ="SELECT * FROM " & Tablename & " WHERE " & Tablename & "." & "ID > 0"
	mdbName = SSProcess.GetProjectFileName 
	SSProcess.OpenAccessMdb mdbName
	SSProcess.OpenAccessRecordset mdbName, sql
	PanKong = SSProcess.GetAccessRecordCount( mdbName,sql )

	SSProcess.CloseAccessRecordset mdbName, sql
	SSProcess.CloseAccessMdb mdbName

End Function