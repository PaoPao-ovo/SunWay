' 表名称
	Table_TD = "PROJECTINFO土地勘测定界"

'字段名称
	FieldStr = "主管部门,勘测单位,勘测定界单位签注,单位地址,单位负责人,土地座落,坐标系,所有制性质,用地单位,用途,申请日期,相关文件,经办人,联系电话,资料复审人,资料审核人,项目名称,项目编号,项目负责人,勘测阶段"

'信息值
	TDValues = ""
	TDZD = ""

Sub OnClick()
	JiLuShu = PanKong(Table_TD)
	
	If JiLuShu = 0 Then 
	
	MsgBox "记录数为空请重新输入"
	
	Exit Sub
	
	Else 
	GetInfo(Table_TD)
	
	Dim arrval(1000)
	SSFunc.ScanString TDValues, "," , arrval, vCount 
	
	Dim arrkey(1000)
	SSFunc.ScanString TDZD, "," , arrkey, kCount
	EmptyPROJECTInfo()
	For i = 2 To 21
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
	
	TDValues = IdValues
	TDZD = Fields

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