' 表名称
	Table_ZFL = "PROJECTINFO正负零"

'字段名称
	FieldStr = "已有资料情况,成果内容说明,控制测量,测绘单位,测绘单位地址,测绘单位电话,测绘单位资质等级,测绘目的,测绘资质证书编号,测量完成时间,测量开始时间,约定完成时间,规划许可证编号,质量控制,项目名称,项目地址,项目类别,项目编号,作业内容,委托单位,编制人员,审核人员,作业依据"

'信息值
	ZFLValues = ""
	ZFLZD = ""

Sub OnClick()
	JiLuShu = PanKong(Table_ZFL)
	
	If JiLuShu = 0 Then 
	
	MsgBox "记录数为空请重新输入"
	
	Exit Sub
	
	Else 
	GetInfo(Table_ZFL)
	
	Dim arrval(1000)
	SSFunc.ScanString ZFLValues, "," , arrval, vCount 
	
	Dim arrkey(1000)
	SSFunc.ScanString ZFLZD, "," , arrkey, kCount
	EmptyPROJECTInfo()
	For i = 2 To 24
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
	
	ZFLValues = IdValues
	ZFLZD = Fields

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