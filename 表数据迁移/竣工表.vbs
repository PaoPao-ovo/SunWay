' 表名称
	Table_JG = "PROJECTINFO竣工"

'字段名称
	FieldStr = "测绘单位地址,委托单位,项目名称,项目编号,项目地址,测绘单位电话,总用地面积(m2),总建筑面积(m2),地下建筑面积(m2),容积率,建筑基底面积(m2),建筑密度(%),绿地率(%),规划许可证编号,测量开始时间,测量完成时间,约定完成时间,测绘目的,项目类别,已有资料情况,控制测量,作业内容,质量控制,成果内容说明,地上建筑面积(m2),装配式建筑面积(m2),测绘单位,测绘单位资质等级,测绘资质证书编号,测绘单位电话,编制人员,审核人员,作业依据,实测住宅户数,规划住宅户数"

'信息值
	JGValues = ""
	JGZD = ""

Sub OnClick()
	JiLuShu = PanKong(Table_JG)
	
	If JiLuShu = 0 Then 
	
	MsgBox "记录数为空请重新输入"
	
	Exit Sub
	
	Else 
	GetInfo(Table_JG)
	
	Dim arrval(1000)
	SSFunc.ScanString JGValues, "," , arrval, vCount 
	
	Dim arrkey(1000)
	SSFunc.ScanString JGZD, "," , arrkey, kCount
	EmptyPROJECTInfo()
	For i = 2 To 35
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
	
	JGValues = IdValues
	JGZD = Fields

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