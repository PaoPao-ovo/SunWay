' ������
	Table_RF = "PROJECTINFO�˷�"

'�ֶ�����
	FieldStr = "�����ṹ,סլ����,��Ŀ����,���Ͻ������(�O),����סլ�������(�O),���������������(�O),���ϲ���,����ƽʱ����,���½������(�O),���²���,������ͨ���,���վ������������,��ǽ������ȣ�С��10��ʱ��д��,��ƺ�߲�������߳�����ʱ��д��,������"

'��Ϣֵ
	RFValues = ""
	RFZD = ""

Sub OnClick()
	JiLuShu = PanKong(Table_RF)
	
	If JiLuShu = 0 Then 
	
	MsgBox "��¼��Ϊ������������"
	
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