' ������
	Table_JG = "PROJECTINFO����"

'�ֶ�����
	FieldStr = "��浥λ��ַ,ί�е�λ,��Ŀ����,��Ŀ���,��Ŀ��ַ,��浥λ�绰,���õ����(m2),�ܽ������(m2),���½������(m2),�ݻ���,�����������(m2),�����ܶ�(%),�̵���(%),�滮���֤���,������ʼʱ��,�������ʱ��,Լ�����ʱ��,���Ŀ��,��Ŀ���,�����������,���Ʋ���,��ҵ����,��������,�ɹ�����˵��,���Ͻ������(m2),װ��ʽ�������(m2),��浥λ,��浥λ���ʵȼ�,�������֤����,��浥λ�绰,������Ա,�����Ա,��ҵ����,ʵ��סլ����,�滮סլ����"

'��Ϣֵ
	JGValues = ""
	JGZD = ""

Sub OnClick()
	JiLuShu = PanKong(Table_JG)
	
	If JiLuShu = 0 Then 
	
	MsgBox "��¼��Ϊ������������"
	
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