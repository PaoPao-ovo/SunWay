' ������
	Table_ZFL = "PROJECTINFO������"

'�ֶ�����
	FieldStr = "�����������,�ɹ�����˵��,���Ʋ���,��浥λ,��浥λ��ַ,��浥λ�绰,��浥λ���ʵȼ�,���Ŀ��,�������֤����,�������ʱ��,������ʼʱ��,Լ�����ʱ��,�滮���֤���,��������,��Ŀ����,��Ŀ��ַ,��Ŀ���,��Ŀ���,��ҵ����,ί�е�λ,������Ա,�����Ա,��ҵ����"

'��Ϣֵ
	ZFLValues = ""
	ZFLZD = ""

Sub OnClick()
	JiLuShu = PanKong(Table_ZFL)
	
	If JiLuShu = 0 Then 
	
	MsgBox "��¼��Ϊ������������"
	
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