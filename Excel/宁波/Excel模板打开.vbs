Sub OnClick()
	'ѡ��·��
	Pathname=SSProcess.SelectPathName 
	If Pathname="" Then Exit Sub
	CreateExcelFile Pathname&"�����������̹滮��Ϣ����ģ��.xlsx"
End Sub

'����excel�ļ�����̨���µ�ģ��excel���Ƶ��˹�ѡ���·����
Function CreateExcelFile(FileName)
	CreateExcelFile=0
	TemplateXlsName = SSProcess.GetSysPathName(7) & "���ģ��\�����������̹滮��Ϣ����ģ��.xlsx"
	Set fso=CreateObject("Scripting.FileSystemObject")
	fso.copyFile TemplateXlsName,FileName
	OpenEXCEL FileName
End Function

'�򿪱�
Function OpenEXCEL(ModelExcel)
	Set oExcel= CreateObject("Excel.Application") 
	oExcel.Application.Visible = true
	Set oWb = oExcel.Workbooks.open(ModelExcel)
End Function
