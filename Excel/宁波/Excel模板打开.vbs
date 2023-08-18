Sub OnClick()
	'选择路径
	Pathname=SSProcess.SelectPathName 
	If Pathname="" Then Exit Sub
	CreateExcelFile Pathname&"宁波建筑工程规划信息导入模板.xlsx"
End Sub

'创建excel文件，将台面下的模版excel复制到人工选择的路径。
Function CreateExcelFile(FileName)
	CreateExcelFile=0
	TemplateXlsName = SSProcess.GetSysPathName(7) & "输出模板\宁波建筑工程规划信息导入模板.xlsx"
	Set fso=CreateObject("Scripting.FileSystemObject")
	fso.copyFile TemplateXlsName,FileName
	OpenEXCEL FileName
End Function

'打开表
Function OpenEXCEL(ModelExcel)
	Set oExcel= CreateObject("Excel.Application") 
	oExcel.Application.Visible = true
	Set oWb = oExcel.Workbooks.open(ModelExcel)
End Function
