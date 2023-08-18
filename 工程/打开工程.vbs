Sub OnClick()
'添加代码
	'生成目标工程路径字符串
	SysPath = SSProcess.GetSysPathName(7)
	FileNmae = "打开工程文件.edb"
	Path = SysPath + FileNmae
	
	'关闭当前工程
	Result = SSProcess.CloseDatabase()
	If Result =1 Then 
		MsgBox "当前工程关闭成功"
	Else
		MsgBox "当前工程关闭失败"
	End If
	
	'打开新工程
	NewResult = SSProcess.OpenDatabase Path
	If NewResult = 0 Then 
	MsgBox "执行失败，请检查路径"
	Else
	MsgBox "执行成功"
	End If
End Sub