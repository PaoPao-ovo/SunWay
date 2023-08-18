
	'配置需要打开的EDB名称
	FileString = "综合管线竣工图,专业管线竣工图"

	Function OpenProject()
		SSProcess.ClearInputParameter
		SSProcess.AddInputParameter "工程名称" , "" , 0 , FileString , "选择工程文件"
		result = SSProcess.ShowInputParameterDlg ("工程选择")
		If result = 1 Then
			SSProcess.UpdateScriptDlgParameter 1
			FileName = SSProcess.GetInputParameter ("工程名称")
			FileName = FileName + ".edb"
			SystemPath = SSProcess.GetProjectFileName()
			SystemPath = Replace(SystemPath,".edb","")
			FilePath = SystemPath + "\"+FileName
			exist = IsFileExists(FilePath)
			If exist = 1 Then
				SSProcess.OpenDatabase(FilePath)
			Else
				MsgBox "文件不存在请检查：" & FilePath & " 是否存在"
			End If 
		End If 
	End Function

	Function IsFileExists(filepath)
		Dim fso
		Set fso=CreateObject("Scripting.FileSystemObject")    
		if fso.fileExists(filepath)= false Then
			IsFileExists = 0
		Else 
			IsFileExists = 1
		End If 
		set fso = nothing 
	End Function 

	Sub OnClick()
	'添加代码
		'窗口配置
		OpenProject()
	End Sub