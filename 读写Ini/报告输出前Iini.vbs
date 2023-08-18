
' [管线报告信息]
' 编号 = ""
' 项目名称 = ""
' 项目地址 = ""
' 设计单位 = ""
' 建设单位 = ""
' 委托单位 = ""
' 测绘单位 = ""
' 外业时间 = ""
' 点最大较差值 = ""
' 高程最大较差值 = ""

'======================================配置Ini&输入框字段============================================

KeyStr = "编号,项目名称,项目地址,设计单位,建设单位,委托单位,外业时间,测绘时间,点最大较差值,高程最大较差值"

'==========================================功能主体=====================================================

'功能入口
Sub OnClick()
    SSProcess.ClearInputParameter
    
    KeyArr = Split(KeyStr,",", - 1,1)
    
    For i = 0 To UBound(KeyArr) - 2
        SSProcess.AddInputParameter KeyArr(i) , SSProcess.ReadEpsIni("管线报告信息", KeyArr(i) ,"") , 0 , "" , ""
    Next 'i
    
    ShowBoolen = SSProcess.ShowInputParameterDlg ("管线报告信息录入")
    
    For i = 0 To UBound(KeyArr)
        SSProcess.WriteEpsIni "管线报告信息", KeyArr(i) ,SSProcess.GetInputParameter(KeyArr(i))
    Next 'i
    SSProcess.OpenAccessMdb SSProcess.GetProjectFileName
    
    SqlStr = "Update 地下管线点属性表 SET XMBH = " & "'" & SSProcess.ReadEpsIni("管线报告信息", "编号" ,"") & "'"
    SsProcess.ExecuteAccessSql SSProcess.GetProjectFileName,SqlStr
    
    SqlStr = "Update 地下管线线属性表 SET XMBH = " & "'" & SSProcess.ReadEpsIni("管线报告信息", "编号" ,"") & "'"
    SsProcess.ExecuteAccessSql SSProcess.GetProjectFileName,SqlStr

    SSProcess.CloseAccessMdb SSProcess.GetProjectFileName
    
    SSProcess.MapMethod "clearattrbuffer",  "地下管线点属性表"
    SSProcess.MapMethod "clearattrbuffer",  "地下管线线属性表"
    SSProcess.RefreshView()
End Sub' OnClick