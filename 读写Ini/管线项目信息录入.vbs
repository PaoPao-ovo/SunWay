
' [管线CAD输出]
' 附注=
' 图幅名称=
' 作业单位=苍南县测绘院
' 委托单位=
' 测量日期=2023年7月计算机成图
' 平面坐标体系=苍南城市坐标系
' 高程体系=1985国家高程基准，等高距0.5米。
' 图式=2017年版图式
' 探测员=张三
' 测量员=张三
' 绘图员=张三
' 检查员=张三


AttrStr = "作业单位,委托单位,测量日期,平面坐标体系,高程体系,图式,探测员,测量员,绘图员,检查员"

Sub OnClick()
    
    SSProcess.ClearInputParameter

    AttrArr = Split(AttrStr,",", - 1,1)
    
    For i = 0 To UBound(AttrArr)
        SSProcess.AddInputParameter AttrArr(i) , SSProcess.ReadEpsIni("管线CAD输出", AttrArr(i) ,"") , 0 , "" , ""
    Next 'i
    
    ShowBoolen = SSProcess.ShowInputParameterDlg ("管线图廓信息录入")
    
    For i = 0 To UBound(AttrArr)
        SSProcess.WriteEpsIni "管线CAD输出", AttrArr(i) ,SSProcess.GetInputParameter(AttrArr(i))
    Next 'i
    
End Sub
