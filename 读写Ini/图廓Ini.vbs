
' 清除之前的输入框

' [验线图廓信息]
' 测绘单位 = 苍南县测绘院
' 测绘时间及方式 = 2020年12月数字测图
' 坐标系 = 苍南城市坐标系
' 高程体系 = 1985国家高程基准，等高距为0.5m。
' 图示 = 2007年版图式
' 测量员 = 陈  兵
' 绘图员 = 林  强
' 检查员 = 方飞亮

KeyStr = "测绘单位,测绘时间及方式,坐标系,高程体系,图示,测量员,绘图员,检查员"

Sub OnClick()
    SSProcess.ClearInputParameter
    
    KeyArr = Split(KeyStr,",", - 1,1)
    
    For i = 0 To UBound(KeyArr)
        SSProcess.AddInputParameter KeyArr(i) , SSProcess.ReadEpsIni("验线图廓信息", KeyArr(i) ,"") , 0 , "" , ""
    Next 'i
    
    ShowBoolen = SSProcess.ShowInputParameterDlg ("图廓信息录入")
    
    For i = 0 To UBound(KeyArr)
        SSProcess.WriteEpsIni "验线图廓信息", KeyArr(i) ,SSProcess.GetInputParameter(KeyArr(i))
    Next 'i
    
End Sub' OnClick