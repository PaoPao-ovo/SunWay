
'====================================================设置图廓的编码================================================

'图廓编码
Dim TkCode
'TkCode = "9310093,9510031,9131013,9420033,9420034,9420035,9420036,9420037,9460093,9699003,9699013,9470105,9430093,9320053"
TkCode = "9420037"
'图廓名称
Dim MapName
MapName = "建设工程实地放线平面图,勘测定界图,宗地图,竣工图,竣工测绘总平面图,竣工规划复核图,用地复核图,基地总平面布局核实测量平面图,停车库核实测量平面图,综合管线竣工图,专业管线竣工图,绿地竣工地形图,总平面测量略图,土地界址点核验平面图"

'================================================文件路径操作对象==================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Word操作对象
Dim Global_Word
Set Global_Word = CreateObject ("asposewordscom.asposewordshelper")

'====================================================功能入口======================================================

'总入口
Sub OnClick()
    InfoWindow SYStr
    TkCodeList = Split(TkCode,",")
    'SSProcess.MapCallBackFunction "OutputMsg", "正在输出PDF......请稍候",0
    For i = 0 To UBound(TkCodeList)
        TkCode = TkCodeList(i)
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "=", TkCode
        SSProcess.SelectFilter
        TkCount = SSProcess.GetSelGeoCount
        MapNameList = Split(MapName,",")
        If TkCount > 1 Then
            MsgBox "图廓不唯一，放弃输出！"
            Exit Sub
        ElseIf TkCount = 1 Then
            TkId = SSProcess.GetSelGeoValue(0,"SSObj_ID")
        End If
        
        PrintPDF TkId,"Foxit PDF Printer","Foxit PDF Printer Driver","FOXIT_PDF:","1",SYStr,DeleteMark,20
        
        DelNote DeleteMark
        
        MsgBox "输出完成"
    Next
    
End Sub' OnClick

Function InfoWindow(ByRef SYStr)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "测绘单位名称" , "宁波市测绘和遥感技术研究院" , 0 , "宁波市测绘和遥感技术研究院" , ""
    Result = SSProcess.ShowInputParameterDlg ("测绘单位名称")
    If Result = 1 Then
        SYStr = SSProcess.GetInputParameter ("测绘单位名称")
        If SYStr <> "" Then
            SYStr = SSProcess.GetInputParameter ("测绘单位名称")
        Else
            SYStr = "宁波市测绘和遥感技术研究院"
        End If
    Else
        SYStr = "宁波市测绘和遥感技术研究院"
    End If
    
End Function' InfoWindow

Function PrintPDF(ByVal TkId,ByVal Printer,ByVal PrinterDriver,ByVal PrinterPort,ByVal PaperCount,ByVal NoteStr,ByRef DeleteMark,ByVal OffsetVal)
    
    '图廓坐标点
    SSProcess.GetObjectPoint TkId,0,X0,Y0,Z0,Ptype0,Name0
    SSProcess.GetObjectPoint TkId,1,X1,Y1,Z1,Ptype1,Name1
    SSProcess.GetObjectPoint TkId,2,X2,Y2,Z2,Ptype2,Name2
    SSProcess.GetObjectPoint TkId,3,X3,Y3,Z3,Ptype3,Name3
    
    CenterX = (X0 + X2) / 2
    CenterY = (Y0 + Y2) / 2
    GetAngle X0,Y0,X2,Y2,Angle,Length
    GetWH Length,Width,Height,NoteStr
    DrawNote Angle,CenterX,CenterY,Width,Height,DeleteMark,NoteStr
    PrintPaper = SSProcess.GetObjectAttr( TkId, "[DaYZZ]")
    If PrintPaper = "" Then
        PrintPaper = "A4纵向"
    End If
    
    If PrintPaper = "A4横向" Then
        Orientation = 2 '横纵向
        PaperW = 210 '纸宽
        PaperH = 297 '纸高
    ElseIf PrintPaper = "A4纵向" Then
        Orientation = 1
        PaperW = 210
        PaperH = 297
    ElseIf PrintPaper = "A3纵向" Then
        Orientation = 1
        PaperW = 297
        PaperH = 420
    ElseIf PrintPaper = "A3横向" Then
        Orientation = 2
        PaperW = 297
        PaperH = 420
    ElseIf PrintPaper = "A2横向" Then
        Orientation = 2
        PaperW = 420
        PaperH = 594
    ElseIf PrintPaper = "A2纵向" Then
        Orientation = 1
        PaperW = 420
        PaperH = 594
    ElseIf PrintPaper = "A1横向" Then
        Orientation = 2
        PaperW = 594
        PaperH = 841
    ElseIf PrintPaper = "A1纵向" Then
        Orientation = 1
        PaperW = 594
        PaperH = 841
    End If
    
    PrintScale = SSProcess.GetObjectAttr(TkId,"[DaYBL]")
    
    If IsNumeric(PrintScale) = False Then
        PrintScale = 500
    Else
        PrintScale = CInt(PrintScale)
    End If
    
    '打印范围线（长宽为Width和Height的长方形）
    Width = X2 - X0
    Height = Y2 - Y0
    
    '打印起点（长方形左上角点，左上向右下画，系统设置后无需调整大小和偏移）
    X = CDbl(X3)
    Y = CDbl(Y3)
    
    
    ModifyPaper Orientation,Width,Height,PaperW,PaperH,LeftMargin,TopMargin,PrintPaper
    
    SetPrinter Printer,PrinterDriver,PrinterPort,PaperCount,PaperW,PaperH,Orientation,TopMargin,LeftMargin,PrintScale
    
    Set WinShellNetwork = CreateObject("wscript.network")
    Hostname = WinShellNetwork.username
    Set Winshell = CreateObject("wscript.shell")
    ComputerStr = "."
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & ComputerStr & "\root\default:StdRegProv")
    EpsTempPath = SSProcess.GetSysPathName(5)
    '设置缺省输出目录
    StringValuesArr = Array(EpsTempPath)
    Const HKEY_CURRENT_USER =  &H80000001
    KeyRootStr = HKEY_CURRENT_USER
    KeyPathStr = "Software\Foxit Software\PDF Creator\" & Hostname & "\"
    oReg.SetMultiStringValue KeyRootStr, KeyPathStr, "Folder", StringValuesArr
    '使用缺省路径
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "UseDefFileName",1
    '是否直接覆盖
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "Overwrite", 1
    '是否输出后自动打开
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "OpenFile", 0
    '是否透明显示
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "GDIPunt", 0
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "PDFVersion", "14"
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "PDFA1B","0"
    Angle = 0
    
    SSParameter.SetParameterSTR "BeforePrintMap", "FrameCoord", ""
    
    SSProcess.PrintMapByCoord X,Y,Width,Height,Angle
    
    SSProcess.MapMethod "SetPrintLineWidthDelta","0"
    
End Function' PrintPDF

'设置打印机参数
Function SetPrinter(ByVal Printer,ByVal PrinterDriver,ByVal PrinterPort,ByVal PaperCount,ByVal PaperW,ByVal PaperH,ByVal Orientation,ByVal TopMargin,ByVal LeftMargin,ByVal Scale)
    
    'Ini字段
    Dim KeyName(9)
    
    '字段原始值
    Dim OldValue(9)
    
    ParameterStr = "PrintParameter"
    keyName(0) = "PrinterName"
    SetNewPrinterParameter  ParameterStr, KeyName(0), Printer, OldValue(0)
    keyName(1) = "PortName"
    SetNewPrinterParameter  ParameterStr, KeyName(1),PrinterPort, OldValue(1)
    keyName(2) = "Driver"
    SetNewPrinterParameter  ParameterStr, KeyName(2),PrinterDriver, OldValue(2)
    keyName(3) = "PaperW"
    SetNewPrinterParameter  ParameterStr, KeyName(3),PaperW, OldValue(3)
    keyName(4) = "PaperH"
    SetNewPrinterParameter  ParameterStr, KeyName(4), PaperH, OldValue(4)
    keyName(5) = "Orientation"
    SetNewPrinterParameter  ParameterStr, KeyName(5), Orientation, OldValue(5)
    keyName(6) = "TopMargin"
    SetNewPrinterParameter  ParameterStr, KeyName(6), TopMargin, OldValue(6)
    keyName(7) = "LeftMargin"
    SetNewPrinterParameter  ParameterStr, KeyName(7),LeftMargin, OldValue(7)
    keyName(8) = "Scale"
    SetNewPrinterParameter  ParameterStr, KeyName(8),Scale, OldValue(8)
    keyName(9) = "PrintOutputMode"
    SetNewPrinterParameter  ParameterStr, KeyName(9),"0", OldValue(9)
    '刷新打印内存参数
    SSProcess.MapMethod "ReadPrinterSetting", Parameters
End Function' SetPrinter

'调整纸张大小和截图位置
Function ModifyPaper(ByVal Orientation,ByVal TkW,ByVal TkH,ByRef PaperW,ByRef PaperH,ByRef LeftMargin,ByRef TopMargin,ByVal PrintPaper)
    LenScale = 1 + 1 / 3 '图上距离和纸张长度的比例系数（图上10表示纸上10*LenScale）
    If Orientation = 1 Then
        LeftMargin = 20 * LenScale
        PaperW = PaperW + 40
        TopMargin = 45 * LenScale
        PaperH = PaperH + 90
    ElseIf Orientation = 2 Then
        LeftMargin = 20 * LenScale
        PaperH = PaperH + 40
        TopMargin = 45 * LenScale
        PaperW = PaperW + 90
    End If
End Function' ModifyPaper

'写Ini
Function SetNewPrinterParameter(ByVal ParameterStr,ByVal KeyStr,ByVal ValueStr,ByRef OldValue)
    OldValue = SSProcess.ReadEpsIni(ParameterStr,KeyStr ,"")
    SSProcess.WriteEpsIni ParameterStr,KeyStr,ValueStr
End Function

'计算角度
Function GetAngle(ByVal X0,ByVal Y0,ByVal X2,ByVal Y2,ByRef Angle,ByRef Length)
    SSProcess.XYSA X0,Y0,X2,Y2,Length,Angle,0
    Angle = SSProcess.RadianToDeg(Angle)
End Function' GetAngle

'获取字宽字高
Function GetWH(ByVal Length,ByRef Width,ByRef Height,ByVal NoteStr)
    WordCount = Len(NoteStr)
    WordXs = 111
    StringLength = Length - 48 * 2
    SingleLength = CInt(StringLength / WordCount)
    Width = WordXs * SingleLength
    Height = WordXs * SingleLength
End Function' GetWH

'绘制注记
Function DrawNote(ByVal Angle,ByVal CenterX,ByVal CenterY,ByVal Width,ByVal Height,ByRef DeleteMark,ByVal NoteStr)
    DeleteMark = 1
    SSProcess.CreateNewObj 3
    SSProcess.AddNewObjPoint CenterX,CenterY,0,0,""
    SSProcess.SetNewObjValue "SSObj_FontString", NoteStr
    SSProcess.SetNewObjValue "SSObj_Color",  "RGB(186, 183, 183)"
    SSProcess.SetNewObjValue "SSObj_DataMark", DeleteMark
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontStringAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontWordAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function' DrawNote

'删除注记
Function DelNote(ByVal DeleteMark)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_DataMark", "==", DeleteMark
    SSProcess.SetSelectCondition "SSObj_Type", "==", "NOTE"
    SSProcess.SelectFilter
    NotecCount = SSProcess.GetSelNoteCount
    If NotecCount > 0 Then
        For i = 0 To NotecCount - 1
            SSProcess.DelSelNote i
        Next
    End If
End Function' DelNote