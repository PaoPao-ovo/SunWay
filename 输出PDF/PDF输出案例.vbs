
'====================================================设置图廓的编码================================================

'图廓编码
Dim TkCode
TkCode = "9420033"

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
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=", TkCode
    SSProcess.SelectFilter
    TkCount = SSProcess.GetSelGeoCount
    If TkCount <> 1 Then
        MsgBox "图廓不唯一，放弃输出！"
        Exit Sub
    Else
        TkId = SSProcess.GetSelGeoValue(0,"SSObj_ID")
        MsgBox TkId
    End If
    
    SSProcess.MapCallBackFunction "OutputMsg", "正在输出PDF......请稍候",0
    
    PrintPDF TkId,"Foxit PDF Printer","Foxit PDF Printer Driver","FOXIT_PDF:","1"
End Sub' OnClick

Function PrintPDF(ByVal TkId,ByVal Printer,ByVal PrinterDriver,ByVal PrinterPort,ByVal PaperCount)
    '图廓坐标点
    SSProcess.GetObjectPoint TkId,0,X0,Y0,Z0,Ptype0,Name0
    SSProcess.GetObjectPoint TkId,1,X1,Y1,Z1,Ptype1,Name1
    SSProcess.GetObjectPoint TkId,2,X2,Y2,Z2,Ptype2,Name2
    SSProcess.GetObjectPoint TkId,3,X3,Y3,Z3,Ptype3,Name3
    
    CenterX = (X0 + X2) / 2
    CenterY = (Y0 + Y2) / 2
    GetAngle X0,Y0,X2,Y2,Angle,Length
    GetWH Length,Width,Height
    DrawNote Angle,CenterX,CenterY,Width,Height,DeleteMark
    
    PrintPaper = SSProcess.GetObjectAttr( TkId, "[DaYZZ]")
    If PrintPaper = "" Then
        PrintPaper = "A4纵向"
    End If
    
    If PrintPaper = "A4横向" Then
        Orientation = 2
        PaperW = 210
        PaperH = 297
        '设置纸张的大小
        LeftMargin = 10
        RightMargin = 10
        TopMargin = 20
        BottomMargin = 10
        H = 17.1
        W = 24.2
        width = 16.345 * W
        height = 16.345 * H
        
    ElseIf PrintPaper = "A4纵向" Then
        Orientation = 1
        '设置宽高
        PaperW = 210
        PaperH = 297
        '设置纸张的大小
        LeftMargin = 10
        RightMargin = 20
        TopMargin = 20
        BottomMargin = 10
        H = 26.8
        W = 21.8
        width = 50.345 * W
        height = 50.345 * H
        
    ElseIf PrintPaper = "A3纵向" Then
        Orientation = 1
        PaperW = 297
        PaperH = 420
        '设置纸张的大小
        LeftMargin = 10
        RightMargin = 10
        TopMargin = 20
        BottomMargin = 10
        H = 37.2
        W = 26.3
        width = 28.345 * W
        height = 28.345 * H
        
    ElseIf PrintPaper = "A3横向" Then
        Orientation = 2
        PaperW = 297
        PaperH = 420
        '设置纸张的大小
        LeftMargin = 10
        RightMargin = 10
        TopMargin = 20
        BottomMargin = 10
        
        'H=25.8: W=36.5
        'width = 28.345*W
        'height = 28.345*H
    End If
    
    
    ' Scale = SSProcess.GetSelGeoValue(0,"[DaYBL]")
    PrintScale = SSProcess.GetObjectAttr( TkId, "[DaYBL]")
    
    If IsNumeric(PrintScale) = False Then
        PrintScale = 500
    Else
        PrintScale = CInt(PrintScale)
    End If
    
    width = X2 - X0 + 10
    height = Y2 - Y0 + 20
    
    X = CDbl(X3)
    
    Y = CDbl(Y3)
    
    
    SetPrinter Printer,PrinterDriver,PrinterPort,PaperCount,PaperW,PaperH,Orientation,TopMargin,LeftMargin,PrintScale,X,Y
    
    Set WinShellNetwork = CreateObject("wscript.network")
    Hostname = WinShellNetwork.username
    Set Winshell = CreateObject("wscript.shell")
    ComputerStr = "."
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & ComputerStr & "\root\default:StdRegProv")
    EpsTempPath = SSProcess.GetSysPathName(5)
    '设置缺省输出目录
    StringValuesArr = Array(EpsTempPath)
    Const HKEY_CURRENT_USER =  & H80000001
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
Function SetPrinter(ByVal Printer,ByVal PrinterDriver,ByVal PrinterPort,ByVal PaperCount,ByVal PaperW,ByVal PaperH,ByVal Orientation,ByVal TopMargin,ByVal LeftMargin,ByVal Scale,ByVal X,ByVal Y)
    
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
Function GetWH(ByVal Length,ByRef Width,ByRef Height)
    WordXs = 222
    StringLength = Length - 48 * 2
    SingleLength = CInt(StringLength / 13)
    Width = WordXs * SingleLength
    Height = WordXs * SingleLength
End Function' GetWH

'绘制注记
Function DrawNote(ByVal Angle,ByVal CenterX,ByVal CenterY,ByVal Width,ByVal Height,ByRef DeleteMark)
    DeleteMark = 1
    SSProcess.CreateNewObj 3
    SSProcess.AddNewObjPoint CenterX,CenterY,0,0,""
    SSProcess.SetNewObjValue "SSObj_FontString", "宁波市测绘和遥感技术研究院"
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