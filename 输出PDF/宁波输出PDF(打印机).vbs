
'====================================================����ͼ���ı���================================================

'ͼ������
Dim TkCode
'TkCode = "9310093,9510031,9131013,9420033,9420034,9420035,9420036,9420037,9460093,9699003,9699013,9470105,9430093,9320053"
TkCode = "9420037"
'ͼ������
Dim MapName
MapName = "���蹤��ʵ�ط���ƽ��ͼ,���ⶨ��ͼ,�ڵ�ͼ,����ͼ,���������ƽ��ͼ,�����滮����ͼ,�õظ���ͼ,������ƽ�沼�ֺ�ʵ����ƽ��ͼ,ͣ�����ʵ����ƽ��ͼ,�ۺϹ��߿���ͼ,רҵ���߿���ͼ,�̵ؿ�������ͼ,��ƽ�������ͼ,���ؽ�ַ�����ƽ��ͼ"

'================================================�ļ�·����������==================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Word��������
Dim Global_Word
Set Global_Word = CreateObject ("asposewordscom.asposewordshelper")

'====================================================�������======================================================

'�����
Sub OnClick()
    InfoWindow SYStr
    TkCodeList = Split(TkCode,",")
    'SSProcess.MapCallBackFunction "OutputMsg", "�������PDF......���Ժ�",0
    For i = 0 To UBound(TkCodeList)
        TkCode = TkCodeList(i)
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "=", TkCode
        SSProcess.SelectFilter
        TkCount = SSProcess.GetSelGeoCount
        MapNameList = Split(MapName,",")
        If TkCount > 1 Then
            MsgBox "ͼ����Ψһ�����������"
            Exit Sub
        ElseIf TkCount = 1 Then
            TkId = SSProcess.GetSelGeoValue(0,"SSObj_ID")
        End If
        
        PrintPDF TkId,"Foxit PDF Printer","Foxit PDF Printer Driver","FOXIT_PDF:","1",SYStr,DeleteMark,20
        
        DelNote DeleteMark
        
        MsgBox "������"
    Next
    
End Sub' OnClick

Function InfoWindow(ByRef SYStr)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "��浥λ����" , "�����в���ң�м����о�Ժ" , 0 , "�����в���ң�м����о�Ժ" , ""
    Result = SSProcess.ShowInputParameterDlg ("��浥λ����")
    If Result = 1 Then
        SYStr = SSProcess.GetInputParameter ("��浥λ����")
        If SYStr <> "" Then
            SYStr = SSProcess.GetInputParameter ("��浥λ����")
        Else
            SYStr = "�����в���ң�м����о�Ժ"
        End If
    Else
        SYStr = "�����в���ң�м����о�Ժ"
    End If
    
End Function' InfoWindow

Function PrintPDF(ByVal TkId,ByVal Printer,ByVal PrinterDriver,ByVal PrinterPort,ByVal PaperCount,ByVal NoteStr,ByRef DeleteMark,ByVal OffsetVal)
    
    'ͼ�������
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
        PrintPaper = "A4����"
    End If
    
    If PrintPaper = "A4����" Then
        Orientation = 2 '������
        PaperW = 210 'ֽ��
        PaperH = 297 'ֽ��
    ElseIf PrintPaper = "A4����" Then
        Orientation = 1
        PaperW = 210
        PaperH = 297
    ElseIf PrintPaper = "A3����" Then
        Orientation = 1
        PaperW = 297
        PaperH = 420
    ElseIf PrintPaper = "A3����" Then
        Orientation = 2
        PaperW = 297
        PaperH = 420
    ElseIf PrintPaper = "A2����" Then
        Orientation = 2
        PaperW = 420
        PaperH = 594
    ElseIf PrintPaper = "A2����" Then
        Orientation = 1
        PaperW = 420
        PaperH = 594
    ElseIf PrintPaper = "A1����" Then
        Orientation = 2
        PaperW = 594
        PaperH = 841
    ElseIf PrintPaper = "A1����" Then
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
    
    '��ӡ��Χ�ߣ�����ΪWidth��Height�ĳ����Σ�
    Width = X2 - X0
    Height = Y2 - Y0
    
    '��ӡ��㣨���������Ͻǵ㣬���������»���ϵͳ���ú����������С��ƫ�ƣ�
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
    '����ȱʡ���Ŀ¼
    StringValuesArr = Array(EpsTempPath)
    Const HKEY_CURRENT_USER =  &H80000001
    KeyRootStr = HKEY_CURRENT_USER
    KeyPathStr = "Software\Foxit Software\PDF Creator\" & Hostname & "\"
    oReg.SetMultiStringValue KeyRootStr, KeyPathStr, "Folder", StringValuesArr
    'ʹ��ȱʡ·��
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "UseDefFileName",1
    '�Ƿ�ֱ�Ӹ���
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "Overwrite", 1
    '�Ƿ�������Զ���
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "OpenFile", 0
    '�Ƿ�͸����ʾ
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "GDIPunt", 0
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "PDFVersion", "14"
    oReg.SetDWORDValue KeyRootStr, KeyPathStr, "PDFA1B","0"
    Angle = 0
    
    SSParameter.SetParameterSTR "BeforePrintMap", "FrameCoord", ""
    
    SSProcess.PrintMapByCoord X,Y,Width,Height,Angle
    
    SSProcess.MapMethod "SetPrintLineWidthDelta","0"
    
End Function' PrintPDF

'���ô�ӡ������
Function SetPrinter(ByVal Printer,ByVal PrinterDriver,ByVal PrinterPort,ByVal PaperCount,ByVal PaperW,ByVal PaperH,ByVal Orientation,ByVal TopMargin,ByVal LeftMargin,ByVal Scale)
    
    'Ini�ֶ�
    Dim KeyName(9)
    
    '�ֶ�ԭʼֵ
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
    'ˢ�´�ӡ�ڴ����
    SSProcess.MapMethod "ReadPrinterSetting", Parameters
End Function' SetPrinter

'����ֽ�Ŵ�С�ͽ�ͼλ��
Function ModifyPaper(ByVal Orientation,ByVal TkW,ByVal TkH,ByRef PaperW,ByRef PaperH,ByRef LeftMargin,ByRef TopMargin,ByVal PrintPaper)
    LenScale = 1 + 1 / 3 'ͼ�Ͼ����ֽ�ų��ȵı���ϵ����ͼ��10��ʾֽ��10*LenScale��
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

'дIni
Function SetNewPrinterParameter(ByVal ParameterStr,ByVal KeyStr,ByVal ValueStr,ByRef OldValue)
    OldValue = SSProcess.ReadEpsIni(ParameterStr,KeyStr ,"")
    SSProcess.WriteEpsIni ParameterStr,KeyStr,ValueStr
End Function

'����Ƕ�
Function GetAngle(ByVal X0,ByVal Y0,ByVal X2,ByVal Y2,ByRef Angle,ByRef Length)
    SSProcess.XYSA X0,Y0,X2,Y2,Length,Angle,0
    Angle = SSProcess.RadianToDeg(Angle)
End Function' GetAngle

'��ȡ�ֿ��ָ�
Function GetWH(ByVal Length,ByRef Width,ByRef Height,ByVal NoteStr)
    WordCount = Len(NoteStr)
    WordXs = 111
    StringLength = Length - 48 * 2
    SingleLength = CInt(StringLength / WordCount)
    Width = WordXs * SingleLength
    Height = WordXs * SingleLength
End Function' GetWH

'����ע��
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

'ɾ��ע��
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