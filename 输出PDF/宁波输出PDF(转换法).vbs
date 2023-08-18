
'====================================================����ͼ���ı���================================================

'ͼ������
Dim TkCode
TkCode = "9450073,9310093,9510031,9131013,9420033,9420034,9420035,9420036,9420037,9460093,9699003,9699013,9470105,9430093,9320053"
MapName = "�˷�,���蹤��ʵ�ط���ƽ��ͼ,���ⶨ��ͼ,�ڵ�ͼ,����ͼ,���������ƽ��ͼ,�����滮����ͼ,�õظ���ͼ,������ƽ�沼�ֺ�ʵ����ƽ��ͼ,ͣ�����ʵ����ƽ��ͼ,�ۺϹ��߿���ͼ,רҵ���߿���ͼ,�̵ؿ�������ͼ,��ƽ�������ͼ,���ؽ�ַ�����ƽ��ͼ"

'================================================�ļ�·����������==================================================

'·����������
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Word��������
Dim Global_Word
Set Global_Word = CreateObject ("asposewordscom.asposewordshelper")

'===============================================���Word����======================================================

'Word�ļ�����
Dim WordFileName
WordFileName = "�ɹ�ͼ��ˮӡ.doc"

'====================================================�������======================================================

'�����
Sub OnClick()
    
    InfoWindow SYStr
    
    If  TypeName (Global_Word) <> "AsposeWordsHelper" Then
        MsgBox "����ע��Aspose.Word���"
        Exit Sub
    End If
    TkCodeList = Split(TkCode,",")
    MapNameList = Split(MapName,",")
    For i = 0 To UBound(TkCodeList)
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "=", TkCodeList(i)
        SSProcess.SelectFilter
        TkCount = SSProcess.GetSelGeoCount
        If TkCount > 1 Then
            MsgBox "ͼ����Ψһ�����������"
            Exit Sub
        ElseIf TkCount = 1 Then
            For i1 = 0 To TkCount - 1
                TkId = SSProcess.GetSelGeoValue(i1,"SSObj_ID")
            Next
            
            '����һ���յ�Word�ļ�
            FilePath = SSProcess.GetSysPathName (5) & WordFileName
            Set WordFile = FileSysObj.CreateTextFile(FilePath,True)
            WordFile.Close
            
            'Word���ΪPDF
            FilePDFPath = SSProcess.GetSysPathName (5) & MapNameList(i) & ".pdf"
            
            '����ͼƬ
            InsertImage TkId,FilePath,FilePDFPath,PrintPaper,DeleteMark,SYStr
            
            'ɾ��ԭ�е�Word�ļ�
            Set WordFile = FileSysObj.GetFile(FilePath)
            WordFile.Delete
            
            'ɾ��ע��
            DelNote DeleteMark
            
        End If
    Next
    MsgBox "���"
End Sub' OnClick

Function InfoWindow(ByRef SYStr)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "��浥λ����" , "�����в���ң�м����о�Ժ" , 0 , "�����в���ң�м����о�Ժ" , ""
    Result = SSProcess.ShowInputParameterDlg ("��浥λ����")
    If Result = 1 Then
        SYStr = SSProcess.GetInputParameter ("��浥λ����")
    Else
        SYStr = "�����в���ң�м����о�Ժ"
    End If
End Function' InfoWindow

'����ˮӡͼƬ
Function InsertSignature()
    FolderPath = SSProcess.GetSysPathName (0) & "\ǩ��\"
    names = "ˮӡ"
    nameList = Split(names,",")
    For i = 0 To UBound(nameList)
        name = nameList(i)
        imageFile = FolderPath & name & ".png"
        If name = "ˮӡ" Then
            If IsFileExists(imageFile) = True Then Global_Word.SetImgWatermark imageFile, 400, 400,0
        Else
            Global_Word.MoveToBookmark name
            If IsFileExists(imageFile) = True Then Global_Word.InsertImageEx imageFile,  0, 250, 0, 390, 150, 150,3, 0
        End If
    Next
End Function

'//�ж��ļ��Ƿ����
Function IsFileExists(ByVal filespec)
    IsFileExists = False
    If (FileSysObj.FileExists(filespec)) = True Then
        IsFileExists = True
    End If
End Function

'���ͼƬ
Function InsertImage(ByVal TkId,ByVal FilePath,ByVal FilePDFPath,ByRef PrintPaper,ByRef DeleteMark,ByVal SYStr)
    
    'ɾ��ͼƬ����
    DeleteAllImage
    
    '��Word
    Global_Word.OpenDocument FilePath
    
    '��ȡͼ����Ϣ
    PrintScale = SSProcess.GetSelGeoValue(0,"[DaYBL]")
    LeftMargin = SSProcess.GetSelGeoValue(0,"[ZuoBJ]")
    UpMargin = SSProcess.GetSelGeoValue(0,"[ShangBJ]")
    PrintPaper = SSProcess.GetSelGeoValue(0,"[DaYZZ]")
    
    '��ȡʧ������Ĭ��ֵ
    If IsNumeric(PrintScale) = False Then PrintScale = 500
    If IsNumeric(LeftMargin) = False Then LeftBoundary = 10
    If IsNumeric(UpMargin) = False Then UpBoundary = 10
    
    'ֽ�ŵĿ��
    PaperWidth = SSProcess.GetSelGeoValue(0,"[ZhiK]")
    PaperHeight = SSProcess.GetSelGeoValue(0,"[ZhiG]")
    
    H = 0
    W = 0
    
    'ֽ������
    If PrintPaper = "" Then
        PrintPaper = "A4����"
    End If
    
    '����ֽ������
    If InStr(PrintPaper,"A4����") > 0 Then
        BaseHeith = 70
        BaseWidth = 70
        PaperWidth = 210
        PaperHeight = 297
        H = 24.9
        W = 18.8
    ElseIf InStr(PrintPaper,"A4����") > 0  Then
        BaseHeith = 105
        BaseWidth = 148.5
        PaperWidth = 297
        PaperHeight = 210
        H = 17.1
        W = 25.6
        ShapeWidth = 26.345 * W
        ShapeHeight = 26.345 * H
    ElseIf InStr(PrintPaper,"A3����") > 0 Then
        BaseHeith = 210
        BaseWidth = 148.5
        PaperWidth = 297
        PaperHeight = 420
        H = 37.2
        W = 26.3
    ElseIf InStr(PrintPaper,"A3����") > 0  Then
        BaseHeith = 148.5
        BaseWidth = 210
        PaperWidth = 420
        PaperHeight = 297
        H = 24.9
        W = 35.2
    ElseIf InStr(PrintPaper,"A2����") > 0  Then
        PaperWidth = 420
        PaperHeight = 594
    ElseIf InStr(PrintPaper,"A2����") > 0 Then
        PaperWidth = 594
        PaperHeight = 420
    ElseIf InStr(PrintPaper,"A1����") > 0  Then
        PaperWidth = 594
        PaperHeight = 841
    ElseIf InStr(PrintPaper,"A1����") > 0 Then
        PaperWidth = 841
        PaperHeight = 594
    Else
        PaperWidth = 297
        PaperHeight = 210
        H = 16.2
        W = 22.9
    End If
    
    If H = 0 Then H = 24.9
    If W = 0 Then W = 17.6
    
    ShapeHeight = 28.345 * H
    ShapeWidth = 28.345 * W
    
    xDist = 1
    yDist = 0.4
    
    'ͼ�������
    SSProcess.GetObjectPoint TkId,0,X0,Y0,Z0,Ptype0,Name0
    SSProcess.GetObjectPoint TkId,1,X1,Y1,Z1,Ptype1,Name1
    SSProcess.GetObjectPoint TkId,2,X2,Y2,Z2,Ptype2,Name2
    
    CenterX = (X0 + X2) / 2
    CenterY = (Y0 + Y2) / 2
    GetAngle X0,Y0,X2,Y2,Angle,Length
    GetWH Length,Width1,Height1,SYStr
    DrawNote Angle,CenterX,CenterY,Width1,Height1,DeleteMark,SYStr
    minX = X0 - 2 * Sqr((X0 - X1) ^ 2 + (Y0 - Y1) ^ 2) / BaseWidth
    minY = Y0 - 4 * Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2) / BaseHeith
    maxX = X2 + 2 * Sqr((X0 - X1) ^ 2 + (Y0 - Y1) ^ 2) / BaseWidth
    maxY = Y2 + 4 * Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2) / BaseHeith
    dpi = 300
    Path = SSProcess.GetSysPathName(7) & "Pictures\"
    strBmpFile = Path & "RFT" & i & ".png"
    SSFunc.DrawToImage minX - 5,minY - 15,maxX + 5,maxY + 15,PaperWidth & "X" & PaperHeight,dpi,strBmpFile '���ָ����Χ�ڵ�ͼ�ε�wmfͼƬ
    'InsertSignature
    SetPaper PrintPaper,strBmpFile,ShapeWidth,ShapeHeight
    Global_Word.SaveEx FilePDFPath
    'Global_Word.SavePdf_2 FilePath
    'Global_Word.SaveEx FilePath
End Function' ExportPDF

'��ȡ�ֿ��ָ�
Function GetWH(ByVal Length,ByRef Width,ByRef Height,ByVal SYStr)
    Count = Len(SYStr)
    WordXs = 222
    StringLength = Length - 48 * 2
    SingleLength = CInt(StringLength / Count)
    Width = WordXs * SingleLength
    Height = WordXs * SingleLength
End Function' GetWH

'����ע��
Function DrawNote(ByVal Angle,ByVal CenterX,ByVal CenterY,ByVal Width,ByVal Height,ByRef DeleteMark,ByVal SYStr)
    DeleteMark = 1
    SSProcess.CreateNewObj 3
    SSProcess.AddNewObjPoint CenterX,CenterY,0,0,""
    SSProcess.SetNewObjValue "SSObj_FontString", SYStr
    SSProcess.SetNewObjValue "SSObj_FontClass", "SY001"
    SSProcess.SetNewObjValue "SSObj_DataMark", DeleteMark
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontStringAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontWordAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontWidth", Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function' DrawNote

'����Ƕ�
Function GetAngle(ByVal X0,ByVal Y0,ByVal X2,ByVal Y2,ByRef Angle,ByRef Length)
    SSProcess.XYSA X0,Y0,X2,Y2,Length,Angle,0
    Angle = SSProcess.RadianToDeg(Angle)
End Function' GetAngle

'����Word��ʽ
Function SetPaper(ByVal PrintPaper,ByVal insertImageFile,ByVal ShapeWidth,ByVal ShapeHeight)
    If PrintPaper = "A4����" Then
        paperSize = 1
        orientation = 2
        pageWidth =  - 1
        pageHeight =  - 1
        H = 17.1
        W = 24.2
        width = 26.345 * W
        height = 26.345 * H
        '����ֽ�ŵĴ�С
        leftMargin = 20'����
        rightMargin = 20
        topMargin = 7
        bottomMargin = 7
    ElseIf PrintPaper = "A4����" Then
        paperSize = 1
        orientation = 1
        pageWidth =  - 1
        pageHeight =  - 1
        '���ÿ��
        H = 26.8
        W = 21.8
        width = 20.245 * W
        height = 10.345 * H
        '����ֽ�ŵĴ�С
        leftMargin = 10'����
        rightMargin = 10
        topMargin = 20
        bottomMargin = 10
    ElseIf PrintPaper = "A3����" Then
        paperSize = 0
        orientation = 1
        pageWidth =  - 1
        pageHeight =  - 1
        H = 37.2
        W = 26.3
        width = 28.345 * W
        height = 28.345 * H
        '����ֽ�ŵĴ�С
        leftMargin = 10'����
        rightMargin = 10
        topMargin = 10
        bottomMargin = 10
    ElseIf PrintPaper = "A3����" Then
        paperSize = 0
        orientation = 2
        pageWidth =  - 1
        pageHeight =  - 1
        H = 25.8
        W = 36.5
        width = 28.345 * W
        height = 28.345 * H
        '����ֽ�ŵĴ�С
        leftMargin = 10'����
        rightMargin = 10
        topMargin = 10
        bottomMargin = 10
    ElseIf PrintPaper = "500*500" Then
        paperSize = 1
        orientation = 1
        pageWidth = 500
        pageHeight = 500
        '���ÿ��
        H = 45.04
        W = 45.01
        width = 30.245 * W
        height = 28.345 * H
        '����ֽ�ŵĴ�С
        leftMargin = 10'����50
        rightMargin = 10
        topMargin = 10
        bottomMargin = 10
    End If
    
    Global_Word.SectionPageSetup 0, paperSize, orientation, pageWidth, pageHeight, leftMargin, rightMargin, topMargin, bottomMargin
    
    horzPos = 0
    left0 = 0
    vertPos = 0
    top0 = 3
    
    wrapType = 0
    '��ת�Ƕ�
    rotation = 0
    Global_Word.InsertImageEx insertImageFile, horzPos, left0, vertPos, top0, ShapeWidth,ShapeHeight, wrapType,rotation
    
End Function' SetPaper

'//��ӡǰ��ɾ��������
Function DeleteAllImage()
    filePath = SSProcess.GetSysPathName (4)
    Dim filenames(10000)
    GetAllFiles filePath,"bmp",filecount,filenames
    For i = 0 To filecount - 1
        projectName = filenames(i)
        If FileExists(projectName) = True Then  FileSysObj.DeleteFile projectName
    Next
End Function

'//��ȡ�����ļ�
Function GetAllFiles(ByRef pathname, ByRef fileExt, ByRef filecount, ByRef filenames())
    Dim folder, file, files, subfolder,folder0, fcount
    If  FileSysObj.FolderExists(pathname) Then
        Set folder = FileSysObj.GetFolder(pathname)
        Set files = folder.Files
        '�����ļ�
        For Each file In files
            extname = FileSysObj.GetExtensionName(file.name)
            If UCase(extname) = UCase(fileExt) Then
                filenames(filecount) = pathname & file.name
                filecount = filecount + 1
            End If
        Next
        '������Ŀ¼
        Set subfolder = folder.SubFolders
        For Each folder0 In subfolder
            GetAllFiles pathname & folder0.name & "\", fileExt, filecount, filenames
        Next
    End If
End Function

'//�ж��ļ��Ƿ����
Function FileExists(ByVal strSrcFilePath)
    res = False
    If (FileSysObj.FileExists(strSrcFilePath)) = True Then res = True
    FileExists = res
End Function

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