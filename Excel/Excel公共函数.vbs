
Function IsFileExist(File)
    Dim fso
    IsFileExist = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.fileExists(File) Then
        IsFileExist = True
    End If
    Set fso = Nothing
End Function

Function DeleteFile(File)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.deleteFile File
    Set fso = Nothing
End Function
Function InitExcel(ByVal xlsFileName)
    Set xlsApp = CreateObject("Excel.Application")
    Set xlsBook = xlsApp.WorkBooks.Open(xlsFileName)
End Function

'设置活动Sheet
Function SetActiveSheet( sheetName )
    Set xlsSheet = xlsBook.Worksheets(sheetName)
    xlsSheet.Activate
End Function

Function SetActiveSheetByIndex( sheetindex )
    Set xlsSheet = xlsBook.Worksheets(sheetindex)
    xlsSheet.Activate
End Function

'复制
Function CopySheet(srcSheetName, tagSheetName)
    Dim xlsSheet0, xlsSheet1
    Set xlsSheet0 = xlsBook.Worksheets(srcSheetName)
    xlsSheet0.Select
    xlsSheet0.Copy xlsSheet0
    Set xlsSheet1 = xlsBook.ActiveSheet
    xlsSheet1.Select
    xlsSheet1.Name = tagSheetName
    
End Function

'删除
Function DeleteSheet( sheetName )
    Dim xlsSheet0
    Set xlsSheet0 = xlsBook.Worksheets(sheetName)
    xlsSheet0.Select
    xlsSheet0.Delete
End Function

'拷贝表格
Function CopySheetTable( beginRow, count, pageRowCount )
    For i = 0 To count - 1
        xlsSheet.Rows( beginRow & ":" & (beginRow + pageRowCount - 1) ).Select
        xlsApp.Selection.Copy
        row = (i + 1) * pageRowCount + beginRow
        rows = row & ":" & row
        xlsSheet.Rows(rows).Select
        xlsSheet.Paste
    Next
End Function

'删除表格
Function DelSheetTable( beginRow, pageRowCount )
    xlsSheet.Rows( beginRow & ":" & (beginRow + pageRowCount - 1) ).Select
    xlsApp.Selection.Delete
End Function

'填充单元格
Function SetCellValue(row, col, value)
    xlsSheet.Cells(row, col) = value
End Function

'合并单元格式
Function MergeCell(row, beginCol, EndCol)
    xlsSheet.Range(xlsSheet.Cells(row, beginCol), xlsSheet.Cells(row, EndCol)).Merge
End Function
