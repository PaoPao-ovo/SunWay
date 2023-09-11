
Sub OnClick()
    XF_Info
End Sub' OnClick

Function XF_Info()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9430013"
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount
    For i = 0 To SelCount - 1
        ID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        GetLineLength ID,Length1,Length2,CneterX1,CneterY1,Angle1,CneterX2,CneterY2,Angle2
        DrawNote Length1,CneterX1,CneterY1,Angle1
        DrawNote Length2,CneterX2,CneterY2,Angle2
    Next 'i
End Function' JdAreaInfo

'获取第一边和第二边的边长
Function GetLineLength(ByVal XFId,ByRef Length1,ByRef Length2,ByRef NoteX1,ByRef NoteY1,ByRef Angle1,ByRef NoteX2,ByRef NoteY2,ByRef Angle2)
    
    SSProcess.GetObjectPoint XFId,0,X0,Y0,Z0,Ptype0,Name0
    SSProcess.GetObjectPoint XFId,1,X1,Y1,Z1,Ptype1,Name1
    SSProcess.GetObjectPoint XFId,2,X2,Y2,Z2,Ptype2,Name2
    SSProcess.GetObjectPoint XFId,3,X3,Y3,Z3,Ptype3,Name3
    
    CneterX1 = (X1 + X0) / 2
    CneterY1 = (Y1 + Y0) / 2
    
    SSProcess.XYSA X0,Y0,X1,Y1,Length1,Angle1,0
    Length1 = CStr(Round(Length1,2)) & "M"
    Sin1 = Sin(Angle1)
    COS1 = Cos(Angle1)
    Angle1 = SSProcess.RadianToDeg(Angle1)
    
    '确定象限
    If Angle1 >= 0 And Angle1 <= 90 Then
        Quadrant1 = 1
    ElseIf Angle1 > 90 And Angle1 <= 180 Then
        Quadrant1 = 2
    ElseIf Angle1 > 180 And  Angle1 <= 270 Then
        Quadrant1 = 3
    Else
        Quadrant1 = 4
    End If
    
    CneterX2 = (X2 + X1) / 2
    CneterY2 = (Y2 + Y1) / 2
    
    SSProcess.XYSA X1,Y1,X2,Y2,Length2,Angle2,0
    Length2 = CStr(Round(Length2,2)) & "M"
    Sin2 = Sin(Angle2)
    COS2 = Cos(Angle2)
    Angle2 = SSProcess.RadianToDeg(Angle2)
    
    '确定象限
    If Angle2 >= 0 And Angle2 <= 90 Then
        Quadrant2 = 1
    ElseIf Angle2 > 90 And Angle2 <= 180 Then
        Quadrant2 = 2
    ElseIf Angle2 > 180 And  Angle2 <= 270 Then
        Quadrant2 = 3
    Else
        Quadrant2 = 4
    End If

    If Quadrant1 = 1 Then
        NoteY2 = CneterY2 - Abs(Sin1 * 2)
        NoteX2 = CneterX2 - Abs(COS1 * 2)
    ElseIf Quadrant1 = 2 Then
        NoteY2 = CneterY2 - Abs(Sin1 * 2)
        NoteX2 = CneterX2 + Abs(COS1 * 2)
    ElseIf Quadrant1 = 3 Then
        NoteY2 = CneterY2 + Abs(Sin1 * 2)
        NoteX2 = CneterX2 + Abs(COS1 * 2)
    Else
        NoteY2 = CneterY2 + Abs(Sin1 * 2)
        NoteX2 = CneterX2 - Abs(COS1 * 2)
    End If
    
    If Quadrant2 = 1 Then
        NoteY1 = CneterY1 + Abs(Sin2 * 2)
        NoteX1 = CneterX1 + Abs(COS2 * 2)
    ElseIf Quadrant2 = 2 Then
        NoteY1 = CneterY1 + Abs(Sin2 * 2)
        NoteX1 = CneterX1 - Abs(COS2 * 2)
    ElseIf Quadrant2 = 3 Then
        NoteY1 = CneterY1 - Abs(Sin2 * 2)
        NoteX1 = CneterX1 - Abs(COS2 * 2)
    Else
        NoteY1 = CneterY1 - Abs(Sin2 * 2)
        NoteX1 = CneterX1 + Abs(COS2 * 2)
    End If
    
End Function' GetLineLength

Function DrawNote(ByVal BZStr,ByVal X,ByVal Y,ByVal Angle)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", BZStr
    SSProcess.SetNewObjValue "SSObj_FontWordAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontStringAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", 200
    SSProcess.SetNewObjValue "SSObj_FontHeight", 160
    SSProcess.SetNewObjValue "SSObj_FontDirection", 0
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function' DrawNote
