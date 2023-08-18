
Sub OnClick()
    JdAreaInfo
End Sub' OnClick

Function JdAreaInfo()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9420025"
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount
    For i = 0 To SelCount - 1
        BZStr = SSProcess.GetSelGeoValue(i,"[JDMC]")
        ID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        SSProcess.GetObjectFocusPoint ID,x,y
        DrawNote BZStr,x,y
    Next 'i
End Function' JdAreaInfo

Function DrawNote(ByVal BZStr,ByVal X,ByVal Y)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontString", BZStr
    SSProcess.SetNewObjValue "SSObj_FontClass", "JD001"
    SSProcess.SetNewObjValue "SSObj_LayerName", "»ùµ××¢¼Ç"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", 1000
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(0,255,191)"
    SSProcess.SetNewObjValue "SSObj_FontHeight", 1000
    SSProcess.SetNewObjValue "SSObj_FontDirection", 0
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function' DrawNote
