
'==================================================图廓和要素的编码配置====================================================

'要素编码
Const CodeStr = "9410001,9410011,9410021,9410031,9410041,9410051,9410061,9410071,9410091,9410101,9410104,9410105;9410021,9410031,9410041,9410051,9410061,9410011,9410001,9430001,9430061,9430051,9430041,9430033,9430023,9430013,9430014,9430015,9430016,9430024,9430071;9470103,9410001,9410011,9410021,9410031,9410041,9410051,9410061,9410071"

'图廓编码
Const TKCodeStr = "9420034,9420035;9430093;9470105"

Sub OnClick()
    
    CodeArr = Split(CodeStr,";", - 1,1)
    TKArr = Split(TKCodeStr,";", - 1,1)
    
    For i = 0 To UBound(TKArr)
        
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "=", TKArr(i)
        SSProcess.SelectFilter
        CodeCount = SSProcess.GetSelGeoCount
        
        If CodeCount > 0 Then
            
            GetExitsCode CodeArr(i),ExistsCodeArr,ExistsCount
            
            CreateWindows ExistsCodeArr,SelArr,SelCount
            
            DrawTuli TKArr(i),SelArr,SelCount
            
        End If
        
    Next 'i
    
End Sub' OnClick


'获取图上存在的Code要素
Function GetExitsCode(ByVal CodeStr,ByRef ExistsCodeArr(),ByRef ExistsCount)
    
    ExistsCount = 0
    
    ReDim ExistsCodeArr(ExistsCount)
    
    CodeArr = Split(CodeStr,",", - 1,1)
    
    For i = 0 To UBound(CodeArr)
        
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "=", CodeArr(i)
        SSProcess.SelectFilter
        CodeCount = SSProcess.GetSelGeoCount
        
        If CodeCount > 0 Then
            ExistsCodeArr(ExistsCount) = SSProcess.GetFeatureCodeInfo(CodeArr(i),"ObjectName") & "【" & CodeArr(i) & "】"
            ExistsCount = ExistsCount + 1
            ReDim Preserve ExistsCodeArr(ExistsCount)
        End If
    Next 'i
    
    ExistsCount = ExistsCount - 1
    
End Function' GetExitsCode

'生成选择弹窗,返回选择的Code
Function CreateWindows(ByVal ExistsCodeArr(),ByRef SelArr(),ByRef SelCount)
    
    SelCount = 0
    
    ReDim SelArr(SelCount)
    
    RecordShortListCount = UBound(ExistsCodeArr) + 1
    ResVal_Dlg = SSFunc.SelectListAttr("选择列表","待选数据列表","选中数据列表",ExistsCodeArr,RecordShortListCount)
    If ResVal_Dlg = 1 Then
        If RecordShortListCount > 0 Then
            For i = 0 To RecordShortListCount - 1
                StrFirst = Replace(ExistsCodeArr(i),"【",",")
                CodeArr = Split(StrFirst,",", - 1,1)
                SelArr(SelCount) = Replace(CodeArr(1),"】","")
                SelCount = SelCount + 1
                ReDim Preserve SelArr(SelCount)
            Next 'i
        End If
    End If
    SelCount = SelCount - 1
End Function' CreateWindows

Function DrawTuli(ByVal TKCode,ByVal CodeArr(),ByVal CodeCount)
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=", TKCode
    SSProcess.SelectFilter
    TKCount = SSProcess.GetSelGeoCount
    
    If TKCode = "9420034,9420035" Then
        
        For i = 0 To TKCount - 1
            TKID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
            DaYBL = SSProcess.GetSelGeoValue(i,"[DaYBL]")
            SSProcess.GetObjectPoint TKID, 1, X, Y, Z, PointType, Name
            For j = 0 To CodeCount
                
                If DrawColor = "" Then
                    DrawColor = SSProcess.GetFeatureCodeInfo(CodeArr(j),"LineColor")
                Else
                    DrawColor = DrawColor & "," & SSProcess.GetFeatureCodeInfo(CodeArr(j),"LineColor")
                End If
                
                If DrawName = "" Then
                    DrawName = SSProcess.GetFeatureCodeInfo(CodeArr(j),"ObjectName")
                Else
                    DrawName = DrawName & "," & SSProcess.GetFeatureCodeInfo(CodeArr(j),"ObjectName")
                End If
                
                If DrawCode = "" Then
                    DrawCode = CodeArr(j)
                Else
                    DrawCode = DrawCode & "," & CodeArr(j)
                End If
                
            Next 'j
            
            JG_ZPT X - 16,Y,TKID,DrawCode,DrawColor,DrawName,500
            
        Next 'i
        
    ElseIf TKCode = "9430093" Then
        
        For i = 0 To TKCount - 1
            TKID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
            SSProcess.GetObjectPoint TKID, 0, X, Y, Z, PointType, Name
            For j = 0 To CodeCount
                
                If DrawColor = "" Then
                    DrawColor = SSProcess.GetFeatureCodeInfo(CodeArr(j),"LineColor")
                Else
                    DrawColor = DrawColor & "," & SSProcess.GetFeatureCodeInfo(CodeArr(j),"LineColor")
                End If
                
                If DrawName = "" Then
                    DrawName = SSProcess.GetFeatureCodeInfo(CodeArr(j),"ObjectName")
                Else
                    DrawName = DrawName & "," & SSProcess.GetFeatureCodeInfo(CodeArr(j),"ObjectName")
                End If
                
                If DrawCode = "" Then
                    DrawCode = CodeArr(j)
                Else
                    DrawCode = DrawCode & "," & CodeArr(j)
                End If
                
            Next 'j
            
            XF_ZPT X,Y,TKID,DrawCode,DrawColor,DrawName
            
        Next 'i
    ElseIf TKCode = "9470105" Then
        For i = 0 To TKCount - 1
            TKID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
            SSProcess.GetObjectPoint TKID, 1, X, Y, Z, PointType, Name
            For j = 0 To CodeCount
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", CodeArr(j)
                SSProcess.SelectFilter
                CodeCount = SSProcess.GetSelGeoCount
                For k = 0 To CodeCount - 1
                    LHLX = SSProcess.GetSelGeoValue(k,"[LHLX]")
                    LHZLX = SSProcess.GetSelGeoValue(k,"[LHZLX]")
                    DrawColor = SSProcess.GetSelGeoValue(k,"SSObj_Color")
                    DrawName = SSProcess.GetFeatureCodeInfo(CodeArr(j),"ObjectName")
                    If LHLX <> "" Then
                        If LHLX <> "休憩场地" Then
                            
                            If ZGNQMC = "" Then
                                ZGNQMC = LHLX
                                If ZDrawCode = "" Then
                                    ZDrawCode = CodeArr(j)
                                Else
                                    ZDrawCode = ZDrawCode & "," & CodeArr(j)
                                End If
                                If ZDrawColor = "" Then
                                    ZDrawColor = DrawColor
                                Else
                                    ZDrawColor = ZDrawColor & "," & DrawColor
                                End If
                            ElseIf Replace(ZGNQMC,LHLX,"") = ZGNQMC Then
                                ZGNQMC = ZGNQMC & "," & LHLX
                                If ZDrawCode = "" Then
                                    ZDrawCode = CodeArr(j)
                                Else
                                    ZDrawCode = ZDrawCode & "," & CodeArr(j)
                                End If
                                If ZDrawColor = "" Then
                                    ZDrawColor = DrawColor
                                Else
                                    ZDrawColor = ZDrawColor & "," & DrawColor
                                End If
                            Else
                                ZGNQMC = ZGNQMC
                            End If
                            
                        ElseIf LHLX = "休憩场地" Then
                            If ZGNQMC = "" Then
                                ZGNQMC = LHZLX
                                If ZDrawCode = "" Then
                                    ZDrawCode = CodeArr(j)
                                Else
                                    ZDrawCode = ZDrawCode & "," & CodeArr(j)
                                End If
                                If ZDrawColor = "" Then
                                    ZDrawColor = DrawColor
                                Else
                                    ZDrawColor = ZDrawColor & "," & DrawColor
                                End If
                            ElseIf Replace(ZGNQMC,LHLX,"") = ZGNQMC Then
                                ZGNQMC = ZGNQMC & "," & LHZLX
                                If ZDrawCode = "" Then
                                    ZDrawCode = CodeArr(j)
                                Else
                                    ZDrawCode = ZDrawCode & "," & CodeArr(j)
                                End If
                                If ZDrawColor = "" Then
                                    ZDrawColor = DrawColor
                                Else
                                    ZDrawColor = ZDrawColor & "," & DrawColor
                                End If
                            Else
                                ZGNQMC = ZGNQMC
                            End If
                        End If
                    Else
                        If ZGNQMC = "" Then
                            ZGNQMC = DrawName
                            If ZDrawCode = "" Then
                                ZDrawCode = CodeArr(j)
                            Else
                                ZDrawCode = ZDrawCode & "," & CodeArr(j)
                            End If
                            If ZDrawColor = "" Then
                                ZDrawColor = DrawColor
                            Else
                                ZDrawColor = ZDrawColor & "," & DrawColor
                            End If
                        ElseIf Replace(ZGNQMC,DrawName,"") = ZGNQMC Then
                            ZGNQMC = ZGNQMC & "," & DrawName
                            If ZDrawCode = "" Then
                                ZDrawCode = CodeArr(j)
                            Else
                                ZDrawCode = ZDrawCode & "," & CodeArr(j)
                            End If
                            If ZDrawColor = "" Then
                                ZDrawColor = DrawColor
                            Else
                                ZDrawColor = ZDrawColor & "," & DrawColor
                            End If
                        Else
                            ZGNQMC = ZGNQMC
                        End If
                    End If
                Next 'k
            Next 'j
            
            LD_ZPT X - 16,Y,ZGNQMC,TKID,ZDrawCode,ZDrawColor
            
        Next 'i
    End If
    
End Function' DrawTuli

'竣工规划总平图图例
Function JG_ZPT(ByVal x0,ByVal y0,ByVal polygonID,ByVal ZDrawCode,ByVal ZDrawColor,ByVal ZDrawName,ByVal DaYBL)
    
    wid2 = (228 * 500) / DaYBL
    heig2 = (286 * 500) / DaYBL
    
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    arDrawName = Split(ZDrawName,",")
    
    count5 = UBound(arDrawCode) + 2
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "=", polygonID
    SSProcess.SelectFilter
    SSProcess.SelectionObjToClipBoard()
    
    SSProcess.DeleteSelectionObj()
    
    AuxiliaryArea x0,y0,x0 + 16,y0,x0 + 16,y0 + count5 * 2 + 2.5,x0,y0 + count5 * 2 + 2.5,AreaId
    
    SSProcess.SelectionObjClip AreaId,0,0.01
    
    SSProcess.AddClipBoardObjToMap 0,0
    SSProcess.DeleteObject AreaId
    
    JG_MakeLine x0,y0,x0,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    JG_MakeLine x0 + 16,y0,x0 + 16,y0 + count5 * 2 + 2.5, 1,"RGB(255,255,255)", polygonID
    JG_MakeLine x0,y0,x0 + 16,y0,1, "RGB(255,255,255)", polygonID
    JG_MakeLine x0,y0 + count5 * 2 + 2.5,x0 + 16,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    JG_MakeNote x0 + 7,y0 + count5 * 2 + 1 , 0, "RGB(255,255,255)", wid2, heig2, "图例",polygonID
    
    
    For j = 0 To UBound(arDrawCode)
        JG_MakeLine x0 + 1,y0 + j * 2 + 1.5,x0 + 7,y0 + j * 2 + 1.5,arDrawCode(j),arDrawColor(j),polygonID
        JG_MakeNote x0 + 10,y0 + 1.5 + j * 2, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID
    Next
    
End Function'JG_ZPT

'消防总平图图例
Function XF_ZPT(ByVal x0,ByVal y0,ByVal polygonID,ByVal ZDrawCode,ByVal ZDrawColor,ByVal ZDrawName)
    
    wid1 = 228
    heig1 = 286
    wid2 = 160
    heig2 = 200
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    arDrawName = Split(ZDrawName,",")
    count5 = UBound(arDrawCode) + 2
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "=", polygonID
    SSProcess.SelectFilter
    SSProcess.SelectionObjToClipBoard()
    
    SSProcess.DeleteSelectionObj()
    
    AuxiliaryArea x0,y0,x0 + 15,y0,x0 + 15,y0 + count5 * 2 + 4,x0,y0 + count5 * 2 + 4,AreaId
    
    SSProcess.SelectionObjClip AreaId,0,0.01
    
    SSProcess.AddClipBoardObjToMap 0,0
    SSProcess.DeleteObject AreaId
    
    XF_MakeLine x0,y0,x0,y0 + count5 * 2 + 4,1, "RGB(255,255,255)", polygonID
    XF_MakeLine x0 + 15,y0,x0 + 15,y0 + count5 * 2 + 4, 1,"RGB(255,255,255)", polygonID
    XF_MakeLine x0,y0,x0 + 15,y0,1, "RGB(255,255,255)", polygonID
    XF_MakeLine x0,y0 + count5 * 2 + 4,x0 + 15,y0 + count5 * 2 + 4,1, "RGB(255,255,255)", polygonID
    XF_MakeNote x0 + 8,y0 + count5 * 2 + 2.5 , 0, "RGB(255,255,255)", wid2, heig2, "图例",polygonID
    
    For j = 0 To UBound(arDrawCode)
        CodeType = SSProcess.GetFeatureCodeInfo(arDrawCode(j), "Type")
        If CodeType = 3 Or  CodeType = 2 Or  CodeType = 1  Then
            XF_MakeLine x0 + 0.5,y0 + 1.5 + j * 2.5,x0 + 5,y0 + 1.5 + j * 2.5 ,arDrawCode(j), arDrawColor(j),polygonID
            XF_MakeNote x0 + 7,y0 + 1.5 + j * 2.5, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID
            
        ElseIf CodeType = 0  Then
            XF_MakePoint x0 + 2 ,y0 + 1 + j * 2.5,arDrawCode(j), arDrawColor(j), polygonID
            XF_MakeNote x0 + 7,y0 + 1.5 + j * 2.5, 0, "RGB(255,255,255)", wid2, heig2,  arDrawName(j),polygonID
            
        ElseIf CodeType = 5  Then
            XF_MakeArea x0 + 0.5,y0 + 0.5 + j * 2.5,x0 + 5,y0 + 0.5 + j * 2.5 ,x0 + 5,y0 + 2.5 + j * 2.5,x0 + 0.5,y0 + 2.5 + j * 2.5 ,arDrawCode(j), arDrawColor(j), polygonID
            XF_MakeNote x0 + 7,y0 + 1.5 + j * 2.5, 0, "RGB(255,255,255)", wid2, heig2,  arDrawName(j),polygonID
        End If
    Next
End Function'XF_ZPT

'绿地总平图图例
Function LD_ZPT(ByVal x0,ByVal y0,ByVal ZGNQMC,ByVal polygonID,ByVal ZDrawCode,ByVal ZDrawColor)
    wid1 = 228
    heig1 = 286
    wid2 = 150
    heig2 = 200
    cvArray1 = Split(ZGNQMC,",")
    count5 = UBound(cvArray1) + 1
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "=", polygonID
    SSProcess.SelectFilter
    SSProcess.SelectionObjToClipBoard()
    
    SSProcess.DeleteSelectionObj()
    
    AuxiliaryArea x0,y0,x0 + 16,y0,x0 + 16,y0 + count5 * 2 + 2.5,x0,y0 + count5 * 2 + 2.5,AreaId
    
    SSProcess.SelectionObjClip AreaId,0,0.01
    
    SSProcess.AddClipBoardObjToMap 0,0
    SSProcess.DeleteObject AreaId
    
    
    LD_MakeLine x0,y0,x0,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    LD_MakeLine x0 + 16,y0,x0 + 16,y0 + count5 * 2 + 2.5, 1,"RGB(255,255,255)", polygonID
    LD_MakeLine x0,y0,x0 + 16,y0,1, "RGB(255,255,255)", polygonID
    LD_MakeLine x0,y0 + count5 * 2 + 2.5,x0 + 16,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    LD_MakeNote x0 + 7,y0 + count5 * 2 + 1.5 , 0, "RGB(255,255,255)", wid2, heig2, "图例",polygonID
    sttr = "地面绿化,地下设施顶面绿化,屋顶绿地"
    For j = 0 To count5 - 1
        If arDrawCode(j) = "9470103" Then
            If InStr(sttr,cvArray1(j)) > 0 Then
                LD_MakeArea x0 + 1,y0 + j * 2 + 0.7,x0 + 7,y0 + j * 2 + 0.7,x0 + 7,y0 + j * 2 + 2.3,x0 + 1,y0 + j * 2 + 2.3,arDrawCode(j), arDrawColor(j), polygonID,"LHLX",cvArray1(j)
            Else
                LD_MakeArea x0 + 1,y0 + j * 2 + 0.7,x0 + 7,y0 + j * 2 + 0.7,x0 + 7,y0 + j * 2 + 2.3,x0 + 1,y0 + j * 2 + 2.3,arDrawCode(j), arDrawColor(j), polygonID,"LHZLX",cvArray1(j)
            End If
        Else
            LD_MakeLine x0 + 1,y0 + j * 2 + 1.5,x0 + 7,y0 + j * 2 + 1.5, arDrawCode(j),"RGB(255,255,255)", polygonID
        End If
        LD_MakeNote x0 + 9,y0 + 1.5 + j * 2, 0, "RGB(255,255,255)", wid2, heig2,cvArray1(j),polygonID
    Next
End Function'LD_ZPT

Function JG_MakeLine(ByVal x1,ByVal y1,ByVal x2,ByVal y2,ByVal code,ByVal color,ByVal polygonID)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function'JG_MakeLine

Function JG_MakeNote(ByVal x,ByVal y,ByVal code,ByVal color,ByVal width,ByVal height,ByVal fontString,ByVal polygonID)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function' JG_MakeNote

Function XF_MakePoint(x,y,code,color,polygonID)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "消防核实总平面测量略图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function XF_MakeLine(x1,y1,x2,y2,code, color, polygonID)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "消防核实总平面测量略图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function XF_MakeArea(x1,y1,x2,y2,x3,y3,x4,y4,code,color,polygonID)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "消防核实总平面测量略图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function XF_MakeNote(x, y, code, color, width, height, fontString,polygonID)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "消防核实总平面测量略图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function LD_MakeLine(x1,y1,x2,y2,code, color, polygonID)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function LD_MakeNote(x, y, code, color, width, height, fontString,polygonID)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "3"
    SSProcess.SetNewObjValue "SSObj_FontWidth", width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function LD_MakeArea(x1,y1,x2,y2,x3,y3,x4,y4,code,color,polygonID,field,LHLX)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    If field = "LHZLX" Then
        SSProcess.SetNewObjValue "[LHLX]", "休憩场所"
        SSProcess.SetNewObjValue "[" & field & "]", LHLX
    Else
        SSProcess.SetNewObjValue "[" & field & "]", LHLX
    End If
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

'绘制辅助面
Function AuxiliaryArea(ByVal X1,ByVal Y1,ByVal X2,ByVal Y2,ByVal X3, ByVal Y3,ByVal X4,ByVal Y4,ByRef AreaId)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_DataMark","辅助面"
    SSProcess.SetNewObjValue "SSObj_Code","2"
    SSProcess.AddNewObjPoint X1, Y1, 0, 0, ""
    SSProcess.AddNewObjPoint X2, Y2, 0, 0, ""
    SSProcess.AddNewObjPoint X3, Y3, 0, 0, ""
    SSProcess.AddNewObjPoint X4, Y4, 0, 0, ""
    SSProcess.AddNewObjPoint X1, Y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
    AreaId = SSProcess.GetGeoMaxID()
    
    IdString = SSProcess.SearchInPolyObjIDs(AreaId,10,"",0,1,1)
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "=", IdString
    SSProcess.SelectFilter
End Function' AuxiliaryArea