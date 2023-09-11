Dim  fileName
Dim xmmc
Dim arID(100000),arID1(100000),arID2(100000)
Dim vArray1(20000), vArray2(20000), vArray3(20000)
Dim cvArray1(20000), cvArray2(20000), cvArray3(20000),vArray(30000)

'一级要素（规划线）倒序数组
strOneOrder = "9410091,9410071,9410061,9410051,9410041,9410031,9410021,9410011,9410001"
Rem special[总平图] 出图前（初始化调用）由此进入
Function VBS_preMap0(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSProcess.PushUndoMark
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
        SSProcess.SelectFilter
        geoCount = SSProcess.GetSelGeoCount()
        For i = 0 To geoCount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            
            
            mdbName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb  mdbName
            sql = "select VALUE from PROJECTINFO where KEY='项目名称'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                XMMC = arSeletionRecord(0)
            Else
                XMMC = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='测量人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                HTRY = arSeletionRecord(0)
            Else
                HTRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='检查人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                JCRY = arSeletionRecord(0)
            Else
                JCRY = ""
            End If
            
            strtemp = XMMC & "," & HTRY & "," & JCRY
            SSProcess.CloseAccessMdb mdbName
            
            SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
        Next
        
        
        
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        fileName = FileFolder & "\竣工测绘总平面图.edb"
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        JGZPTKEY    selectID
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function

Rem special[总平图] 出图完成由此进入
Function VBS_postMap0(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    SSProcess.SetMapScale "500"
    
    'DaHui
    'DeleteFeature "9410091","9420033"
    'DeleteFeature "9420035","9999403"
    DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,基底注记,TKZSX,TKZSM,DEFAULT"
    
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420034
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        
        ids = SSProcess.SearchInnerObjIDs(id, 1, "9410001", 0)
        
        If ids <> "" Then
            idsList = Split(ids,",")
            strtemp = SSProcess.GetObjectAttr(idsList(0), "SSObj_DataMark")
            artemp = Split(strtemp,",")
        Else
            strtemp = ",,"
            artemp = Split(strtemp,",")
        End If
        SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
        SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
        SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
        '图形重新生成
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    Next
    
    CreateKEYZPT
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

Rem special[对比图] 出图前（初始化调用）由此进入
Function VBS_preMap1(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSParameter.SetParameterINT "printMap", "return", 1
        Dim path_print
        If CheckReportPath(path_print) = False    Then
            
            MsgBox "无法完成出图：成果目录未创建、无法完成出图"
            Exit Function
        End If
        
        
        If GetXMMC(xmmc) = False Then
            MsgBox "无法完成出图：请检查项目名称是否正确？"
            Exit Function
        End If
        
        
        fileName = path_print & "\" & xmmc & "四至尺寸对比图.edb"
        
        If FileExists(fileName) Then
            MsgBox fileName & "文件已存在、无法完成输出、请手动检查删除后重试"
            Exit Function
        End If
        SSParameter.SetParameterINT "printMap", "return", 0
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[对比图] 出图完成由此进入
Function VBS_postMap1(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    
    
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

Rem special[土地核验] 出图前（初始化调用）由此进入
Function VBS_preMap2(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSParameter.SetParameterINT "printMap", "return", 1
        Dim path_print
        If CheckReportPath(path_print) = False    Then
            
            MsgBox "无法完成出图：成果目录未创建、无法完成出图"
            Exit Function
        End If
        
        
        If GetXMMC(xmmc) = False Then
            MsgBox "无法完成出图：请检查项目名称是否正确？"
            Exit Function
        End If
        
        
        fileName = path_print & "\" & xmmc & "土地核验测量图.edb"
        
        If FileExists(fileName) Then
            MsgBox fileName & "文件已存在、无法完成输出、请手动检查删除后重试"
            Exit Function
        End If
        SSParameter.SetParameterINT "printMap", "return", 0
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[土地核验] 出图完成由此进入
Function VBS_postMap2(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    
    
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

Rem special[基底图] 出图前（初始化调用）由此进入
Function VBS_preMap3(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        '输出地块内图斑总面积
        SSProcess.PushUndoMark
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "==", 9420025
        SSProcess.SelectFilter
        geoCount = SSProcess.GetSelGeoCount()
        For i = 0 To geoCount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            
            
            mdbName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb  mdbName
            sql = "select VALUE from PROJECTINFO where KEY='项目名称'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                XMMC = arSeletionRecord(0)
            Else
                XMMC = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='测量人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                HTRY = arSeletionRecord(0)
            Else
                HTRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='检查人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                JCRY = arSeletionRecord(0)
            Else
                JCRY = ""
            End If
            
            strtemp = XMMC & "," & HTRY & "," & JCRY
            SSProcess.CloseAccessMdb mdbName
            
            SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
        Next
        
        
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        LiMianCL "建筑基底平面示意图",FileFolder, intCount
        fileName = FileFolder & "\建筑基底平面示意图" & intCount & ".edb"
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
        
    End If
    
End Function


Rem special[基底图] 出图完成由此进入
Function VBS_postMap3(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    '// 添加您的成果图处理代码 
    GHXKZGUID = SSProcess.GetObjectAttr (selectID,"[JSGHXKZGUID]")
    jzdtguid = SSProcess.GetObjectAttr (selectID,"[JZWMCGUID]")
    GHXKZHoutmap = SSProcess.GetObjectAttr (selectID,"[GuiHXKZBH]")
    JianZWMC = SSProcess.GetObjectAttr (selectID,"[JianZWMC]")
    JiDMJ = SSProcess.GetObjectAttr (selectID,"[JiDMJ]")
    
    SSProcess.SetObjectAttr tk_id,"[JSGHXKZGUID]",GHXKZGUID
    SSProcess.SetObjectAttr tk_id,"[JZWMCGUID]",jzdtguid
    SSProcess.SetObjectAttr tk_id,"[GuiHXKZBH]",GHXKZHoutmap
    SSProcess.SetObjectAttr tk_id,"[JianZWMC]",JianZWMC
    SSProcess.SetObjectAttr tk_id,"[JiDMJ]",JiDMJ
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    CreateKEYJD()
    reset()
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420032
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        
        ids = SSProcess.SearchInnerObjIDs(id, 2, "9420025", 0)
        
        If ids <> "" Then
            idsList = Split(ids,",")
            strtemp = SSProcess.GetObjectAttr(idsList(0), "SSObj_DataMark")
            artemp = Split(strtemp,",")
        Else
            strtemp = ",,"
            artemp = Split(strtemp,",")
        End If
        
        SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
        SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
        SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
        '图形重新生成
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    Next
    
    
    XF_Info()
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

'.................................................................................辅助基底图............................................

Function XF_Info()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9420025,9420024,9420026"
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount
    For i = 0 To SelCount - 1
        ID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        pointcount = SSProcess.GetSelGeoPointCount(i)
        IsClockwise = SSProcess.IsClockwise( ID )
        JDMC = SSProcess.GetObjectAttr (ID, "[JDMC]")
        If JDMC <> "" Then
            SSProcess.GetObjectFocusPoint ID, x2, y2
            CreatNote x2, y2,JDMC
            For j = 0 To PointCount - 1
                SSProcess.GetSelGeoPoint i, j, x, y, z, pointtype, name
                SSProcess.GetSelGeoPoint i, j + 1, x1, y1, z1, pointtype, name
                '取长度大于2的边进行标注
                If  ((x1 - x) * (x1 - x) + (y1 - y) * (y1 - y)) ^ 0.5 > 4then
                    GetLineLength IsClockwise,x,y,x1,y1,Length1,CneterX1,CneterY1,Angle1
                    DrawNote Length1,CneterX1,CneterY1,Angle1
                End If
            Next
        End If
        'GetLineLength ID,Length1,Length2,CneterX1,CneterY1,Angle1,CneterX2,CneterY2,Angle2
        'DrawNote Length1,CneterX1,CneterY1,Angle1
        'DrawNote Length2,CneterX2,CneterY2,Angle2
    Next 'i
End Function' JdAreaInfo

'获取第一边和第二边的边长
Function GetLineLength(ByRef IsClockwise, ByRef x ,ByRef y,ByRef x1,ByRef y1,ByRef Length1,ByRef NoteX1,ByRef NoteY1,ByRef Angle1)
    
    
    CneterX1 = (x + x1) / 2
    CneterY1 = (y + y1) / 2
    
    SSProcess.XYSA x,y,x1,y1,Length1,Angle1,0
    Length1 = CStr(Round(Length1,2))
    Sin1 = Sin(Angle1)
    COS1 = Cos(Angle1)
    
    Angle1 = SSProcess.RadianToDeg(Angle1)
    
    '确定象限
    If Angle1 = 0 Then
        Quadrant1 = 11
    ElseIf Angle1 = 90 Then
        Quadrant1 = 22
    ElseIf Angle1 = 180 Then
        Quadrant1 = 33
    ElseIf Angle1 = 270 Then
        Quadrant1 = 44
    ElseIf Angle1 > 0 And Angle1 < 90 Then
        Quadrant1 = 1
    ElseIf Angle1 > 90 And Angle1 < 180 Then
        Quadrant1 = 2
    ElseIf Angle1 > 180 And  Angle1 < 270 Then
        Quadrant1 = 3
    ElseIf Angle1 > 270 And  Angle1 < 360 Then
        Quadrant1 = 4
    End If
    '顺时针
    
    If IsClockwise = 0 Then
        
        If Quadrant1 = 11 Then
            NoteY1 = CneterY1 - 0.5
            NoteX1 = CneterX1
        ElseIf Quadrant1 = 22 Then
            NoteY1 = CneterY1
            NoteX1 = CneterX1 + 0.5
        ElseIf Quadrant1 = 33 Then
            NoteY1 = CneterY1 + 0.5
            NoteX1 = CneterX1
        ElseIf Quadrant1 = 44 Then
            NoteY1 = CneterY1
            NoteX1 = CneterX1 - 0.5
        ElseIf Quadrant1 = 1 Then
            NoteY1 = CneterY1 + Abs(Sin1 * 2)
            NoteX1 = CneterX1 - Abs(COS1 * 0.5)
        ElseIf Quadrant1 = 2 Then
            NoteY1 = CneterY1 - Abs(Sin1 * 2)
            NoteX1 = CneterX1 + Abs(COS1 * 0.5)
        ElseIf Quadrant1 = 3 Then
            NoteY1 = CneterY1 - Abs(Sin1 * 0.5)
            NoteX1 = CneterX1 + Abs(COS1 * 2)
        ElseIf  Quadrant1 = 4 Then
            NoteY1 = CneterY1 - Abs(Sin1 * 0.5)
            NoteX1 = CneterX1 - Abs(COS1 * 2)
        End If
        
    Else
        
        '逆时针
        If Quadrant1 = 11 Then
            NoteY1 = CneterY1 + 0.5
            NoteX1 = CneterX1
        ElseIf Quadrant1 = 22 Then
            NoteY1 = CneterY1
            NoteX1 = CneterX1 - 0.5
        ElseIf Quadrant1 = 33 Then
            NoteY1 = CneterY1 - 0.5
            NoteX1 = CneterX1
        ElseIf Quadrant1 = 44 Then
            NoteY1 = CneterY1
            NoteX1 = CneterX1 + 0.5
            
        ElseIf Quadrant1 = 1 Then
            NoteY1 = CneterY1 - Abs(Sin1 * 0.5)
            NoteX1 = CneterX1 + Abs(COS1 * 2)
        ElseIf Quadrant1 = 2 Then
            NoteY1 = CneterY1 + Abs(Sin1 * 2)
            NoteX1 = CneterX1 + Abs(COS1 * 0.5)
        ElseIf Quadrant1 = 3 Then
            NoteY1 = CneterY1 - Abs(Sin1 * 0.5)
            NoteX1 = CneterX1 + Abs(COS1 * 2)
        ElseIf  Quadrant1 = 4 Then
            NoteY1 = CneterY1 - Abs(Sin1 * 0.5)
            NoteX1 = CneterX1 - Abs(COS1 * 2)
        End If
    End If
    
    
End Function' GetLineLength

Function DrawNote(ByVal Bras,ByVal X,ByVal Y,ByVal Angle)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", Bras
    SSProcess.SetNewObjValue "SSObj_FontWordAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontStringAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", 200
    SSProcess.SetNewObjValue "SSObj_FontHeight", 200
    SSProcess.SetNewObjValue "SSObj_FontDirection", 0
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function' DrawNote

Function CreatNote(x,y,note)
    SSProcess.PushUndoMark
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "JD001"
    SSProcess.SetNewObjValue "SSObj_FontString", note
    SSProcess.SetNewObjValue "SSObj_Color", "RGB(255,255,255)"
    SSProcess.SetNewObjValue "SSObj_FontWidth", "350"
    SSProcess.SetNewObjValue "SSObj_FontHeight", "350"
    SSProcess.AddNewObjPoint x, y + 1.5,z, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function
'.................................................................................辅助基底图............................................



Rem special[分层图] 出图前（初始化调用）由此进入
Function VBS_preMap4(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        TKFZ1
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        
        ZRZH = SSProcess.GetObjectAttr (selectID, "[LD]")
        fileName = FileFolder & "\" & ZRZH & "建筑功能分区竣工实测平面图.edb"
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        SSParameter.SetParameterINT "printMap", "return", 1
        If GetXMMC(xmmc) = False Then
            MsgBox "无法完成出图：请检查项目名称是否正确？"
            Exit Function
        End If
        'SSProcess.WriteEpsIni "批前测量分层图", "项目名称" , xmmc
        
        dh = SSProcess.GetObjectAttr (selectID,"[JianZWMC]")
        If (dh = "" Or dh = "*") Then
            MsgBox "无法完成出图：请检查项目[建筑物名称]是否合法？"
            Exit Function
        End If
        SSParameter.SetParameterINT "printMap", "return", 0
        
    End If
    
End Function




Rem special[分层图] 出图完成由此进入
Function VBS_postMap4(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterSTR "printMap", "TKIDS", - 1, tk_ids
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    
    
    '// 添加您的成果图处理代码 
    
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    '重置
    reset()
    '图廓赋值
    TKFZ2 mark
    If mark = True Then
        '分层图
        FChandle()
        '标注
        TextEXE()
        '删除楼层
        FCTDeleteLC()
        '生成图例
        CreateKEY()
        '生成边长标注
        FCTBC_Info()
    End If
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function


Function FCTBC_Info()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9420023"
    SSProcess.SetSelectCondition "[YSBM]", "<>", ""
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount
    For i = 0 To SelCount - 1
        ID = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        pointcount = SSProcess.GetSelGeoPointCount(i)
        IsClockwise = SSProcess.IsClockwise( ID )
        For j = 0 To PointCount - 1
            SSProcess.GetSelGeoPoint i, j, x, y, z, pointtype, name
            SSProcess.GetSelGeoPoint i, j + 1, x1, y1, z1, pointtype, name
            '取长度大于2的边进行标注
            If  ((x1 - x) * (x1 - x) + (y1 - y) * (y1 - y)) ^ 0.5 > 4 Then
                GetLineLength IsClockwise,x,y,x1,y1,Length1,CneterX1,CneterY1,Angle1
                DrawNote1 Length1,CneterX1,CneterY1,Angle1
            End If
        Next
        'GetLineLength ID,Length1,Length2,CneterX1,CneterY1,Angle1,CneterX2,CneterY2,Angle2
        'DrawNote Length1,CneterX1,CneterY1,Angle1
        'DrawNote Length2,CneterX2,CneterY2,Angle2
    Next 'i
End Function' JdAreaInfo


Function DrawNote1(ByVal Bras,ByVal X,ByVal Y,ByVal Angle)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontString", Bras
    SSProcess.SetNewObjValue "SSObj_FontWordAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontStringAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth", 120
    SSProcess.SetNewObjValue "SSObj_FontHeight", 120
    SSProcess.SetNewObjValue "SSObj_FontDirection", 0
    SSProcess.AddNewObjPoint X,Y,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function' DrawNote



Rem special[立面图 出图前（初始化调用）由此进入
Function VBS_preMap5(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        
        fileName = FileFolder & "\" & "立面图.edb"
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        SSParameter.SetParameterINT "printMap", "return", 1
        If GetXMMC(xmmc) = False Then
            MsgBox "无法完成出图：请检查项目名称是否正确？"
            Exit Function
        End If
        'SSProcess.WriteEpsIni "批前测量分层图", "项目名称" , xmmc
        
        Dim dh
        dh = SSProcess.GetObjectAttr (selectID,"[JianZWMC]")
        
        If (dh = "" Or dh = "*") Then
            MsgBox "无法完成出图：请检查项目[建筑物名称]是否合法？"
            Exit Function
        End If
        SSParameter.SetParameterINT "printMap", "return", 0
        
    End If
    
End Function

Rem special[立面图] 出图完成由此进入
Function VBS_postMap5(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    JianZWMC = SSProcess.GetObjectAttr (selectID,"[JianZWMC]")
    GuiHXKZBH = SSProcess.GetObjectAttr (selectID,"[GuiHXKZBH]")
    SSProcess.SetObjectAttr tk_id,"[JianZWMC]",JianZWMC
    SSProcess.SetObjectAttr tk_id,"[GuiHXKZBH]",GuiHXKZBH
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    
    
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function


Rem special[地上停车位] 出图前（初始化调用）由此进入
Function VBS_preMap6(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSParameter.SetParameterINT "printMap", "return", 1
        Dim path_print
        If CheckReportPath(path_print) = False    Then
            
            MsgBox "无法完成出图：成果目录未创建、无法完成出图"
            Exit Function
        End If
        
        
        If GetXMMC(xmmc) = False Then
            MsgBox "无法完成出图：请检查项目名称是否正确？"
            Exit Function
        End If
        
        
        fileName = path_print & "\" & xmmc & "地上停车位分布图.edb"
        
        If FileExists(fileName) Then
            MsgBox fileName & "文件已存在、无法完成输出、请手动检查删除后重试"
            Exit Function
        End If
        SSParameter.SetParameterINT "printMap", "return", 0
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[地上停车位] 出图完成由此进入
Function VBS_postMap6(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    
    
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

Rem special[地下停车位] 出图前（初始化调用）由此进入
Function VBS_preMap7(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSParameter.SetParameterINT "printMap", "return", 1
        Dim path_print
        If CheckReportPath(path_print) = False    Then
            
            MsgBox "无法完成出图：成果目录未创建、无法完成出图"
            Exit Function
        End If
        
        
        If GetXMMC(xmmc) = False Then
            MsgBox "无法完成出图：请检查项目名称是否正确？"
            Exit Function
        End If
        
        
        fileName = path_print & "\" & xmmc & "地下停车位分布图.edb"
        
        If FileExists(fileName) Then
            MsgBox fileName & "文件已存在、无法完成输出、请手动检查删除后重试"
            Exit Function
        End If
        SSParameter.SetParameterINT "printMap", "return", 0
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function

Rem special[地下停车位图] 出图完成由此进入
Function VBS_postMap7(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterSTR "printMap", "TKIDS", - 1, tk_ids
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    GHXKZGUID = SSProcess.GetObjectAttr (selectID,"[JSGHXKZGUID]")
    jzdtguid = SSProcess.GetObjectAttr (selectID,"[JZWMCGUID]")
    GHXKZHoutmap = SSProcess.GetObjectAttr (selectID,"[GuiHXKZBH]")
    ' GetXKZXX GHXKZHoutmap,JZWMCoutmap,GHXKZGUID,jzdtguid
    BLC = SSProcess.GetMapScale
    
    'msgbox jzdtguid
    '// 添加您的成果图处理代码 
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "==", tk_ids
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    Dim arcc(1000)
    For i = 0 To geoCount - 1
        tk_id = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        
        cc = SSProcess.GetSelGeoValue(i,"[CengC]")
        
        '    GetCGXX GHXKZHoutmap,JZWMCoutmap,cc,GHCD,SJCG,CMXX,sjcs
        
        If InStr(cc,"-") = 1 Then
            DSDXBS = "地下"
        Else
            DSDXBS = "地上"
        End If
        'If GHCD <> "" Then  SSProcess.SetObjectAttr tk_id,"[PiZCG]",GHCD
        'If SJCG <> "" Then  SSProcess.SetObjectAttr tk_id,"[ShiCCG]",SJCG
        SSProcess.SetObjectAttr tk_id,"[BiLC]",BLC
        'SSProcess.SetObjectAttr tk_id,"[DiSDXBS]",DSDXBS
        'SSProcess.SetObjectAttr tk_id,"[CengS]",sjcs
        SSProcess.SetObjectAttr tk_id,"[JSGHXKZGUID]",GHXKZGUID
        SSProcess.SetObjectAttr tk_id,"[JZWMCGUID]",jzdtguid
        SSProcess.SetObjectAttr tk_id,"[GuiHXKZBH]",GHXKZHoutmap
        'SSProcess.SetObjectAttr tk_id,"[BeiZ]","说明：1、该层建筑面积按实测尺寸计算。     \2、实测尺寸已知扣除抹灰厚度（抹灰厚度平均0.03m）。"
        If InStr(cc,"-") > 0 Then
            cc = Split(cc,"+")
            If UBound(cc) = 1 Then
                If InStr(cc(0),".") = 0 Then
                    str0 = SSFunc.GetChineseDigit(Abs (cc(0)))
                Else
                    str0 = cc(0)
                End If
                If InStr(cc(1),".") = 0 Then
                    str1 = SSFunc.GetChineseDigit(Abs(cc(1)))
                Else
                    str1 = cc(1)
                End If
                str111 = "地下" & str0 & "层至" & str1 & "层"
            ElseIf UBound(cc) = 0 Then
                If InStr(cc(0),".") = 0 Then
                    str0 = SSFunc.GetChineseDigit(Abs (cc(0)))
                Else
                    str0 = cc(0)
                End If
                str111 = "地下" & str0 & "层"
            End If
        Else
            SSFunc.ScanString cc, ",", arcc, arccCount
            For c = 0 To arccCount - 1
                cc = Split(arcc(c),"+")
                If UBound(cc) = 1 Then
                    If InStr(cc(0),".") = 0 Then
                        str0 = SSFunc.GetChineseDigit(Abs (cc(0)))
                    Else
                        str0 = cc(0)
                    End If
                    
                    If InStr(cc(1),".") = 0 Then
                        str1 = SSFunc.GetChineseDigit(Abs(cc(1)))
                    Else
                        str1 = cc(1)
                    End If
                    str = str0 & "层至" & str1 & "层"
                ElseIf UBound(cc) = 0 Then
                    If InStr(cc(0),".") = 0 Then
                        str0 = SSFunc.GetChineseDigit(Abs (cc(0)))
                    Else
                        str0 = cc(0)
                    End If
                    str = str0 & "层"
                End If
                If str111 = "" Then
                    str111 = str
                Else
                    str111 = str111 & "、" & str
                End If
                str = ""
            Next
            
        End If
        
        SSProcess.SetObjectAttr tk_id,"[CengM]",str111
        
        str111 = ""
        
        ids = SSProcess.SearchInnerObjIDs( tk_id, 10, "1", 0 )
        If ids <> "" Then
            If change_ids = "" Then
                change_ids = ids
            Else
                change_ids = change_ids & "," & ids
            End If
            
        End If
        
    Next
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

Rem special[绿地图] 出图前（初始化调用）由此进入
Function VBS_preMap8(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSParameter.SetParameterINT "printMap", "return", 1
        Dim path_print
        If CheckReportPath(path_print) = False    Then
            
            MsgBox "无法完成出图：成果目录未创建、无法完成出图"
            Exit Function
        End If
        
        
        If GetXMMC(xmmc) = False Then
            MsgBox "无法完成出图：请检查项目名称是否正确？"
            Exit Function
        End If
        
        
        fileName = path_print & "\" & xmmc & "绿地面积统计图.edb"
        
        If FileExists(fileName) Then
            MsgBox fileName & "文件已存在、无法完成输出、请手动检查删除后重试"
            Exit Function
        End If
        SSParameter.SetParameterINT "printMap", "return", 0
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[绿地图] 出图完成由此进入
Function VBS_postMap8(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    
    
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function


Rem special[竣工图] 出图前（初始化调用）由此进入
Function VBS_preMap9(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSProcess.PushUndoMark
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
        SSProcess.SelectFilter
        geoCount = SSProcess.GetSelGeoCount()
        For i = 0 To geoCount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            
            
            mdbName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb  mdbName
            sql = "select VALUE from PROJECTINFO where KEY='项目名称'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            
            If nSeletionCount > 0 Then
                XMMC = arSeletionRecord(0)
            Else
                XMMC = ""
            End If
            
            
            sql = "select VALUE from PROJECTINFO where KEY='测量人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                HTRY = arSeletionRecord(0)
            Else
                HTRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='检查人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                JCRY = arSeletionRecord(0)
            Else
                JCRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='测绘单位'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                CLDW = arSeletionRecord(0)
            Else
                CLDW = ""
            End If
            
            strtemp = XMMC & "," & HTRY & "," & JCRY & "," & CLDW
            SSProcess.CloseAccessMdb mdbName
            
            SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
        Next
        
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        fileName = FileFolder & "\竣工图.edb"
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[竣工图] 出图完成由此进入
Function VBS_postMap9(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    'DeleteFeature "9410101","9420032"
    'DeleteFeature "9420034","9999403"
    'DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,基底注记,TKZSX,TKZSM"
    DaHui
    SSProcess.SetMapScale "500"
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420033
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        
        ids = SSProcess.SearchInnerObjIDs(id, 1, "9410001", 0)
        
        If ids <> "" Then
            idsList = Split(ids,",")
            strtemp = SSProcess.GetObjectAttr(idsList(0), "SSObj_DataMark")
            artemp = Split(strtemp,",")
        Else
            strtemp = ",,,"
            artemp = Split(strtemp,",")
        End If
        
        SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
        SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
        SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
        SSProcess.SetObjectAttr id, "[测量单位名称]", artemp(3)
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
        '图形重新生成
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    Next
    DeleteFeatureLayerName "规划线,GHX"
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function


Rem special[竣工规划复核图] 出图前（初始化调用）由此进入
Function VBS_preMap10(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSProcess.PushUndoMark
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
        SSProcess.SelectFilter
        geoCount = SSProcess.GetSelGeoCount()
        For i = 0 To geoCount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            
            
            mdbName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb  mdbName
            sql = "select VALUE from PROJECTINFO where KEY='项目名称'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                XMMC = arSeletionRecord(0)
            Else
                XMMC = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='测量人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                HTRY = arSeletionRecord(0)
            Else
                HTRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='检查人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                JCRY = arSeletionRecord(0)
            Else
                JCRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='测绘单位'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                CLDW = arSeletionRecord(0)
            Else
                CLDW = ""
            End If
            
            
            sql = "select VALUE from PROJECTINFO where KEY='审核人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                SHRY = arSeletionRecord(0)
            Else
                SHRY = ""
            End If
            
            
            strtemp = XMMC & "," & HTRY & "," & JCRY & "," & CLDW & "," & SHRY
            SSProcess.CloseAccessMdb mdbName
            
            SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
        Next
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        fileName = FileFolder & "\竣工规划复核图.edb"
        
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[竣工规划复核图] 出图完成由此进入
Function VBS_postMap10(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    'DaHui
    DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,基底注记,TKZSX,TKZSM"
    SSProcess.SetMapScale "500"
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420035
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        
        ids = SSProcess.SearchInnerObjIDs(id, 1, "9410001", 0)
        
        If ids <> "" Then
            idsList = Split(ids,",")
            strtemp = SSProcess.GetObjectAttr (idsList(0), "SSObj_DataMark")
            artemp = Split(strtemp,",")
        Else
            strtemp = ",,,,"
            artemp = Split(strtemp,",")
        End If
        
        SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
        SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
        SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
        SSProcess.SetObjectAttr id, "[测量单位名称]", artemp(3)
        SSProcess.SetObjectAttr id, "[ShenHY]", artemp(4)
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
        '图形重新生成
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    Next
    
    CreateKEYZPT()
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

Rem special[用地复核图] 出图前（初始化调用）由此进入
Function VBS_preMap11(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSProcess.PushUndoMark
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
        SSProcess.SelectFilter
        geoCount = SSProcess.GetSelGeoCount()
        For i = 0 To geoCount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            
            
            mdbName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb  mdbName
            sql = "select VALUE from PROJECTINFO where KEY='项目名称'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                XMMC = arSeletionRecord(0)
            Else
                XMMC = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='测量人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                HTRY = arSeletionRecord(0)
            Else
                HTRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='检查人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                JCRY = arSeletionRecord(0)
            Else
                JCRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='测绘单位'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                CLDW = arSeletionRecord(0)
            Else
                CLDW = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='审核人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                SHRY = arSeletionRecord(0)
            Else
                SHRY = ""
            End If
            
            strtemp = XMMC & "," & HTRY & "," & JCRY & "," & CLDW & "," & SHRY
            SSProcess.CloseAccessMdb mdbName
            
            SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
        Next
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        fileName = FileFolder & "\用地复核图.edb"
        
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[用地复核图] 出图完成由此进入
Function VBS_postMap11(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    DaHui
    DeleteFeature "9410011","9420035"
    DeleteFeature "9420037","9999403"
    DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,基底注记,TKZSX,TKZSM"
    
    JZDSC
    JZX
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420036
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        
        ids = SSProcess.SearchInnerObjIDs(id, 1, "9410001", 0)
        
        If ids <> "" Then
            idsList = Split(ids,",")
            strtemp = SSProcess.GetObjectAttr (idsList(0), "SSObj_DataMark")
            artemp = Split(strtemp,",")
        Else
            strtemp = ",,,,"
            artemp = Split(strtemp,",")
        End If
        
        
        
        
        SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
        SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
        SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
        SSProcess.SetObjectAttr id, "[测量单位名称]", artemp(3)
        SSProcess.SetObjectAttr id, "[ShenHY]", artemp(4)
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
        '图形重新生成
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    Next
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

Rem special[基地总平面布置核实测量平面图] 出图前（初始化调用）由此进入
Function VBS_preMap12(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        SSProcess.PushUndoMark
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
        SSProcess.SelectFilter
        geoCount = SSProcess.GetSelGeoCount()
        For i = 0 To geoCount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            
            
            mdbName = SSProcess.GetProjectFileName
            SSProcess.OpenAccessMdb  mdbName
            sql = "select VALUE from PROJECTINFO where KEY='项目名称'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                XMMC = arSeletionRecord(0)
            Else
                XMMC = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='测量人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                HTRY = arSeletionRecord(0)
            Else
                HTRY = ""
            End If
            
            sql = "select VALUE from PROJECTINFO where KEY='检查人员'"
            GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
            If nSeletionCount > 0 Then
                JCRY = arSeletionRecord(0)
            Else
                JCRY = ""
            End If
            
            strtemp = XMMC & "," & HTRY & "," & JCRY
            SSProcess.CloseAccessMdb mdbName
            
            SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
        Next
        
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        fileName = FileFolder & "\基地总平面布置核实测量平面图.edb"
        
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[基地总平面布置核实测量平面图] 出图完成由此进入
Function VBS_postMap12(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    DaHui
    
    DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,基底注记,TKZSX,TKZSM"
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420037
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        
        ids = SSProcess.SearchInnerObjIDs(id, 1, "9410001", 0)
        
        If ids <> "" Then
            idsList = Split(ids,",")
            strtemp = SSProcess.GetObjectAttr(idsList(0), "SSObj_DataMark")
            artemp = Split(strtemp,",")
        Else
            strtemp = ",,"
            artemp = Split(strtemp,",")
        End If
        SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
        SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
        SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
        '图形重新生成
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    Next
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function

Rem special[总平面测量略图] 出图前（初始化调用）由此进入
Function VBS_preMap13(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        fileName = FileFolder & "\总平面测量略图.edb"
        
        
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[总平面测量略图] 出图完成由此进入
Function VBS_postMap13(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    DaHui
    DeleteFeature "9410011","9420037"
    DeleteFeature "9420039","9420108"
    DeleteFeature "9450013","9999403"
    DeleteFeatureLayerName "DT_POLYGON,DT_LINE,DT_POINT,DT_ZJ,LCGZKJ,LCGZKJZJ,GHFSQLJX,LMT_ZJ,基底注记,TKZSX,TKZSM"
    
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function


Rem special[立面图] 出图前（初始化调用）由此进入
Function VBS_preMap14(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... 
        strProjectName = SSProcess.GetProjectFileName()
        FileFolder = Replace(strProjectName,".edb","")
        If IsfolderExists(FileFolder) = False Then CreateFolders FileFolder
        
        LiMianCL "立面图",FileFolder, intCount
        fileName = FileFolder & "\立面图" & intCount & ".edb"
        
        SSParameter.SetParameterSTR "printMap","NewedbName",fileName
        
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    End If
    
End Function




Rem special[立面图] 出图完成由此进入
Function VBS_postMap14(MSGID,mapName,selectID)
    
    Rem 图廓ID,脚本处理项个数
    Dim tk_id,tk_innerids,ScriptChangeCount
    Rem 脚本处理项名称,脚本处理项参数,脚本处理项附加参数
    Dim str_Name,str_para,str_paraex
    Rem 获取分层图图廓IDS,多个英文逗号相隔
    SSParameter.GetParameterINT "printMap", "TKID", - 1, tk_id
    Rem 获取图廓内地物IDS
    SSParameter.GetParameterSTR "printMap", "TKInerobjIDS", "", tk_innerids
    Rem 获取脚本处理项个数
    SSParameter.GetParameterINT "printMap", "ScriptChangeCount", - 1, ScriptChangeCount
    
    '// 添加您的成果图处理代码 
    
    SSProcess.SetObjectAttr tk_id,"[XiangMMC]",xmmc
    
    
    '// 添加您的成果图处理代码 
    debug_print String(50,"-")
    debug_print "输出完成。"
    debug_print String(50,"-")
    ViewExtend()
    
    
    
    Rem 成果图细节分开处理
    For i = 0 To ScriptChangeCount - 1
        Rem 获取处理项名称
        SSParameter.GetParameterSTR "printMap", i & "Name", "", str_Name
        Rem 获取处理项参数
        SSParameter.GetParameterSTR "printMap", i & "ParaString", "", str_para
        Rem 获取处理项附加参数
        SSParameter.GetParameterSTR "printMap", i & "ParaStringEX", "", str_paraex
        '// 此处无代码、说明没有脚本处理项..
    Next
    
End Function


Function LiMianCL(ByVal mapname,ByVal FileFolder,ByRef intCount)
    intCount = 1
    '创建FileSystemObject对象
    Set objFso = CreateObject("Scripting.FileSystemObject")
    '使用GetFolder()获得文件夹对象
    Set objGetFolder = objFso.GetFolder(FileFolder)
    '遍历Files集合并显示文件夹中所有的文件名
    For Each strFile In objGetFolder.Files
        If objFso.GetExtensionName(strFile) = "edb" Then
            If InStr(strFile.Name ,mapname) > 0 Then intCount = intCount + 1
        End If
    Next
    
End Function


Dim g_MapList,g_MapPrePtrfun,g_MapPostPtrfun

Sub OnClick()
    
    Rem 初始化
    g_MapList = Array("总平面图","四至尺寸对比图","土地核验图","建筑基底图","分层面积计算图","建筑立面示意图","地上停车位分布图","地下停车位分布图","绿地面积统计图","竣工图","竣工规划复核图","用地复核图","基地总平面布置核实测量平面图","总平面测量略图","立面图")
    g_MapPrePtrfun = Array("VBS_preMap0","VBS_preMap1","VBS_preMap2","VBS_preMap3","VBS_preMap4","VBS_preMap5","VBS_preMap6","VBS_preMap7","VBS_preMap8","VBS_preMap9","VBS_preMap10","VBS_preMap11","VBS_preMap12","VBS_preMap13","VBS_preMap14")
    g_MapPostPtrfun = Array("VBS_postMap0","VBS_postMap1","VBS_postMap2","VBS_postMap3","VBS_postMap4","VBS_postMap5","VBS_postMap6","VBS_postMap7","VBS_postMap8","VBS_postMap9","VBS_postMap10","VBS_postMap11","VBS_postMap12","VBS_postMap13","VBS_postMap14")
    
    Rem 系统传来的消息,用户选择的范围线ID,成果图名称
    Dim str_msg,str_selectObjid,str_mapName
    
    Rem 获取系统参数 -  - 用户选择范围线ID
    SSParameter.GetParameterINT "printMap", "SelectID", - 1, str_selectObjid
    
    Rem 获取系统参数 -  - 系统消息 （0：本工程出图初始化消息 1：新工程固定目录出图初始化消息  2新工程自定义目录出图初始化消息 - 1：出图已完成交付于脚本处理细节）
    SSParameter.GetParameterINT "printMap", "printMSG", - 1, str_msg
    
    Rem 获取系统参数 -  - 专题名称
    SSParameter.GetParameterSTR "printMap", "SpecialMapName", "", str_mapName
    DistributeMSG str_msg,str_mapName,str_selectObjid
End Sub

Sub ViewExtend()
    
    '图形范围全视
    
    SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
    
    '图形重新生成
    
    SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    
End Sub

'// 判断文件是否存在
Function FileExists(fileName)
    
    
    Dim fso
    Set fso = CreateObject("scripting.filesystemobject")
    FileExists = fso.FileExists(fileName)
    
End Function

'创建文件夹
Function CreateFolders(path)
    Set fso = CreateObject("scripting.filesystemobject")
    CreateFolderEx fso,path
    Set fso = Nothing
End Function

Function CreateFolderEx(fso,path)
    If fso.FolderExists(path) Then
        Exit Function
    End If
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then
        CreateFolderEx fso,fso.GetParentFolderName(path)
    End If
    fso.CreateFolder(path)
End Function

'判断文件夹是否存在
Function IsfolderExists(folder)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.folderExists(folder) Then
        IsfolderExists = True
    Else
        IsfolderExists = False
    End If
End Function


Rem 此虑数函数无需修改
Function DistributeMSG(MSGid,str_MapName,selectID)
    Dim pFun
    
    For i = 0 To UBound(g_MapList)
        If UCase(g_MapList(i)) = UCase(str_MapName) Then
            If MSGid = 3  Then
                
                Set pFun = GetRef(g_MapPostPtrfun(i))
                Call pFun(MSGid,str_MapName,selectID)
                
            Else
                
                Set pFun = GetRef(g_MapPrePtrfun(i))
                Call pFun(MSGid,str_MapName,selectID)
                
            End If
            Exit For
        End If
    Next
End Function


Function debug_print(str)
    
    SSProcess.MapCallBackFunction "OutputMsg", STR & "    " & Now , 0
    
End Function

'// 检查成果目录是否存在、如果不存在放弃出图
Function CheckReportPath(path_print)
    
    Dim fso
    Set fso = CreateObject("scripting.filesystemobject")
    
    Dim path_thisedb
    path_thisedb = SSProcess.GetSysPathName( 5)
    
    path_print = path_thisedb
    
    b1 = fso.FolderExists(path_print)
    
    If  b1 = False  Then
        
        CheckReportPath = False
    Else
        CheckReportPath = True
        
    End If
    
End Function

'// 获取本工程项目名称
Function GetXMMC(xmmc)
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9410001"
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    
    If geocount <> 1 Then GetXMMC = 0
    Exit Function
    
    xmmc = SSProcess.GetSelGeoValue(0,"[XiangMMC]")
    
    If xmmc = "" Or xmmc = "*" Then Exit Function
    
    GetXMMC = 1
    
End Function

'获取规划许可guid
Function GetXKZXX(ghxkzbh,jzwmc,GHXKZGUID,jzdtguid)
    
    Dim arID(20)
    sql = "SELECT JG_建设工程建筑单体信息属性表.GuiHXKZGUID,JZWMCGUID,GuiHXKZBH FROM JG_建设工程建筑单体信息属性表  WHERE (JG_建设工程建筑单体信息属性表.GHXKZBH = '" & ghxkzbh & "' AND JG_建设工程建筑单体信息属性表.JianZWMC = '" & jzwmc & " ');"
    projectname = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectname
    SSProcess.OpenAccessRecordset projectname, sql
    recordCount = SSProcess.GetAccessRecordCount (projectname, sql )
    If recordCount > 0 Then
        SSProcess.AccessMoveFirst projectname,sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            SSFunc.ScanString values, ",", arID, idCount
            GHXKZGUID = arID(0)
            jzdtguid = arID(1)
            ghxkzbh = arID(2)
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
    
End Function

Function DaHui
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "DEFAULT,标注层,测量控制点,数学基础,水系点,水系线,水系面,居民地点,居民地线,居民地面,交通点,交通线,交通面,管线点,管线线,管线面,境界点,境界线,境界面,地貌点,地貌线,地貌面,植被与土质点,植被与土质线,植被与土质面,更新区域,水系中心线,道路中心线,水系注记,居民地注记,交通注记,管线注记,境界注记,地貌注记,植被注记,待更新区域,工作区域,图例层,房屋面,设施点,设施线,设施面,其它境界线,其它境界面,原始观测点,门牌号,等高线,高程点,三维视角点层,管线设施点,管线设施线,管线设施面,废弃工程线,废弃工程点,废弃工程面,DMTZ,GXYZ,KZD,JMD,DLDW,ZBTZ,SXSS,GCD,DGX,ZDH,ZJ,单位名"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For i = 0 To geocount - 1
        geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        SSProcess.SetObjectAttr geoID, "SSObj_Color", RGB(0,0,0)
    Next
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "DEFAULT,标注层,测量控制点,数学基础,水系点,水系线,水系面,居民地点,居民地线,居民地面,交通点,交通线,交通面,管线点,管线线,管线面,境界点,境界线,境界面,地貌点,地貌线,地貌面,植被与土质点,植被与土质线,植被与土质面,更新区域,水系中心线,道路中心线,水系注记,居民地注记,交通注记,境界注记,地貌注记,植被注记,待更新区域,工作区域,图例层,房屋面,设施点,设施线,设施面,其它境界线,其它境界面,原始观测点,门牌号,等高线,高程点,三维视角点层,管线设施点,管线设施线,管线设施面,废弃工程线,废弃工程点,废弃工程面,DMTZ,GXYZ,KZD,JMD,DLDW,ZBTZ,SXSS,GCD,DGX,ZDH,ZJ,单位名"
    SSProcess.SelectFilter
    notecount = SSProcess.GetSelNoteCount()
    For i1 = 0 To notecount - 1
        id = SSProcess.GetSelNoteValue(i1 ,"SSObj_ID" )
        SSProcess.SetObjectAttr id, "SSObj_Color", RGB(0,0,0)
    Next
End Function

Function DeleteFeature(StartCode,EndCode)
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", ">=", StartCode
    SSProcess.SetSelectCondition "SSObj_Code", "<=", EndCode
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj
End Function

Function DeleteFeatureLayerName(strLayerName)
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", strLayerName
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj
End Function

Function JZDSC
    Const JZDBM = "9510041"
    Const QSMBM = "9410001"
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9510041"
    SSProcess.SelectFilter
    geoecount = SSProcess.GetSelgeoCount
    For i = 0 To geoecount - 1
        SSProcess.DelSelgeo i
    Next
    SSProcess.ClearSelection
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
    SSProcess.SelectFilter
    GeoCount = SSProcess.GetSelGeoCount
    For i = 0 To GeoCount - 1
        AreaPNum = SSProcess.GetSelGeoPointCount(i)
        'Msgbox AreaPNum
        For j = 0 To AreaPNum - 2
            SSProcess.GetSelGeoPoint i, j, x,  y,  z,  ptype,  name
            ids = SSProcess.SearchNearObjIDs(x, y, 0.001, 0, JZDBM, 0 )
            If ids = "" Then
                'Msgbox ids
                SSProcess.CreateNewObjByCode JZDBM
                SSProcess.AddNewObjPoint x, y, 9999, 0, "J" & J + 1
                SSProcess.AddNewObjToSaveObjList
            End If
        Next
    Next
    SSProcess.SaveBufferObjToDatabase
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", JZDBM
    SSProcess.SelectFilter
    SSProcess.ChangeSelectionObjAttr "SSObj_PointType", "2"
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
End Function

Function JZX
    SSProcess.ClearSelection
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
    SSProcess.SelectFilter
    GeoCount = SSProcess.GetSelGeoCount
    For i = 0 To GeoCount - 1
        geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        SSProcess.ChangeCodeCopy geoID,"9510042"
        Maxid = SSProcess.GetGeoMaxID()
        SSProcess.LineCrack Maxid ,0
    Next
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
End Function


Function FChandle()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "==", 1
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    Dim strCopy(1000,1000),strCopyID(10000),strCopyID1(10000),strCopyID2(10000),strCopyID3(10000),strCopyID4(10000)
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ids = SSProcess.SearchInnerObjIDs(id, 2, "9420008", 0)
        artemp = Split(ids,",")
        '面心点坐标
        SSProcess.GetObjectFocusPoint  id , x0,  y0
        '面心点到图廓正东距离
        SSProcess.GetObjectPoint     id, 0, x1, y1, z1, pointtype, name
        SSProcess.GetObjectPoint     id, 1, x2, y2, z2, pointtype, name
        SSProcess.GetObjectPoint     id, 2, x3, y3, z3, pointtype, name
        SSProcess.GetObjectPoint     id, 3, x4, y4, z4, pointtype, name
        Length = (x2 - x1) + (x2 - x0)
        For j = 0 To UBound(artemp)
            LC = SSProcess.GetObjectAttr (artemp(j), "[LC]")
            LCID = SSProcess.GetObjectAttr (artemp(j), "[ID_LC]")
            '获取图廓、楼层、功能区ID
            mdbName = SSProcess.GetProjectFileName   '当前工程工程名
            SSProcess.OpenAccessMdb mdbName
            '功能区            
            strSQL = "SELECT JG_规划功能区属性表.ID from  JG_规划功能区属性表 inner join GeoAreaTB on JG_规划功能区属性表.ID = GeoAreaTB.ID where ([GeoAreaTB].[Mark] Mod 2)<>0 and JG_规划功能区属性表.LC = '" & LC & "'"
            GetSQLRecordAll mdbName, strSQL, arSeletionRecord, nSeletionCount
            If nSeletionCount > 0 Then
                strtemp1 = ""
                For k = 0 To nSeletionCount - 1
                    strtemp = arSeletionRecord(k)
                    If strtemp1 = "" Then
                        strtemp1 = strtemp
                    Else
                        strtemp1 = strtemp1 & "," & strtemp
                    End If
                Next
                strCopyID1(j) = id & "," & artemp(j) & "," & strtemp1
            Else
                strCopyID1(j) = id & "," & artemp(j)
            End If
            '附属功能区
            strSQL = "SELECT JG_规划附属功能区属性表.ID from  JG_规划附属功能区属性表 inner join GeoAreaTB on JG_规划附属功能区属性表.ID = GeoAreaTB.ID where ([GeoAreaTB].[Mark] Mod 2)<>0 and JG_规划附属功能区属性表.LC = '" & LC & "'"
            GetSQLRecordAll mdbName, strSQL, arSeletionRecord, nSeletionCount
            If nSeletionCount > 0 Then
                strtemp1 = ""
                For k = 0 To nSeletionCount - 1
                    strtemp = arSeletionRecord(k)
                    If strtemp1 = "" Then
                        strtemp1 = strtemp
                    Else
                        strtemp1 = strtemp1 & "," & strtemp
                    End If
                Next
                strCopyID2(j) = id & "," & artemp(j) & "," & strtemp1
            Else
                strCopyID2(j) = id & "," & artemp(j)
            End If
            '公用区
            strSQL = "SELECT JG_规划公用区属性表.ID from  JG_规划公用区属性表 inner join GeoAreaTB on JG_规划公用区属性表.ID = GeoAreaTB.ID where ([GeoAreaTB].[Mark] Mod 2)<>0 and JG_规划公用区属性表.LC = '" & LC & "'"
            GetSQLRecordAll mdbName, strSQL, arSeletionRecord, nSeletionCount
            If nSeletionCount > 0 Then
                strtemp1 = ""
                For k = 0 To nSeletionCount - 1
                    strtemp = arSeletionRecord(k)
                    If strtemp1 = "" Then
                        strtemp1 = strtemp
                    Else
                        strtemp1 = strtemp1 & "," & strtemp
                    End If
                Next
                strCopyID3(j) = id & "," & artemp(j) & "," & strtemp1
            Else
                strCopyID3(j) = id & "," & artemp(j)
            End If
            '注记
            strSQL = "SELECT JG_底图注记属性表.ID from  JG_底图注记属性表  where  JG_底图注记属性表.ID_LC = '" & LCID & "'"
            GetSQLRecordAll mdbName, strSQL, arSeletionRecord, nSeletionCount
            If nSeletionCount > 0 Then
                strtemp1 = ""
                For k = 0 To nSeletionCount - 1
                    strtemp = arSeletionRecord(k)
                    If strtemp1 = "" Then
                        strtemp1 = strtemp
                    Else
                        strtemp1 = strtemp1 & "," & strtemp
                    End If
                Next
                strCopyID4(j) = id & "," & artemp(j) & "," & strtemp1
            Else
                strCopyID4(j) = id & "," & artemp(j)
            End If
            strCopyID(j) = strCopyID1(j) & "," & strCopyID2(j) & "," & strCopyID3(j) & "," & strCopyID4(j)
            SSProcess.CloseAccessMdb mdbName
        Next
    Next
    
    '粘帖图廓
    For j = 0 To UBound(artemp)
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_ID", "==", strCopyID (j)
        SSProcess.SelectFilter
        SSProcess.SelectionObjToClipBoard
        SSProcess.AddClipBoardObjToMap Length * (j + 1), 0
    Next
    
    '删除原数据
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_ID", "==", 1
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        ids = SSProcess.SearchInnerObjIDs(id, 10, "", 0)
        arids = Split(ids,",")
        For j = 0 To UBound(arids)
            SSProcess.DeleteObject arids(j)
        Next
        SSProcess.DeleteObject id
        SSProcess.RefreshView
    Next
    
End Function


Function reset
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9420008,9420025"
    SSProcess.SelectFilter
    SSProcess.UpdateObjAttrByFeatureCode "FeatureCodeTB_500", "Feature.Code=SSObj_Code", "SSObj_Color=Feature.LineColor,SSObj_LineWidth=Feature.LineWidth,SSObj_LayerName=Feature.LayerName,SSObj_Type=Feature.Type"
    
    
End Function

Function TextEXE
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420031
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    Dim strCopy(1000,1000),strCopyID(10000)
    For i = 0 To geoCount - 1
        '注记生成坐标
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        SSProcess.GetObjectPoint id, 0, x0, y0, z0, pointtype, name
        SSProcess.GetObjectPoint id, 1, x1, y1, z1, pointtype, name
        x = (x1 - x0) / 2 + x0
        y = y0 + 10
        '生成注记
        ids = SSProcess.SearchInnerObjIDs(id, 2, "9420008", 0)
        LC = SSProcess.GetObjectAttr (ids, "[LC]")
        CQC = SSProcess.GetObjectAttr (ids, "[CQC]")
        If InStr(LC,"～") = 0 Then
            If CQC <> "屋面层" Then
                artemplc = Split(LC,".")
                If UBound(artemplc) = 0 Then
                    '标准层
                    strText = ""
                    GetLCXX LC,strText
                    strText = strText & "平面图"
                    CreateText strText,x,y,z
                ElseIf UBound(artemplc) > 0 Then
                    '夹层
                    LC = Mid(LC,1,InStr(LC,".") - 1)
                    strText = ""
                    GetLCXX LC,strText
                    strText = strText & "夹层平面图"
                    CreateText strText,x,y,z
                End If
            Else
                strText = "屋面层平面图"
                CreateText strText,x,y,z
            End If
        Else
            artemplc = Split(LC,"～")
            strText = ""
            GetLCXX artemplc(0),strText0
            GetLCXX    artemplc(1),strText1
            strText = strText0 & "～" & strText1 & "平面图"
            CreateText strText,x,y,z
        End If
    Next
End Function


Function NumberChange(Number,BigNumber)
    strNumer = "1,2,3,4,5,6,7,8,9"
    strBigNumber = "一,二,三,四,五,六,七,八,九"
    artempNumber = Split(strNumer,",")
    artempBigNumber = Split(strBigNumber,",")
    For i = 0 To 8
        If  artempNumber(i) = Number  Then
            BigNumber = artempBigNumber(i)
        End If
    Next
End Function

Function CreateText(strText,x,y,z)
    SSProcess.CreateNewObjByClass "0"
    SSProcess.SetNewObjValue "SSObj_FontString", strText
    SSProcess.SetNewObjValue "SSObj_FontWidth", 1000
    SSProcess.SetNewObjValue "SSObj_FontHeight", 1000
    
    SSProcess.AddNewObjPoint x, y, z, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function GetLCXX(LC,strText)
    If Len(LC) = 1 Then
        For j = 0 To Len(LC) - 1
            IntLC = Mid(LC,j + 1,1)
            NumberChange IntLC,BigNumber
            strText = strText & BigNumber & "层"
        Next
    ElseIf Len(LC) > 1 And InStr(LC,"-") = 0 Then
        For j = 0 To Len(LC) - 1
            IntLC = Mid(LC,j + 1,1)
            NumberChange IntLC,BigNumber
            If strText = "" Then
                strText = BigNumber
            ElseIf IntLC <> 0 Then
                strText = strText & "十" & BigNumber & "层"
            ElseIf IntLC = 0 Then
                strText = strText & "十" & "层"
            End If
        Next
        If  Mid(LC,1,1) = 1 Then
            strText = Mid(strText,2,Len(strText))
        End If
    ElseIf Len(LC) > 1 And InStr(LC,"-") = 1 Then
        For j = 1 To Len(LC) - 1
            IntLC = Mid(LC,j + 1,1)
            NumberChange IntLC,BigNumber
            strText = "地下" & BigNumber & "层"
        Next
    End If
    
End Function

Function TKFZ1
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420004
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        
        
        mdbName = SSProcess.GetProjectFileName
        SSProcess.OpenAccessMdb  mdbName
        sql = "select VALUE from PROJECTINFO where KEY='项目名称'"
        GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
        If nSeletionCount > 0 Then
            XMMC = arSeletionRecord(0)
        Else
            XMMC = ""
        End If
        
        sql = "select VALUE from PROJECTINFO where KEY='测量人员'"
        GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
        If nSeletionCount > 0 Then
            HTRY = arSeletionRecord(0)
        Else
            HTRY = ""
        End If
        
        sql = "select VALUE from PROJECTINFO where KEY='检查人员'"
        GetSQLRecordAll mdbName,sql,arSeletionRecord,nSeletionCount
        If nSeletionCount > 0 Then
            JCRY = arSeletionRecord(0)
        Else
            JCRY = ""
        End If
        
        strtemp = XMMC & "," & HTRY & "," & JCRY
        SSProcess.CloseAccessMdb mdbName
        
        SSProcess.SetObjectAttr id, "SSObj_DataMark", strtemp
    Next
End Function

Function TKFZ2(ByRef mark)
    mark = True
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420031
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        
        ids = SSProcess.SearchInnerObjIDs(id, 2, "9420004", 0)
        
        FWJG = SSProcess.GetObjectAttr (ids, "[FWJG]")
        ZRZH = SSProcess.GetObjectAttr (ids, "[ZRZH]")
        ZCS = SSProcess.GetObjectAttr (ids, "[ZCS]")
        FWZL = SSProcess.GetObjectAttr (ids, "[FWZL]")
        
        idsList = Split(ids,",")
        'if ubound(idsList)>0 then msgbox "出图位置有重叠自然幢，请确认数据是否正确！":mark=false:exit function
        strtemp = SSProcess.GetObjectAttr (idsList(0), "SSObj_DataMark")
        artemp = Split(strtemp,",")
        If UBound(artemp) < 0 Then
            ReDim artemp(2)
            artemp(0) = ""
            artemp(1) = ""
            artemp(2) = ""
        End If
        SSProcess.SetObjectAttr id, "[FWJG]", FWJG
        SSProcess.SetObjectAttr id, "[ZRZH]", ZRZH
        SSProcess.SetObjectAttr id, "[ZCS]", ZCS
        SSProcess.SetObjectAttr id, "[FWZL]", FWZL
        SSProcess.SetObjectAttr id, "[XiangMMC]", artemp(0)
        SSProcess.SetObjectAttr id, "[HuiTY]", artemp(1)
        SSProcess.SetObjectAttr id, "[JianCY]", artemp(2)
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.View.Extend", 0
        '图形重新生成
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    Next
End Function

Function FCTDeleteLC
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9420003,9420008,9410001"
    SSProcess.SelectFilter
    SSProcess.DeleteSelectionObj
End Function

Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    'SQL语句
    sql = StrSqlStatement
    '打开记录集
    SSProcess.OpenAccessRecordset mdbName, sql
    '获取记录总数
    RecordCount = SSProcess.GetAccessRecordCount (mdbName, sql)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim arSQLRecord(RecordCount)
        '将记录游标移到第一行
        SSProcess.AccessMoveFirst mdbName, sql
        iRecordCount = 0
        '浏览记录
        While SSProcess.AccessIsEOF (mdbName, sql) = 0
            fields = ""
            values = ""
            '获取当前记录内容
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            arSQLRecord(iRecordCount) = values                                        '查询记录
            iRecordCount = iRecordCount + 1                                                    '查询记录数
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, sql
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, sql
End Function

'分层图图例
Function CreateKEY
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420031
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            SSProcess.GetObjectPoint TKID, 1, x, y, z, pointtype, name
            ids = SSProcess.SearchInnerObjIDs(TKID , 10 ,"9420021,9420022,9420023", 0)
            ZGNQMC = ""
            ZDrawCode = ""
            If ids <> "" Then
                SSFunc.ScanString ids, ",", vArray, nCount
                For j = 0 To nCount - 1
                    DrawCode = SSProcess.GetObjectAttr(vArray(j), "SSObj_Code")
                    DrawColor = SSProcess.GetObjectAttr(vArray(j), "SSObj_Color")
                    GNQMC = SSProcess.GetObjectAttr(vArray(j), "[MC]")
                    If GNQMC <> "" Then
                        If ZGNQMC = "" Then
                            ZDrawCode = DrawCode
                            ZGNQMC = GNQMC
                            ZDrawColor = DrawColor
                        Else
                            If Replace(ZGNQMC,GNQMC,"") = ZGNQMC Then
                                ZGNQMC = ZGNQMC & "," & GNQMC
                                ZDrawCode = ZDrawCode & "," & DrawCode
                                ZDrawColor = ZDrawColor & "," & DrawColor
                            End If
                        End If
                    End If
                Next
            End If
            
            LvDiTuLi x - 21,y,ZGNQMC,TKID,ZDrawCode,ZDrawColor
        Next
    End If
End Function

'分层图图例
Function LvDiTuLi(x0,y0,ZGNQMC,polygonID,ZDrawCode,ZDrawColor)
    wid1 = 228
    heig1 = 286
    wid2 = 228
    heig2 = 286
    SSFunc.ScanString ZGNQMC, ",", cvArray1, count5
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    '竖线
    makeLine x0,y0,x0,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    'makeLine x0+0.2,y0+0.2,x0+0.2,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
    
    makeLine x0 + 21,y0,x0 + 21,y0 + count5 * 2 + 2.5, 1,"RGB(255,255,255)", polygonID
    makeLine x0 + 8,y0,x0 + 8,y0 + count5 * 2 + 2.5, 1,"RGB(255,255,255)", polygonID
    'makeLine x0+16.8,y0+0.2,x0+16.8,y0+count5*2+2.3, 1,"RGB(255,255,255)", polygonID
    '横线
    'makeLine x0+0.2,y0+0.2,x0+16.8,y0+0.2,1, "RGB(255,255,255)", polygonID
    makeLine x0,y0,x0 + 21,y0,1, "RGB(255,255,255)", polygonID
    'makeLine x0+0.2,y0+count5*2+2.3 ,x0+16.8,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
    makeLine x0,y0 + count5 * 2 + 2.5,x0 + 21,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    makeNote x0 + 2.5,y0 + count5 * 2 + 1.5 , 0, "RGB(255,255,255)", wid2, heig2, "图例",polygonID
    makeNote x0 + 12.5,y0 + count5 * 2 + 1.5 , 0, "RGB(255,255,255)", wid2, heig2, "备注",polygonID
    
    For j = 0 To count5 - 1
        '竖线
        makeArea2 x0 + 1,y0 + j * 2 + 0.7,x0 + 7,y0 + j * 2 + 0.7,x0 + 7,y0 + j * 2 + 2.3,x0 + 1,y0 + j * 2 + 2.3,arDrawCode(j), arDrawColor(j), polygonID,"YT",cvArray1(j)
        makeLine x0,y0 + j * 2 + 2.5,x0 + 21,y0 + j * 2 + 2.5, 1,"RGB(255,255,255)", polygonID
        makeNote x0 + 10,y0 + 1.5 + j * 2, 0, "RGB(255,255,255)", wid2, heig2, cvArray1(j),polygonID
    Next
End Function

'竣工规划总平图,规划复核图
Function CreateKEYZPT
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9420034,9420035"
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            DaYBL = SSProcess.GetSelGeoValue( i, "[DaYBL]" )
            SSProcess.GetObjectPoint TKID, 1, x, y, z, pointtype, name
            ids = SSProcess.SearchInnerObjIDs(TKID , 10 ,"JZ014,FX002,9410001,9410011,9410021,9410031,9410041,9410051,9410061,9410071,9410091,9410101,9410104,9410105,9420005,9420006", 0)
            If ids <> "" Then
                SSFunc.ScanString ids, ",", vArray, nCount
                ZDrawCode = ""
                For j = 0 To nCount - 1
                    DrawCode = SSProcess.GetObjectAttr(vArray(j), "SSObj_Code")
                    DrawColor = SSProcess.GetObjectAttr(vArray(j), "SSObj_Color")
                    DrawName = SSProcess.GetFeatureCodeInfo (DrawCode,"ObjectName")
                    If ZDrawCode = "" Then
                        ZDrawCode = DrawCode
                        ZDrawColor = DrawColor
                        ZDrawName = DrawName
                    Else
                        If Replace(ZDrawCode,DrawCode,"") = ZDrawCode Then
                            ZDrawCode = ZDrawCode & "," & DrawCode
                            ZDrawColor = ZDrawColor & "," & DrawColor
                            ZDrawName = ZDrawName & "," & DrawName
                        End If
                    End If
                Next
            End If
            LvDiTuLiZPT x - 21,y,TKID,ZDrawCode,ZDrawColor,ZDrawName,500
        Next
    End If
    SSProcess.SetMapScale "500"
End Function
'竣工规划总平图
Function LvDiTuLiZPT(x0,y0,polygonID,ZDrawCode,ZDrawColor,ZDrawName,DaYBL)
    wid1 = (228 * 500) / DaYBL
    heig1 = (286 * 500) / DaYBL
    wid2 = (228 * 500) / DaYBL
    heig2 = (286 * 500) / DaYBL
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    arDrawName = Split(ZDrawName,",")
    count5 = UBound(arDrawCode) + 2
    '竖线
    makeLine x0,y0,x0,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    'makeLine x0+0.2,y0+0.2,x0+0.2,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
    
    makeLine x0 + 21,y0,x0 + 21,y0 + count5 * 2 + 2.5, 1,"RGB(255,255,255)", polygonID
    'makeLine x0+8,y0,x0+8,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
    'makeLine x0+16.8,y0+0.2,x0+16.8,y0+count5*2+2.3, 1,"RGB(255,255,255)", polygonID
    '横线
    'makeLine x0+0.2,y0+0.2,x0+16.8,y0+0.2,1, "RGB(255,255,255)", polygonID
    makeLine x0,y0,x0 + 21,y0,1, "RGB(255,255,255)", polygonID
    'makeLine x0+0.2,y0+count5*2+2.3 ,x0+16.8,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
    makeLine x0,y0 + count5 * 2 + 2.5,x0 + 21,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    makeNote x0 + 9.5,y0 + count5 * 2 + 1 , 0, "RGB(255,255,255)", wid2, heig2, "图例",polygonID
    
    For j = 0 To UBound(arDrawCode)
        '竖线
        makeLine x0 + 1,y0 + j * 2 + 1.5,x0 + 7,y0 + j * 2 + 1.5,arDrawCode(j), arDrawColor(j), polygonID
        'makeLine x0,y0+j*2+2.5,x0+16,y0+j*2+2.5, 1,"RGB(255,255,255)", polygonID
        makeNote x0 + 10.5,y0 + 1.5 + j * 2, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j),polygonID
    Next
End Function

'基底图
Function CreateKEYJD
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9420032
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            DaYBL = SSProcess.GetSelGeoValue( i, "[DaYBL]" )
            SSProcess.GetObjectPoint TKID, 1, x, y, z, pointtype, name
            ids = SSProcess.SearchInnerObjIDs(TKID , 10 ,"9420025,9420026,9420027", 0)
            If ids <> "" Then
                SSFunc.ScanString ids, ",", vArray, nCount
                ZDrawCode = ""
                For j = 0 To nCount - 1
                    DrawCode = SSProcess.GetObjectAttr(vArray(j), "SSObj_Code")
                    DrawColor = SSProcess.GetObjectAttr(vArray(j), "SSObj_Color")
                    DrawName = SSProcess.GetFeatureCodeInfo (DrawCode,"ObjectName")
                    If ZDrawCode = "" Then
                        ZDrawCode = DrawCode
                        ZDrawColor = DrawColor
                        ZDrawName = DrawName
                    Else
                        If Replace(ZDrawCode,DrawCode,"") = ZDrawCode Then
                            ZDrawCode = ZDrawCode & "," & DrawCode
                            ZDrawColor = ZDrawColor & "," & DrawColor
                            ZDrawName = ZDrawName & "," & DrawName
                        End If
                    End If
                Next
            End If
            LvDiTuLiJD x - 20,y,TKID,ZDrawCode,ZDrawColor,ZDrawName,DaYBL
        Next
    End If
End Function
'基底图
Function LvDiTuLiJD(x0,y0,polygonID,ZDrawCode,ZDrawColor,ZDrawName,DaYBL)
    wid1 = 228 * 500 / DaYBL
    heig1 = 286 * 500 / DaYBL
    wid2 = 228 * 500 / DaYBL
    heig2 = 286 * 500 / DaYBL
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    arDrawName = Split(ZDrawName,",")
    count5 = UBound(arDrawCode) + 2
    '竖线
    makeLine x0,y0,x0,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    'makeLine x0+0.2,y0+0.2,x0+0.2,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
    
    makeLine x0 + 20,y0,x0 + 20,y0 + count5 * 2 + 2.5, 1,"RGB(255,255,255)", polygonID
    'makeLine x0+8,y0,x0+8,y0+count5*2+2.5, 1,"RGB(255,255,255)", polygonID
    'makeLine x0+16.8,y0+0.2,x0+16.8,y0+count5*2+2.3, 1,"RGB(255,255,255)", polygonID
    '横线
    'makeLine x0+0.2,y0+0.2,x0+16.8,y0+0.2,1, "RGB(255,255,255)", polygonID
    makeLine x0,y0,x0 + 16,y0,1, "RGB(255,255,255)", polygonID
    'makeLine x0+0.2,y0+count5*2+2.3 ,x0+16.8,y0+count5*2+2.3,1, "RGB(255,255,255)", polygonID
    makeLine x0,y0 + count5 * 2 + 2.5,x0 + 20,y0 + count5 * 2 + 2.5,1, "RGB(255,255,255)", polygonID
    makeNote x0 + 8,y0 + count5 * 2 + 1 , 0, "RGB(255,255,255)", wid2, heig2, "图例",polygonID
    
    For j = 0 To UBound(arDrawCode)
        '竖线
        makeNote x0 + 1,y0 + 1.5 + j * 2, 0, "RGB(255,255,255)", wid2, heig2, arDrawName(j) & ":",polygonID
        makeArea x0 + 10,y0 + j * 2 + 0.7,x0 + 17,y0 + j * 2 + 0.7,x0 + 17,y0 + j * 2 + 2.3,x0 + 10,y0 + j * 2 + 2.3,arDrawCode(j), arDrawColor(j), polygonID
    Next
End Function

Function makePoint(x,y,code,color,polygonID)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function makeLine(x1,y1,x2,y2,code, color, polygonID)
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

Function makeArea(x1,y1,x2,y2,x3,y3,x4,y4,code,color,polygonID)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function makeArea1(x1,y1,x2,y2,x3,y3,x4,y4,code,color,polygonID)
    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function makeArea2(x1,y1,x2,y2,x3,y3,x4,y4,code,color,polygonID,field,fieldvalue)
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工测量成果图图廓信息"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x2, y2, 0, 0, ""
    SSProcess.AddNewObjPoint x3, y3, 0, 0, ""
    SSProcess.AddNewObjPoint x4,y4, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
    maxID = SSProcess.GetGeoMaxID
    SSProcess.SetObjectAttr maxID, "[" & field & "]", fieldvalue
    
End Function

Function makeNote(x, y, code, color, width, height, fontString,polygonID)
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

Dim adoConnection
Function InitDB()
    accessName = SSProcess.GetProjectFileName
    Set adoConnection = CreateObject("adodb.connection")
    strcon = "DBQ=" & accessName & ";DRIVER={Microsoft Access Driver (*.mdb)};"
    adoConnection.Open strcon
End Function

'//关库
Function ReleaseDB()
    adoConnection.Close
    Set adoConnection = Nothing
End Function
'//判断表是否存在
Function IsTableExits(ByVal  strMdbName,ByVal  strTableName_s)
    strMdbName = SSProcess.GetProjectFileName
    IsTableExits = False
    strTableName_s = UCase(strTableName_s)
    '判断文件DB文件是否存在
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.fileExists(strMdbName) = False Then Exit Function
    '获取DB文件后缀名
    Set f = fso.getfile(strMdbName)
    dbType = fso.GetExtensionName(f)
    Set f = Nothing
    Set fso = Nothing
    '分DB类型查找
    If dbType = "dbf" Then
        strMdbName = Replace(strMdbName,"/","\")
        ipos = InStrRev(strMdbName,"\")
        strMdbName = Left(strMdbName,ipos)
        strcon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMdbName + ";Extended Properties=dBASE IV;User ID=;Password="
    Else
        strcon = "DBQ=" & strMdbName & ";DRIVER={Microsoft Access Driver (*.mdb)};"
    End If
    '开库
    '''set adoConnection=createobject("adodb.connection")
    '''adoConnection.Open strcon
    Set rsSchema = adoConnection.OpenSchema(20)
    '获取当前DB文件所有表
    strAllTableName = ""
    Do While Not rsSchema.EOF
        strTableName = "" & UCase (rsSchema.Fields("TABLE_NAME")) & ""
        If Left(strTableName,4) <> "MSYS" Then
            If strTableName_s = strTableName Then IsTableExits = True
            Exit Do
        End If
        rsSchema.MoveNext
    Loop
    rsSchema.Close
    Set rsSchema = Nothing
    If IsTableExits = False Then addloginfo "【" & strTableName_s & "】表在edb中不存在"
    ''adoConnection.Close
    ''Set adoConnection = Nothing
End Function

Function GetProjectTableList(ByVal strTableName,ByVal strFields,ByVal strAddCondition,ByVal strTableType,ByVal strGeoType,ByRef rs(),ByRef fieldCount)
    GetProjectTableList = 0
    values = ""
    rsCount = 0
    fieldCount = 0
    If strTableName = "" Or strFields = "" Then Exit Function
    If IsTableExits("",strTableName) = False Then Exit Function
    'strFields=GetTableAllFields ("", strTableName, strFields)
    If  strFields = "" Then Exit Function
    '设置地物类型
    If strGeoType = "0" Then
        GeoType = "GeoPointTB"
    ElseIf strGeoType = "1" Then
        GeoType = "GeoLineTB"
    ElseIf strGeoType = "2" Then
        GeoType = "GeoAreaTB"
    ElseIf strGeoType = "3" Then
        GeoType = "MarkNoteTB"
    Else
        GeoType = "GeoAreaTB"
    End If
    If strTableType = "SpatialData" Then
        strCondition = " (" & GeoType & ".Mark Mod 2)<>0"
        If strAddCondition <> "" Then strCondition = " (" & GeoType & ".Mark Mod 2)<>0 and " & strAddCondition & ""
        sql = "select  " & strFields & " from " & strTableName & "  INNER JOIN " & GeoType & " ON " & strTableName & ".ID = " & GeoType & ".ID WHERE " & strCondition & ""
    Else
        If strAddCondition <> "" Then
            strCondition = strAddCondition
            sql = "select  " & strFields & " from " & strTableName & "  WHERE  " & strCondition & ""
        Else
            sql = "select  " & strFields & " from " & strTableName & ""
        End If
    End If
    
    ''addloginfo sql
    'if instr(sql,"scpcjzmj")>0 then  addloginfo sql
    '获取当前工程edb表记录
    AccessName = SSProcess.GetProjectFileName
    '判断表是否存在
    'if  IsTableExits(AccessName,strTableName)=false then exit function 
    'set adoConnection=createobject("adodb.connection")
    'strcon="DBQ="& AccessName &";DRIVER={Microsoft Access Driver (*.mdb)};"  
    'adoConnection.Open strcon
    Set adoRs = CreateObject("ADODB.recordset")
    count = 0
    adoRs.cursorLocation = 3
    adoRs.cursorType = 3
    adoRs.open sql,adoConnection,3,3
    rcdCount = adoRs.RecordCount
    fieldCount = adoRs.Fields.Count
    ReDim rs(rcdCount,fieldCount)
    'erase rs
    While adoRs.Eof = False
        nowValues = ""
        For i = 0 To fieldCount - 1
            value = adoRs(i)
            If IsNull(value) Then value = ""
            value = Replace(value,",","，")
            rs(rsCount,i) = value
        Next
        rsCount = rsCount + 1
        adoRs.MoveNext
    WEnd
    adoRs.Close
    Set adoRs = Nothing
    'adoConnection.Close
    'Set adoConnection = Nothing
    GetProjectTableList = rsCount
End Function
Function YDHXYDMJ(YongDMJ)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9410001"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    If geocount = 1 Then
        For i = 0 To geocount - 1
            id = SSProcess.GetSelGeoValue(i, "SSObj_ID")
            YongDMJ = SSProcess.GetObjectAttr (id, "SSObj_Area")
        Next
    End If
    YDHXYDMJ = YongDMJ
End Function

Function GetSubArea(cellList,cellCount,ByVal scArea,ByVal ghArea,ByVal numberDigit,ByVal startCol)
    If IsNumeric(numberDigit) = False Then numberDigit = 2
    scArea = GetFormatNumber(scArea,numberDigit)
    ghArea = GetFormatNumber(ghArea,numberDigit)
    subNum = CDbl(scArea) - CDbl(ghArea)
    subNum = GetFormatNumber(subNum,numberDigit)'差值-建筑基底面积
    If scArea = "0.00" Or scArea = "0" Then scArea = ""
    If ghArea = "0.00" Or ghArea = "0" Then ghArea = ""
    If subNum = "0.00" Or subNum = "0" Then subNum = ""
    
    If startCol = 2 Then      cellValue = scArea & "||" & ghArea & "||" & subNum & "||" & ""  Else     cellValue = scArea & "||" & ghArea & "||" & subNum
    ReDim Preserve cellList(cellCount)
    cellList(cellCount) = cellValue
    cellCount = cellCount + 1
End Function
Function OutputTable11( )
    cellCount = 0
    ReDim cellList(cellCount)
    '**************************************************************总用地面积
    ydhxTableName = "JG_用地红线信息属性表"
    fields = "GuiHSPZYDMJ"
    listCount = GetProjectTableList (ydhxTableName,"GuiHSPZYDMJ","","SpatialData","1",list,fieldCount)
    If listCount = 1 Then gh_YongDMJ = list(0,0)
    gh_YongDMJ = GetFormatNumber(gh_YongDMJ,2)'规划-总用地面积
    sc_YongDMJ = YDHXYDMJ(YongDMJ)
    If sc_YongDMJ <> "" Then sc_YongDMJ = GetFormatNumber(sc_YongDMJ,2)
    GetSubArea cellList,cellCount, sc_YongDMJ, gh_YongDMJ,2,1
    
    '**************************************************************总建筑面积
    zrzCount = GetProjectTableList ("FC_自然幢信息属性表","sum(SCJZMJ)","","SpatialData","2",zrzList,fieldCount)
    If zrzCount = 1 Then sc_SCJZMJ = zrzList(0,0)
    sc_SCJZMJ = GetFormatNumber(sc_SCJZMJ,2)'实测-总建筑面积
    
    ghxkTableName = "JG_建设工程规划许可证信息属性表"
    'exCondition="YDHXGUID In (select YDHXGUID from "&ydhxTableName&"  INNER JOIN GeoLineTB ON "&ydhxTableName&".ID = GeoLineTB.ID WHERE (GeoLineTB.Mark Mod 2)<>0)"
    
    exCondition = "ID>0"
    ghxkCount = GetProjectTableList (ghxkTableName,"sum(GuiHSPZJZMJ)",exCondition,"","",ghxkList,fieldCount)
    If ghxkCount = 1 Then gh_SCJZMJ = ghxkList(0,0)
    gh_SCJZMJ = GetFormatNumber(gh_SCJZMJ,2)'规划-总建筑面积
    GetSubArea cellList,cellCount, sc_SCJZMJ, gh_SCJZMJ,2,1
    
    iniRow = 1
    iniCol = 1
    startRow = iniRow
    startCol = iniCol
    ghgnqTableName = "JG_规划功能区属性表"
    '**************************************************************地上建筑面积
    GetGnqAreaList cellList,cellCount, ghxkTableName, ghgnqTableName, "int(sjcs)>0","GuiHSPDSJZMJ",exCondition,copyCount,sumDsGNQMJ
    '复制地上功能区
    startRow = iniRow + 2
    startCol = iniCol + 1
    
    '**************************************************************地下建筑面积
    startRow = startRow + copyCount + 1
    GetGnqAreaList cellList,cellCount, ghxkTableName, ghgnqTableName, "int(sjcs)<0","GuiHSPDXJZMJ",exCondition,copyCount1,sumDxGNQMJ
    '复制地下功能区
    
    
    
    '**************************************************************建筑基底面积
    jdCount = GetProjectTableList ("JG_建筑物基底面属性表","sum(JDMJ)","","SpatialData","2",jdList,fieldCount)
    If jdCount = 1 Then sc_JDMJ = jdList(0,0)
    sc_JDMJ = GetFormatNumber(sc_JDMJ,2)'实测-建筑基底面积
    ghxkCount = GetProjectTableList (ghxkTableName,"sum(GuiHSPJDMJ),sum(GuiHSPRJL),sum(GuiHSPJZMD),sum(GuiHSPLHL),sum(ZpsJZMJ),sum(ScZZHS),sum(GhZZHS)",exCondition,"","",ghxkList,fieldCount)
    If ghxkCount = 1 Then
        gh_JDMJ = ghxkList(0,0)
        gh_GuiHSPRJL = ghxkList(0,1)
        gh_GuiHSPJZMD = ghxkList(0,2)
        gh_GuiHSPLHL = ghxkList(0,3)
        gh_ZpsJZMJ = ghxkList(0,4)
        ScZZHS = ghxkList(0,5)
        GhZZHS = ghxkList(0,6)
    End If
    gh_JDMJ = GetFormatNumber(gh_JDMJ,2)'规划-建筑基底面积
    GetSubArea cellList,cellCount, sc_JDMJ, gh_JDMJ,2,1
    
    ldCount = GetProjectTableList ("GH_绿化要素属性表","sum(LHMJ)","ID>0","","",sclhmjList,fieldCount)
    If ldCount = 1 Then sc_lhmj = sclhmjList(0,0)
    gh_lhmj = ""
    GetSubArea cellList,cellCount, sc_lhmj, gh_lhmj,2,1'绿地面积
    
    If  sc_YongDMJ = 0 Then sc_Rjl = 0 Else    sc_Rjl = sumDsGNQMJ / sc_YongDMJ
    GetSubArea cellList,cellCount, sc_Rjl, gh_GuiHSPRJL,2,1'容积率
    
    If  sc_YongDMJ = 0 Then sc_Jzmd = 0 Else    sc_Jzmd = (sc_JDMJ / sc_YongDMJ) * 100
    GetSubArea cellList,cellCount, sc_Jzmd, gh_GuiHSPJZMD,2,1'建筑密度
    
    ldCount = GetProjectTableList ("GH_绿化要素属性表","sum(LHMJ/ZSBL)","ID>0","","",sclhYdmjList,fieldCount)
    If ldCount = 1 Then sc_lhYdmj = sclhYdmjList(0,0)
    If  sc_YongDMJ = 0 Then sc_lhl = 0 Else    sc_lhl = (sc_lhYdmj / sc_YongDMJ) * 100
    GetSubArea cellList,cellCount, sc_lhl, gh_GuiHSPLHL,2,1'绿地率
    
    sc_ZpsJZMJ = ""
    If gh_ZpsJZMJ = 0 Then gh_ZpsJZMJ = ""
    GetSubArea cellList,cellCount, sc_ZpsJZMJ, gh_ZpsJZMJ,2,1'装配式建筑面积
    
    cwTableName = "CWSCXX"
    cwCount = GetProjectTableList (cwTableName,"sum(DSCWSL)+sum(DXCWSL),sum(DSCWSL),sum(DXCWSL)","CWLX='普通机动车位'","","",cwList,fieldCount)
    If  cwCount = 1 Then    sc_Jdcw = cwList(0,0)
    sc_ds_Jdcw = cwList(0,1)
    sc_dx_Jdcw = cwList(0,2)
    
    ghcwTableName = "CWGHXX"
    cwCount = GetProjectTableList (ghcwTableName,"sum(DSCWSL)+sum(DXCWSL),sum(DSCWSL),sum(DXCWSL)","CWLX='普通机动车位'","","",ghcwList,fieldCount)
    If  cwCount = 1 Then    gh_Jdcw = ghcwList(0,0)
    gh_ds_Jdcw = ghcwList(0,1)
    gh_dx_Jdcw = ghcwList(0,2)
    
    GetSubArea cellList,cellCount, sc_Jdcw, gh_Jdcw,0,1'机动车位
    GetSubArea cellList,cellCount, sc_ds_Jdcw, gh_ds_Jdcw,0,2'地上机动车位
    GetSubArea cellList,cellCount, sc_dx_Jdcw, gh_dx_Jdcw,0,2'地下机动车位
    GetSubArea cellList,cellCount, ScZZHS, GhZZHS,0,1'住宅户数
    
    cwCount = GetProjectTableList (cwTableName,"sum(DSCWSL)+sum(DXCWSL)","CWLX='非机动车位'","","",cwList,fieldCount)
    If  cwCount = 1 Then    sc_Fjdcw = cwList(0,0)
    ghcwCount = GetProjectTableList (ghcwTableName,"sum(DSCWSL)+sum(DXCWSL)","CWLX='非机动车位'","","",ghcwList,fieldCount)
    If  ghcwCount = 1 Then    gh_Fjdcw = ghcwList(0,0)
    GetSubArea cellList,cellCount, sc_Fjdcw, gh_Fjdcw,0,1'非机动车位
    
    
End Function

'//获取功能区分类面积
Function GetGnqAreaList(cellList,cellCount,ByVal ghxkTableName,ByVal ghgnqTableName,ByVal strConditon,ByVal field,exCondition,copyCount,sc_GNQMJ)
    copyCount = 0
    '**************************************************************建筑面积
    ghgnqCount = GetProjectTableList (ghgnqTableName,"SUM(GNQMJ)",strConditon,"SpatialData","2",ghgnqList,fieldCount)
    If ghgnqCount = 1  Then sc_GNQMJ = ghgnqList(0,0)
    sc_GNQMJ = GetFormatNumber(sc_GNQMJ,2)'实测-建筑面积
    
    ghxkCount = GetProjectTableList (ghxkTableName,"sum(" & field & ")",exCondition,"","",ghxkList,fieldCount)
    If ghxkCount = 1 Then gh_GNQMJ = ghxkList(0,0)
    gh_GNQMJ = GetFormatNumber(gh_GNQMJ,2)'规划-建筑面积
    GetSubArea cellList,cellCount, sc_GNQMJ, gh_GNQMJ,2,2
    '**************************************************************建筑面积-各功能区面积
    ghgnqCount = GetProjectTableList (ghgnqTableName,"SUM(JZMJ),YT","" & strConditon & " group by YT","SpatialData","2",ghgnqList,fieldCount)
    
    ghldxxCount = GetProjectTableList ("GHLDXX","SUM(JZMJ),GHYT","GHYT<>'' group by GHYT","AttributeData","0",ghldxxList,ghldxxfieldCount)
    
    If ghgnqCount > 0 Then
        For i = 0 To ghgnqCount - 1
            sc_gnq_GNQMJ = ghgnqList(i,0)
            sc_gnq_GNQMJ = GetFormatNumber(sc_gnq_GNQMJ,2)
            gnqName = ghgnqList(i,1)
            ghldxx_gnqmj = ""
            If ghldxxCount > 0 Then
                For i1 = 0 To ghldxxCount - 1
                    If ghldxxList(i1,1) = gnqName Then    ghldxx_gnqmj = ghldxxList(i1,0)
                    ghldxx_gnqmj = GetFormatNumber(ghldxx_gnqmj,2)
                Next
            End If
            If sc_gnq_GNQMJ = "" Then sc_gnq_GNQMJ = 0
            If ghldxx_gnqmj = "" Then ghldxx_gnqmj = 0
            change_gnqmj = GetFormatNumber(sc_gnq_GNQMJ - ghldxx_gnqmj,2)
            If sc_gnq_GNQMJ = "0.00" Or sc_gnq_GNQMJ = "0" Then sc_gnq_GNQMJ = ""
            If ghldxx_gnqmj = "0.00" Or ghldxx_gnqmj = "0" Then ghldxx_gnqmj = ""
            If change_gnqmj = "0.00" Or change_gnqmj = "0" Then change_gnqmj = ""
            cellValue = gnqName & "||" & sc_gnq_GNQMJ & "||" & ghldxx_gnqmj & "||" & change_gnqmj
            ReDim Preserve cellList(cellCount)
            cellList(cellCount) = cellValue
            cellCount = cellCount + 1
            copyCount = copyCount + 1
        Next
    Else
        cellValue = gnqName & "||" & "" & "||" & "" & "||" & ""
        ReDim Preserve cellList(cellCount)
        cellList(cellCount) = cellValue
        cellCount = cellCount + 1
    End If
End Function

Function JGZPTKEY(ByVal TKID)
    
    InitDB()
    sSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 9410001
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    
    xmmc = SSProcess.GetSelGeoValue( 0, "[XiangMMC]" )
    '找JG_规划功能区属性表的yt 字段 去重 得到数量a
    ydhxTableName = "ZYJJZBMJHZB"
    GNQCount = GetProjectTableList ("ZYJJZBMJHZB","distinct LXMC"," ID>0","","",list,fieldCount)
    
    '功能区名称 数组
    '总行数= a+13
    ZHS = GNQCount + 12
    '画外框
    '获取图廓坐标位置
    
    SSProcess.GetObjectPoint TKID, 2, x, y, z, pointtype, name
    x1 = x - 10
    y1 = y - 10
    ztmc = "黑体"
    ztdx = 150
    ztkd = 200
    yzl = y1 - 7 - 10 - GNQCount * 2 - 24 - 1
    
    '表头
    makeNote1 x1 - 10, y1 + 2, code, color, ztdx, ztkd, xmmc & "项目规划核实表",TKID,ztmc
    makeArea1 x1 + 3,y1,x1 - 17,y1,x1 - 17,yzl + 1,x1 + 3,yzl + 1,1,color,TKID
    
    '第一个竖线
    
    makeLine x1 - 15,y1,x1 - 15,yzl + 1,1, color, TKID
    '第一个横线
    makeLine x1 - 17,y1 - 3,x1 + 3,y1 - 3,1, color, TKID
    '序号
    makeNote1 x1 - 16.3, y1 - 0.5, code, color, ztdx, ztkd, "序",TKID,ztmc
    makeNote1 x1 - 16.3, y1 - 1.5, code, color, ztdx, ztkd, "号",TKID,ztmc
    '第二个竖线
    makeLine x1 - 3,y1,x1 - 3,yzl + 10,1, color, TKID
    makeNote1 x1 - 14, y1 - 1.5, code, color, ztdx, ztkd, "用地信息或技术指标名称",TKID,ztmc
    makeNote1 x1 - 1.5, y1 - 1.5, code, color, ztdx, ztkd, "核实数据",TKID,ztmc
    
    '第二个横线
    makeLine x1 - 17,y1 - 5,x1 + 3,y1 - 5,1, color, TKID
    makeNote1 x1 - 16, y1 - 3.5, code, color, ztdx, ztkd, "1",TKID,ztmc
    makeNote1 x1 - 10, y1 - 3.5, code, color, ztdx, ztkd, "容积率",TKID,ztmc
    
    
    ghgnqCount = GetProjectTableList ("JGSCHZXX","DSJZMJ"," ID>0","","",ghgnqList,fieldCount)
    If ghgnqCount = 1  Then sumDsGNQMJ = ghgnqList(0,0)
    If sumDsGNQMJ = "" Then sumDsGNQMJ = 0
    sumDsGNQMJ = GetFormatNumber(sumDsGNQMJ,2)'实测-建筑面积
    sc_YongDMJ = YDHXYDMJ(YongDMJ)
    listCount = GetProjectTableList ("JGSCHZXX","RJV","ID>0","","",RJLlist,fieldCount)'容积率
    If listCount = 1 Then sc_Rjl = RJLlist(0,0)
    makeNote1 x1 - 1.5, y1 - 4, code, color, ztdx, ztkd, sc_Rjl,TKID,ztmc
    '第三个横线
    makeLine x1 - 17,y1 - 7,x1 + 3,y1 - 7,1, color, TKID
    makeNote1 x1 - 16, y1 - 6, code, color, ztdx, ztkd, "2",TKID,ztmc
    makeNote1 x1 - 12, y1 - 5.5, code, color, ztdx, ztkd, "计算容积率建筑面积",TKID,ztmc
    makeNote1 x1 - 1.5, y1 - 5.5, code, color, ztdx, ztkd, sumDsGNQMJ,TKID,ztmc
    '第四个横线
    makeLine x1 - 15,y1 - 18,x1 + 3,y1 - 18,1, color, TKID
    makeNote1 x1 - 16, y1 - 12 - GNQCount, code, color, ztdx, ztkd, "3",TKID,ztmc
    
    '第1个短横线
    makeLine x1 - 15,y1 - 9,x1 + 3,y1 - 9,1, color, TKID
    makeNote1 x1 - 10.5, y1 - 7.5, code, color, ztdx, ztkd, "总建筑面积",TKID,ztmc
    zrzCount = GetProjectTableList ("JGSCHZXX","JZMJ","ID>0","","",zrzList,fieldCount)
    If zrzCount = 1 Then sc_SCJZMJ = zrzList(0,0)
    sc_SCJZMJ = GetFormatNumber(sc_SCJZMJ,2)'实测-总建筑面积
    makeNote1 x1 - 1.5, y1 - 7.5, code, color, ztdx, ztkd, sc_SCJZMJ,TKID,ztmc
    
    '第2个短横线
    makeLine x1 - 9,y1 - 13,x1 + 3,y1 - 13,1, color, TKID
    makeNote1 x1 - 8, y1 - 10, code, color, ztdx, ztkd, "地上建筑",TKID,ztmc
    makeNote1 x1 - 7, y1 - 12, code, color, ztdx, ztkd, "面积",TKID,ztmc
    makeNote1 x1 - 1.5, y1 - 11, code, color, ztdx, ztkd, sumDsGNQMJ,TKID,ztmc
    
    makeNote1 x1 - 8, y1 - 14, code, color, ztdx, ztkd, "地下建筑",TKID,ztmc
    makeNote1 x1 - 7, y1 - 16, code, color, ztdx, ztkd, "面积",TKID,ztmc
    ghgnqCount1 = GetProjectTableList ("JGSCHZXX","DXJZMJ"," ID>0","","",ghgnqList1,fieldCount)
    If ghgnqCount1 = 1  Then sumDsGNQMJ1 = ghgnqList1(0,0)
    sumDsGNQMJ1 = GetFormatNumber(sumDsGNQMJ1,2)'实测-建筑面积
    makeNote1 x1 - 1.5, y1 - 15, code, color, ztdx, ztkd, sumDsGNQMJ1,TKID,ztmc
    makeNote1 x1 - 14, y1 - 12, code, color, ztdx, ztkd, "按空间",TKID,ztmc
    makeNote1 x1 - 14, y1 - 14, code, color, ztdx, ztkd, "位置分类",TKID,ztmc
    
    makeNote1 x1 - 13, y1 - 17 - GNQCount - 2, code, color, ztdx, ztkd, "按使用",TKID,ztmc
    makeNote1 x1 - 13, y1 - 17 - GNQCount - 4, code, color, ztdx, ztkd, "用途分类",TKID,ztmc
    '循环遍历用途
    For j = 0 To GNQCount - 1
        ytname = list(j,0)
        makeLine x1 - 9,y1 - 17 - j * 2 - 3,x1 + 3,y1 - 17 - j * 2 - 3,1, color, TKID
        makeNote1 x1 - 8, y1 - 17 - j * 2 - 1.5, code, color, ztdx, ztkd,ytname ,TKID,ztmc
        '查对应面积
        ytCount = GetProjectTableList (ydhxTableName,"SCJZMJ","LXMC='" & ytname & "'","","",list1,fieldCount1)
        ytmj = list1(0,0)
        makeNote1 x1 - 1.5, y1 - 17 - j * 2 - 1.5, code, color, ztdx, ztkd,ytmj,TKID,ztmc
    Next
    '竖线
    makeLine x1 - 9,y1 - 9,x1 - 9,y1 - 7 - 10 - GNQCount * 2 - 1,1, color, TKID
    '第五个横线
    makeLine x1 - 15,y1 - 7 - 10 - GNQCount * 2 - 1,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 1,1, color, TKID
    '第六个横线
    makeLine x1 - 17,y1 - 7 - 10 - GNQCount * 2 - 5,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 5,1, color, TKID
    makeNote1 x1 - 16, y1 - 7 - 10 - GNQCount * 2 - 3, code, color, ztdx, ztkd, "4",TKID,ztmc
    '短横线
    makeLine x1 - 15,y1 - 7 - 10 - GNQCount * 2 - 3,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 3,1, color, TKID
    makeNote1 x1 - 10, y1 - 7 - 10 - GNQCount * 2 - 1.5, code, color, ztdx, ztkd, "地上车位",TKID,ztmc
    cwTableName = "CWSCXX"
    cwCount = GetProjectTableList ("JGSCHZXX","DSJDCWGS+DXJDCWGS,DSJDCWGS,DXJDCWGS","ID>0","","",cwList,fieldCount)
    If  cwCount = 1 Then
        sc_ds_Jdcw = cwList(0,1)
        sc_dx_Jdcw = cwList(0,2)
        If sc_ds_Jdcw = "" Then sc_ds_Jdcw = 0
        If sc_dx_Jdcw = "" Then sc_dx_Jdcw = 0
        sc_Jdcw = Int(sc_ds_Jdcw) + Int(sc_dx_Jdcw)
    End If
    cwCount = GetProjectTableList ("JGSCHZXX","DSFJDCWGS,DXFJDCWGS","ID>0","","",cwList,fieldCount)
    If  cwCount = 1 Then
        DSFJDCWGS = cwList(0,0)
        DXFJDCWGS = cwList(0,1)
        If DSFJDCWGS = "" Then DSFJDCWGS = 0
        If DXFJDCWGS = "" Then DXFJDCWGS = 0
        sc_Fjdcw = Int(DSFJDCWGS) + Int(DXFJDCWGS)
    End If
    
    dscezsl = Int(sc_ds_Jdcw) + Int(DSFJDCWGS)
    dxcezsl = Int(sc_dx_Jdcw) + Int(DSFJDCWGS)
    makeNote1 x1 - 1.5, y1 - 7 - 10 - GNQCount * 2 - 1.5, code, color, ztdx, ztkd, dscezsl,TKID,ztmc
    makeNote1 x1 - 10, y1 - 7 - 10 - GNQCount * 2 - 3.5, code, color, ztdx, ztkd, "地下车位",TKID,ztmc
    makeNote1 x1 - 1.5, y1 - 7 - 10 - GNQCount * 2 - 3.5, code, color, ztdx, ztkd,dxcezsl,TKID,ztmc
    '第七个横线
    makeLine x1 - 17,y1 - 7 - 10 - GNQCount * 2 - 7,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 7,1, color, TKID
    makeNote1 x1 - 16, y1 - 7 - 10 - GNQCount * 2 - 6, code, color, ztdx, ztkd, "5",TKID,ztmc
    makeNote1 x1 - 10, y1 - 7 - 10 - GNQCount * 2 - 5.5, code, color, ztdx, ztkd, "绿地率",TKID,ztmc
    
    ldCount = GetProjectTableList ("JGSCHZXX","LVL","ID>0","","",sclhYdmjList,fieldCount)
    If ldCount = 1 Then sc_lhYdmj = sclhYdmjList(0,0)
    sc_lhl = GetFormatNumber(sc_lhl,2)
    makeNote1 x1 - 1.5, y1 - 7 - 10 - GNQCount * 2 - 5.5, code, color, ztdx, ztkd, sc_lhYdmj & "%",TKID,ztmc
    '第八个横线
    makeLine x1 - 15,y1 - 7 - 10 - GNQCount * 2 - 9,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 9,1, color, TKID
    makeNote1 x1 - 16, y1 - 7 - 10 - GNQCount * 2 - 11, code, color, ztdx, ztkd, "6",TKID,ztmc
    makeNote1 x1 - 10, y1 - 7 - 10 - GNQCount * 2 - 7.5, code, color, ztdx, ztkd, "土地竣工面积",TKID,ztmc
    '待开放makeNote x1, y1-7-10-GNQCount*2-7, code, color, ztdx, ztkd, "土地竣工面积数据",TKID,ztmc
    '第九个横线
    makeLine x1 - 17,y1 - 7 - 10 - GNQCount * 2 - 15,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 15,1, color, TKID
    '三分之二竖线
    makeLine x1 - 7,y1 - 7 - 10 - GNQCount * 2 - 9,x1 - 7,yzl + 1,1, color, TKID
    makeNote1 x1 - 13, y1 - 7 - 10 - GNQCount * 2 - 11.5, code, color, ztdx, ztkd, "按宗地分类",TKID,ztmc
    '短横线
    makeLine x1 - 7,y1 - 7 - 10 - GNQCount * 2 - 11,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 11,1, color, TKID
    makeNote1 x1 - 6, y1 - 7 - 10 - GNQCount * 2 - 9.5, code, color, ztdx, ztkd, "宗地一",TKID,ztmc
    '宗地有数据取值时 下面放开
    'makeNote x1, y1-7-10-GNQCount*2-9, code, color, ztdx, ztkd, "宗地一数据",TKID,ztmc
    '短横线
    makeLine x1 - 7,y1 - 7 - 10 - GNQCount * 2 - 13,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 13,1, color, TKID
    makeNote1 x1 - 6, y1 - 7 - 10 - GNQCount * 2 - 11.5, code, color, ztdx, ztkd, "宗地二",TKID,ztmc
    '宗地有数据取值时 下面放开
    'makeNote x1, y1-7-10-GNQCount*2-11, code, color, ztdx, ztkd, "宗地二数据",TKID,ztmc
    '短横线
    makeNote1 x1 - 6, y1 - 7 - 10 - GNQCount * 2 - 13.5, code, color, ztdx, ztkd, "宗地三",TKID,ztmc
    '宗地有数据取值时 下面放开
    'makeNote x1, y1-7-10-GNQCount*2-13, code, color, ztdx, ztkd, "宗地二数据",TKID,ztmc
    '第十个横线
    makeLine x1 - 17,y1 - 7 - 10 - GNQCount * 2 - 19,x1 + 3,y1 - 7 - 10 - GNQCount * 2 - 19,1, color, TKID
    makeNote1 x1 - 16, y1 - 7 - 10 - GNQCount * 2 - 17, code, color, ztdx, ztkd, "7",TKID,ztmc
    makeNote1 x1 - 13, y1 - 7 - 10 - GNQCount * 2 - 15.5, code, color, ztdx, ztkd, "不动产权证或",TKID,ztmc
    makeNote1 x1 - 13.5, y1 - 7 - 10 - GNQCount * 2 - 17.5, code, color, ztdx, ztkd, "（土地证）证号",TKID,ztmc
    ''待开放makeNote x1,  y1-7-10-GNQCount*2-16, code, color, ztdx, ztkd, "数据",TKID,ztmc
    '最后一行
    makeNote1 x1 - 16, y1 - 7 - 10 - GNQCount * 2 - 21, code, color, ztdx, ztkd, "8",TKID,ztmc
    makeNote1 x1 - 13, y1 - 7 - 10 - GNQCount * 2 - 21, code, color, ztdx, ztkd, "土地用途",TKID,ztmc
    ''待开放makeNote x1,  y1-7-10-GNQCount*2-19, code, color, ztdx, ztkd, "数据",TKID,ztmc
    '文字
    makeNote1 x1 - 17, y1 - 7 - 10 - GNQCount * 2 - 27, code, color, ztdx, ztkd, "说明：",TKID,ztmc
    makeNote1 x1 - 17, y1 - 7 - 10 - GNQCount * 2 - 29, code, color, ztdx, ztkd, "1、城市市政，指城市电力、水利、给排水等设施。",TKID,ztmc
    makeNote1 x1 - 17, y1 - 7 - 10 - GNQCount * 2 - 31, code, color, ztdx, ztkd, "2、项目配套市政，指为本项目配套的电力、电信、给排水、设备等用房。",TKID,ztmc
    makeNote1 x1 - 17, y1 - 7 - 10 - GNQCount * 2 - 33, code, color, ztdx, ztkd, "3、空格为根据项目的具体情况各分局另行填写的内容。",TKID,ztmc
    makeNote1 x1 - 17, y1 - 7 - 10 - GNQCount * 2 - 35, code, color, ztdx, ztkd, "4、各使用用途的房间其空间位置详见施工平面图。",TKID,ztmc
    makeNote1 x1 - 17, y1 - 7 - 10 - GNQCount * 2 - 37, code, color, ztdx, ztkd, "5、表中面积是按照《建筑工程建筑面积计算和竣工综合测量技术规程》",TKID,ztmc
    makeNote1 x1 - 17, y1 - 7 - 10 - GNQCount * 2 - 39, code, color, ztdx, ztkd, "（DB33/T 1152-2018）进行核实",TKID,ztmc
    makeNote1 x1 - 17, y1 - 7 - 10 - GNQCount * 2 - 41, code, color, ztdx, ztkd, "6、建筑物内挑空、镂空、夹层等未表示，层次与总平面图保持一致",TKID,ztmc
    
    '画内框
    
    ReleaseDB()
End Function

Function makeNote1(x, y, code, color, width, height, fontString,polygonID,ztmc)
    SSProcess.CreateNewObj 3
    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontInterval", "80"
    SSProcess.SetNewObjValue "SSObj_FontString", fontString
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_FontName", ztmc
    SSProcess.SetNewObjValue "SSObj_LayerName", "竣工图廓"
    SSProcess.SetNewObjValue "SSObj_FontInterval", "8"
    SSProcess.SetNewObjValue "SSObj_GroupID", polygonID
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "1"
    SSProcess.SetNewObjValue "SSObj_FontWidth",width
    SSProcess.SetNewObjValue "SSObj_FontHeight", height
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function GetFormatNumber(ByVal number,ByVal numberDigit)
    If IsNumeric(numberDigit) = False Then numberDigit = 2
    If IsNumeric(number) = False Then number = 0
    number = FormatNumber(Round(number + 0.00000001,numberDigit),numberDigit, - 1,0,0)
    GetFormatNumber = (number)
End Function

