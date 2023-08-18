Rem autor
 < Administrator > 
Rem email
XXXX@xxx.com
Rem 脚本文件名
F
 \ 1208 \ 苍南勘测 \ DeskTop \ 放验线 \ Script \ 放线 \ 放样图输出.vbs
Rem 对应方案文件名
F
 \ 1208 \ 苍南勘测 \ DeskTop \ 放验线 \ 功能模板 \ 放样图.Map
Rem 方案名称
放样图输出
Rem 本脚本文件应放置于 EPS安装目录 \ desktop \ XX台面 \ Script \ 
Rem framework
gq
Rem framework
471b1e20fe69040339fca38c3d3a189b




Rem special
[放样图] 出图前（初始化调用）由此进入
Function VBS_preMap0(MSGID,mapName,selectID)
    
    Rem 本函数关键参数：SSParameter.SetParameterINT "printMap", "return", 1
    Rem return = 1 停止输出成果图
    Rem return = 0 继续输出成果图（无需设置、默认值为0）
    
    If MSGID = 0 Then '// 新工程出图 
        '// 添加您的代码.... 
        '// 设置出图工程名称、必须调用.... ,批量出图的路径每次会调用脚本传回的路径，工程不能同名，通常可以用范围线地物的扩展属性拼接
        FileFolder = SSProcess.GetSysPathName (5)
        'CreateFolders FileFolder,"多测合一数据成果"
        SaveFile = FileFolder & "\3成果\放样图.edb"
        SSParameter.SetParameterSTR "printMap","NewedbName",SaveFile
        
    ElseIf MSGID = 1 Then '// 本工程出图 
        '// 添加您的代码.... 
        
    ElseIf MSGID = 2 Then '// 新工程自定义目录出图(自主选择保存路径) 
        '// 添加您的代码.... 
        
    End If
    
End Function



Rem special
[放样图] 出图完成由此进入
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
    
    GetTKSX
    DaHui
End Function



Dim g_MapList,g_MapPrePtrfun,g_MapPostPtrfun
Rem 主函数无需修改
Sub OnClick()
    
    Rem 初始化
    g_MapList = Array("放样图")
    g_MapPrePtrfun = Array("VBS_preMap0")
    g_MapPostPtrfun = Array("VBS_postMap0")
    
    Rem 系统传来的消息,用户选择的范围线ID,成果图名称
    Dim str_msg,str_selectObjid,str_mapName
    
    Rem 获取系统参数 -  - 用户选择范围线ID
    SSParameter.GetParameterINT "printMap", "SelectID", - 1, str_selectObjid
    
    Rem 获取系统参数 -  - 系统消息 （0：新工程固定目录出图初始化消息  1：本工程出图初始化消息  2
    新工程自定义目录出图初始化消息  3：出图已完成交付于脚本处理细节）
    SSParameter.GetParameterINT "printMap", "printMSG", - 1, str_msg
    
    Rem 获取系统参数 -  - 专题名称
    SSParameter.GetParameterSTR "printMap", "SpecialMapName", "", str_mapName
    
    DistributeMSG str_msg,str_mapName,str_selectObjid
End Sub




Rem 此虑数函数无需修改
Function DistributeMSG(MSGid,str_MapName,selectID)
    Dim pFun
    
    For i = 0 To UBound(g_MapList)
        If UCase(g_MapList(i)) = UCase(str_MapName) Then
            If MSGid = 3 Then
                
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

Function GetTKSX
    If 0 Then '隐藏
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_CODE", "==", "9130224"
        SSProcess.SelectFilter
        geocount = SSProcess.GetSelGeoCount
        If  geocount > 0 Then
            hxid = SSProcess.GetSelGeoValue(0,"SSObj_ID")
            XMMC = SSProcess.GetObjectAttr (hxid,"[XiangMMC]")
        End If
    End If
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_CODE", "==", "9130224"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount
    If  geocount > 0 Then
        id = SSProcess.GetSelGeoValue(i, "SSObj_id")
        SSProcess.SetObjectAttr id, "[图幅名称]","放样图"
    End If
End Function

Function DaHui
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_code", "<>", "9130224,9130411,9310013,9130611"
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    For i = 0 To geocount - 1
        geoID = SSProcess.GetSelGeoValue(i, "SSObj_ID")
        SSProcess.SetObjectAttr geoID, "SSObj_Color", RGB(0,0,0)
    Next
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "<>", "地类图斑,勘测村界,境界,征地项目面,征地界址点,勘测图廓层,征地注记,征地村注记,村属性点,乡镇属性点"
    SSProcess.SelectFilter
    notecount = SSProcess.GetSelNoteCount()
    For i1 = 0 To notecount - 1
        id = SSProcess.GetSelNoteValue(i1 ,"SSObj_ID" )
        SSProcess.SetObjectAttr id, "SSObj_Color", RGB(0,0,0)
    Next
    
    
End Function
