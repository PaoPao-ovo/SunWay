'房屋结构
Const FWJG_HX = "砼,砖,钢,混,石,木,毡,草,土,竹,秫秸,玻璃,其他"

'房屋层数
Const FWCS_HX = "2,3,4,5,6,7,8,9,10,11,12,13"


Const TDYT = "011:水田,012:水浇地,013:旱地,021:果园,022:茶园,023:其它园地,031:有林地,032:灌木林地,033:其它林地,041:天然牧草地,042:人工牧草地,043:其它草地,051:批发零售用地,052:住宿餐饮用地,053:商务金融用地,054:其它商服用地,0601:工业用地,0602:采矿用地,0604:仓储用地,0701:城镇住宅用地,0702:农村宅基地,0801:机关团体用地,0802:新闻出版用地,0803:科教用地,0804:医卫慈善用地,0805:文体娱乐用地,0806:公共设施用地,0807:公园与绿地,0808:风景名胜设施用地,091:军事设施用地,092:使领馆用地,093:监教场所用地,094:宗教用地,095:殡葬用地,101:铁路用地,102:公路用地,103:街巷用地,104:农村道路,105:机场用地,106:港口码头用地,107:管道运输用地,111:河流水面,112:湖泊水面,113:水库水面,114:坑塘水面,115:沿海滩涂,116:内陆滩涂,117:沟渠,118:水工建筑用地,119:冰川及永久积雪,121:空闲地,122:设施农用地,123:田坎,124:盐碱地,125:沼泽地,126:沙地,127:裸地"

QuanShuLeiXing = "A:集体土地所有权宗地,B:建设用地使用权宗地（地表）,S:建设用地使用权宗地（地上）,X:建设用地使用权宗地（地下）,C:宅基地使用权宗地,D:土地承包经营权宗地（耕地）,E:土地承包经营权宗地（林地）,F:土地承包经营权宗地（草地）,H:海域使用权宗海,G:无居民海岛使用权,W:使用权未确定或有争议的土地或海域海岛,Y:其它使用权土地、海域、海岛"                            '土地权属类型


Const QLLXS = "1:集体土地所有权,2:国家土地所有权,3:国有建设用地使用权,4:国有建设用地使用权/房屋（构筑物）所有权,5:宅基地使用权,6:宅基地使用权/房屋（构筑物）所有权,7:集体建设用地使用权,8:集体建设用地使用权/房屋（构筑物）所有权,9:土地承包经营权,10:土地承包经营权/森林、林木所有权,11:林地使用权,12:林地使用权/森林、林木使用权,13:草原使用权,14:水域滩涂养殖权,15:海域使用权,16:海域使用权/构（建）筑物所有权,17:无居民海岛使用权,18:无居民海岛使用权/构（建）筑物所有权,19:地役权,20:取水权,21:探矿权,22:采矿权,23:其它权利"

Const QLXZS = "100:国有土地,101:划拨,102:出让,103:作价出资（入股）,104:租赁,105:授权经营,200:集体土地,201:家庭承包,202:其它方式承包,203:批准拨用,204:入股,205:联营"

'线上点的数组(x,y,z,name)
Dim PointArr1(2,4)

'检查集组名
Dim strGroupName1
strGroupName1 = "绘线检查"

'检查集检查名
Dim strCheckName1
strCheckName1 = "检查线检查"

'检查日志
Dim strPromptMessage1
strPromptMessage1 = "请手动填写测站点号和检查点号"

'===================================================================================================================

'线上点的数组(x,y,z,name) —— 测站点、检查点
Dim PointArr2(2,4)
'检查集组名

Dim strGroupName2
strGroupName2 = "绘线检查"
'检查集检查名

Dim strCheckName2
strCheckName2 = "方向线检查"
'检查日志

Dim strPromptMessage2
strPromptMessage2 = "请手动填写测站点号和方向点号"

'=============================================================================================================================
'线上点的数组
Dim PointArr3(2,4)

'检查集组名
Dim strGroupName3
strGroupName3 = "绘线检查"

'检查集检查名
Dim strCheckName3
strCheckName3 = "控制点检查线线检查"

'检查日志
Dim strPromptMessage3
strPromptMessage3 = "请手动填写测站点号和检查点号"


#include"支点线_支点线.vbs"


Sub OnClick()
    
    
    SSParameter.GetParameterINT "AfterAddLine", "CurrentObjID", "0", ObjID
    If ObjID = 0 Then Exit Sub
    ObjCode = SSProcess.GetObjectAttr (objID, "SSObj_Code")
    If ObjCode = "" Then
        OBJID = SSProcess.GetGeoMaxID()
        ObjCode = SSProcess.GetObjectAttr (OBJID, "SSObj_Code")
        If ObjCode = "" Then  Exit Sub
    End If
    
    '=============================================================================================================================================================
    If objcode = 9130241 Then
        GetOnlinePoint1(objID)
        SearchNear1(objID)
    End If
    
    If objcode = 9130251 Then
        GetOnlinePoint2(objID)
        SearchNear2(objID)
    End If
    
    If objcode = 1130212 Then
        GetOnlinePoint3(objID)
        SearchNear3(objID)
        SetYZBC(objID)
        comparelong(objID)
    End If
    
    If objcode = 9414032  Then'用地红线标注
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "标注内容", "用地红线",0, "地下室范围线", ""
        res = SSProcess.ShowInputParameterDlg ("标注内容" )
        bzny = SSProcess.GetInputParameter ("标注内容" )
        
        SSProcess.SetObjectAttr CLng(objID), "[BiaoZNR]", bzny
        If bzny = "用地红线"  Then
            SSProcess.SetObjectAttr CLng(objID), "SSObj_Color", RGB(255,0,0)
        ElseIf bzny = "地下室范围线"  Then
            SSProcess.SetObjectAttr CLng(objID), "SSObj_Color", RGB(0,0,255)
        Else
            SSProcess.SetObjectAttr CLng(objID), "SSObj_Color", RGB(255,255,255)
        End If
    End If
    
    If objCode = 9470013 Then'绿化范围线属性表
        SSProcess.ClearInputParameter
        LDBH = SSProcess.ReadEpsDBIni("绿地编号", "编号" ,"")
        SCMJ = SSProcess.GetObjectAttr( ObjID, "SSObj_Area")
        SCMJ = FormatNumber(SCMJ,3, - 1,0,0)
        SSProcess.AddInputParameter "绿地图斑号",LDBH, 0, "", ""
        SSProcess.AddInputParameter "绿地类型", "地面绿化",  0, "地面绿化,地下室顶绿化,屋顶绿化,园路及园林铺装,景观水体", ""
        SSProcess.AddInputParameter "绿地细分", "园路及园林铺装",  0, "园路及园林铺装,景观水体,地面,地下室顶,屋顶", ""
        SSProcess.AddInputParameter "覆土厚度", "",0,"", "填写数值"
        SSProcess.AddInputParameter "是否集中绿地", "否", 0, "是,否", ""
        'SSProcess.AddInputParameter "规划审批绿地面积", "", 0, "", ""
        result = SSProcess.ShowInputParameterDlg ("录入属性")
        If result = 1 Then
            bh = SSProcess.GetInputParameter ("绿地图斑号")
            lx = SSProcess.GetInputParameter ("绿地类型")
            xl = SSProcess.GetInputParameter ("绿地细分")
            SFJZLD = SSProcess.GetInputParameter ("是否集中绿地")
            fthd = SSProcess.GetInputParameter ("覆土厚度")
            
            If ghspmj = "" Then ghspmj = 0
            SSProcess.SetObjectAttr ObjID, "[LvHTBH]", bh
            SSProcess.SetObjectAttr ObjID, "[LvHLX]", lx
            SSProcess.SetObjectAttr ObjID, "[LvHXL]", xl
            SSProcess.SetObjectAttr ObjID, "[TuBMJ]",SCMJ
            SSProcess.SetObjectAttr ObjID, "[FuTHD]",fthd
            SSProcess.SetObjectAttr ObjID, "[SFJZLD]",SFJZLD
            If  IsNumeric (bh) = True Then
                LDBH = CDbl(bh) + 1
                SSProcess.WriteEpsDBIni "绿地编号", "编号" ,LDBH
            Else
                SSProcess.WriteEpsDBIni "绿地编号", "编号" ,bh
            End If
        End If
    End If
    
    '自然幢绘制
    If objCode = 9210123 Then
        '自然幢信息
        strLCXX_ZRZ = "LCFZXX"                     '楼层信息-自然幢
        strCHZT_ZRZ = "CHZT"                         '测绘状态-自然幢
        strLJZLB = "LJZHLB"                         '逻辑幢列表
        strZRZH_ZRZ = "ZRZH"                          '幢号-自然幢
        strFWJG_ZRZ = "FWJG"                          '房屋结构-自然幢
        strFWJGM_ZRZ = "FWJGNAME"                  '房屋结构名-自然幢
        
        '属性对话框
        ZRZ_AttrDlg strFWZH, strFWJG, strZCS, strZTS, strCHZT, strLCFZ2, strsxh,strQH
        SSProcess.SetObjectAttr ObjID, "[QiuH]", strQH
        SSProcess.SetObjectAttr ObjID, "[ZRZH]", strFWZH
        SSProcess.SetObjectAttr ObjID, "[FWJG]", strFWJG
        SSProcess.SetObjectAttr ObjID, "[ZCS]", strZCS
        SSProcess.SetObjectAttr ObjID, "[ZTS]", strZTS
        SSProcess.SetObjectAttr ObjID, "[CHZT]", strCHZT
        SSProcess.SetObjectAttr ObjID, "[LCFZXX]", strLCFZ2
        SSProcess.SetObjectAttr ObjID, "[ZRZSXH]", strsxh
        
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    End If
    
    
    If objCode = 9210872 Then
        '楼层中轴线
        strIDs = SSProcess.SearchNearObjIDs2 (ObjID, 2, "9210123", 0 )
        If strIDs <> ""  And InStr(strIDs,",") = 0 Then
            zrzguid = SSProcess.GetObjectAttr (strIDs, "[ZRZGUID]" )
            strFields = "LCGUID"
            fieldsCount = 1
            sql = "select " & strFields & " from FC_楼层信息属性表 inner join GeoAreaTB on FC_楼层信息属性表.ID=GeoAreaTB.ID where (GeoAreaTB.mark mod 2) <> 0 and ZRZGUID=" & zrzguid & " order BY val(CH)"
            GetMdbValues sql,strFields,fieldsCount,lcAr,lcCount
            lcguid = lcAr(0,0)
            SSProcess.SetObjectAttr ObjID, "SSObj_DataMark", lcguid
        End If
        SSProcess.ExecuteSDLFunction "$SDL.SSProject.Display.RedrawExtend", 0
    End If
    
    '====================================================放验线
    If objcode = 9130221 Then '支点线
        zd(objID)
    End If
    '====================================================规划建筑轴线
    If objcode = 9310013 Then '规划建筑轴线
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "建筑物名称", "",0, "1#楼", ""
        res = SSProcess.ShowInputParameterDlg ("建筑物名称" )
        If res = 1 Then
            JianZWMC = SSProcess.GetInputParameter ("建筑物名称" )
            SSProcess.SetObjectAttr objID, "[JianZWMC]",  JianZWMC
        End If
    End If
    
    
    '=============================================================================================================================================
    
    If objCode = 3103013 Or objCode = 3103014 Or objCode = 3104003 Or objCode = 3105003 Or objCode = 3108003 Or objCode = 31030131 Then'310301301  Then
        obj_area = SSProcess.GetObjectAttr (objID, "SSObj_Area")
        obj_area = FormatNumber(obj_area,3)
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "建筑物结构", "混",0, FWJG_HX, ""
        SSProcess.AddInputParameter "楼层数目", "1",0, FWCS_HX, ""
        
        
        res = SSProcess.ShowInputParameterDlg ("建筑物结构" )
        If res = 0 Then
            '符号化刷新
            SSProcess.ObjectDeal objID, "AddToSelection", "", result
            SSProcess.ObjectDeal 0, "FreeSelectionObjectDisplayList", "", result
            Exit Sub
        End If
        
        FWJG = SSProcess.GetInputParameter ("建筑物结构" )
        FWCS = SSProcess.GetInputParameter ("楼层数目" )
        
        SSProcess.SetObjectAttr CLng(objID), "[CONSTRUCT]", FWJG
        SSProcess.SetObjectAttr CLng(objID), "[OGLAYER]", FWCS
        
        
        pointcount = SSProcess.GetObjectAttr (objID, "SSObj_PointCount")
        
        SSProcess.GetObjectPoint objID, 0, x0,  y0,  z0,  ptype0,  name0
        For i = 1 To pointcount - 1
            SSProcess.GetObjectPoint objID, i, x1,  y1,  z1,  ptype1,  name1
            SSProcess.SetObjectPoint objID, i, x1,  y1,  z0,  ptype1,  name1, 1
        Next
        
        
    End If
    '符号化刷新
    SSProcess.ObjectDeal objID, "AddToSelection", "", result
    SSProcess.ObjectDeal 0, "FreeSelectionObjectDisplayList", "", result
    SSProcess.RefreshView
End Sub

'获取MDB信息
Function GetMdbValues(ByVal sql,ByVal strFields,ByVal fieldsCount,ByRef rs,ByRef rscount)
    
    mdbName = SSProcess.GetProjectFileName()
    SSProcess.OpenAccessMdb mdbName
    SSProcess.OpenAccessRecordset mdbName, sql
    rscount = SSProcess.GetAccessRecordCount (mdbName, sql)
    ReDim rs(rscount,fieldsCount)
    'addloginfo "sql=" & sql & ",fieldsCount=" & fieldsCount
    If rscount > 0 Then
        SSProcess.AccessMoveFirst mdbName, sql
        n = 0
        While SSProcess.AccessIsEOF (mdbName, sql) = False
            SSProcess.GetAccessRecord mdbName, sql, fields, values
            If IsNull(values) Then values = ""
            values = Replace(values,"|","，")
            strs = Split(values,",")
            If UBound(strs) <> - 1 Then
                For i = 0 To fieldsCount - 1
                    rs(n,i) = strs(i)
                Next
            End If
            SSProcess.AccessMoveNext mdbName, sql
            n = n + 1
        WEnd
    End If
    SSProcess.CloseAccessRecordset mdbName, sql
    SSProcess.CloseAccessMdb mdbName
    
End Function

'宗地构面
Function ZDGM()
    SSProcess.PushUndoMark
    
    '删除原来的宗地面
    'SSProcess.ClearSelection
    'SSProcess.ClearSelectCondition
    'SSProcess.SetSelectCondition  "SSObj_Type","=","AREA"
    'SSProcess.SetSelectCondition  "SSObj_Code","=","6803153"
    'SSProcess.SetSelectCondition  "SSObj_LayerName","=","宗地"
    'SSProcess.SelectFilter
    'SSProcess.DeleteSelectionObj
    
    SSProcess.ClearFunctionParameter
    '悬挂点处理限距
    SSProcess.AddFunctionParameter "limitdist=0.0001"
    '拓扑弧段编码
    SSProcess.AddFunctionParameter "SrcArcCodes=9130242,6801332,6803232,6803152"
    '删除源弧段
    SSProcess.AddFunctionParameter "DelSrcArc=0"
    '删除上次生成的重叠弧段
    SSProcess.AddFunctionParameter "DelNewArc=0"
    '删除上次生成的原拓扑面
    SSProcess.AddFunctionParameter "DelOldTopArea=0"
    '数据处理后是否存盘
    SSProcess.AddFunctionParameter "SaveDB=1"
    '是否生成拓扑面
    SSProcess.AddFunctionParameter "CreateTopArea=1"
    '拓扑面编码设置  属性点编码1,面编码1,图层名称1/属性点编码2,面编码2,图层名称2
    SSProcess.AddFunctionParameter "NewObject=913022301,9130223"
    '判断属性点重复的关键字
    SSProcess.AddFunctionParameter "LabelKeyFields="
    '生成拓扑弧段选择：
    '0 不生成弧段
    '1 生成统一编码弧段，编码由UniqueArcCode指定
    '2 生成弧段, 当有多种线状地物重叠时，按 ReserveArcOrder设置的编码顺序优先从前选取
    '3 自动生成与其他弧段重叠的新弧段, 按CreateOverlayArc设置
    SSProcess.AddFunctionParameter "CreateTopArc=0"
    '弧生成方法
    jx = ""
    jx = jx & "250200/行政区划海岸线/733001/宗地拓扑线"
    jx = jx & ",250201/行政区划高潮线/733001/宗地拓扑线"
    SSProcess.AddFunctionParameter "CreateOverlayArc=" & jx
    
    SSProcess.TopProcess "宗地构面"
End Function

'自然幢属性对话框
Function ZRZ_AttrDlg(ByRef strFWZH,ByRef strFWJG,ByRef strZCS,ByRef strZTS,ByRef strCHZT,ByRef strLCFZ2,ByRef strsxh,ByRef strQH)
    SSProcess.ClearInputParameter
    SSProcess.AddInputParameter "丘号", "", 0, "", ""
    SSProcess.AddInputParameter "自然幢号", "1幢", 0, "1幢,2幢,3幢", ""
    SSProcess.AddInputParameter "幢顺序号", "0001", 0, "0001,0002,0003", ""
    SSProcess.AddInputParameter "房屋结构", "5:砖木结构", 0, "1:钢结构,2:钢和钢筋混凝土结构,3:钢筋混凝土结构,4:混合结构,5:砖木结构,6:其它结构", "房屋结构取值 1:钢结构,2:钢和钢筋混凝土结构,3:钢筋混凝土结构,4:混合结构,5:砖木结构,6:其它结构"
    SSProcess.AddInputParameter "总层数", "1", 0, "2,3,4,5,6,7,8,9,10,11,12,13,14,15,16", ""
    SSProcess.AddInputParameter "总套数", "1", 0, "2,3,4,5,6,7,8,9,10,11,12,13,14,15,16", ""
    SSProcess.AddInputParameter "测绘状态", "2:实测", 0, "1:预测,2:实测", ""
    SSProcess.AddInputParameter "楼层分组信息", "1", 0, "", strLCFZ2
    'SSProcess.AddInputParameter "逻辑幢号列表", "1", 0, "1,1、2,1、2、3", ""
    
    SSProcess.ShowInputParameterDlg title
    strQH = SSProcess.GetInputParameter ("丘号")
    strFWZH = SSProcess.GetInputParameter ("自然幢号")
    strFWJG = SSProcess.GetInputParameter ("房屋结构")
    strZCS = SSProcess.GetInputParameter ("总层数")
    strZTS = SSProcess.GetInputParameter ("总套数")
    strCHZT = SSProcess.GetInputParameter ("测绘状态")
    strLCFZ2 = SSProcess.GetInputParameter ("楼层分组信息")
    strsxh = SSProcess.GetInputParameter ("幢顺序号")
    
    '房屋结构
    If Replace(strFWJG,":","") <> strFWJG Then
        arFWJG = Split(strFWJG,":")
        strFWJG = arFWJG(0)
        strFWJGMC = arFWJG(1)
    End If
    '测绘状态
    If Replace(strCHZT,":","") <> strCHZT Then
        arCHZT = Split(strCHZT,":")
        strCHZT = arCHZT(0)
        'strFWJGMC =arCHZT(1)
    End If
End Function

'获取竣工信息
'获取许可证及建筑物名称
Function GetDTByJSGCGHXKZBH (DT)
    
    Dim Fvalues(1000)
    DT = ""
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_建设工程建筑单体信息属性表.JianZWMC,GuiHXKZBH FROM (JG_建设工程建筑单体信息属性表 inner join JG_用地红线信息属性表 on JG_建设工程建筑单体信息属性表.YDHXGUID = JG_用地红线信息属性表.YDHXGUID)  inner join GeoAreaTB on GeoAreaTB.ID = JG_用地红线信息属性表.ID  WHERE ((GeoAreaTB.mark mod 2) <> 0)  ORDER BY JG_建设工程建筑单体信息属性表.JianZWMC;"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            
            If values <> "" And  values <> "*" And values <> "NULL" Then
                SSFunc.ScanString values, ",", Fvalues, FvaluesCount
                GHXKZH = Fvalues(1)
                DTMC = Fvalues(0)
                'GHSPJDMJ=Fvalues(2)
                If DT = "" Then
                    DT = GHXKZH & "|" & DTMC
                Else
                    DT = DT & "," & GHXKZH & "|" & DTMC
                End If
            End If
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function


'*************获取竣工单体建筑信息**************************
Function GetjgclGDxx(ghxkzbh,jzwmc,GHSPZFL,GHSPZGD,GHSPDXZGD,JGDSCS,JGDXCS,JGCLJGLX,GHSPJDMJ,YDHXGUID,JSGHXKZGUID,JZWMCGUID,GuiHYDXKZBH)
    Dim Fvalues(1000)
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_建设工程建筑单体信息属性表.GuiHSPZFL,GuiHSPDSCS,GuiHSPDXCS,JunGCLDSCS,JunGCLDXCS,JunGCLJGLX,GuiHSPJDMJ,YDHXGUID,JSGHXKZGUID,JZWMCGUID,GuiHYDXKZBH FROM JG_建设工程建筑单体信息属性表 WHERE ([JG_建设工程建筑单体信息属性表].[ID] > 0 And ([JG_建设工程建筑单体信息属性表].[JianZWMC] = '" & jzwmc & "') And ([JG_建设工程建筑单体信息属性表].[GuiHXKZBH] = '" & ghxkzbh & "'));"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            SSFunc.Scanstring values,",",Fvalues,Fvaluescount
            GHSPZFL = Fvalues(0)
            GHSPZGD = Fvalues(1)
            GHSPDXZGD = Fvalues(2)
            JGDSCS = Fvalues(3)
            JGDXCS = Fvalues(4)
            JGCLJGLX = Fvalues(5)
            GHSPJDMJ = Fvalues(6)
            YDHXGUID = Fvalues(7)
            JSGHXKZGUID = Fvalues(8)
            JZWMCGUID = Fvalues(9)
            GuiHYDXKZBH = Fvalues(10)
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function




'*************获取规划许可证信息**************************
Function Getjhxkzxx(ghxkzbh,jzwmc,XiangMMC)
    Dim Fvalues(6)
    'MSGBOX ghxkzbh
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_建设工程规划许可证信息属性表.XiangMMC FROM JG_建设工程规划许可证信息属性表 WHERE ([JG_建设工程规划许可证信息属性表].[ID] > 0  And ([JG_建设工程规划许可证信息属性表].[GuiHXKZBH] = '" & ghxkzbh & "'));"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            SSFunc.Scanstring values,",",Fvalues,Fvaluescount
            XiangMMC = Fvalues(0)
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function



Function GetDTJZWJJXX (DT)
    
    Dim Fvalues(1000)
    DT = ""
    projectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb projectName
    sql = "SELECT JG_建设工程建筑单体信息属性表.JianZWMC,GuiHXKZBH,JZWMCGUID,JSGHXKZGUID FROM (JG_建设工程建筑单体信息属性表 inner join JG_用地红线信息属性表 on JG_建设工程建筑单体信息属性表.YDHXGUID = JG_用地红线信息属性表.YDHXGUID)  inner join GeoAreaTB on GeoAreaTB.ID = JG_用地红线信息属性表.ID  WHERE ((GeoAreaTB.mark mod 2) <> 0)  ORDER BY JG_建设工程建筑单体信息属性表.JianZWMC;"
    SSProcess.OpenAccessRecordset projectName, sql
    rscount = SSProcess.GetAccessRecordCount (projectName, sql )
    If rscount > 0 Then
        SSProcess.AccessMoveFirst projectName, sql
        While (SSProcess.AccessIsEOF (projectName, sql ) = False)
            SSProcess.GetAccessRecord projectName, sql, fields, values
            
            If values <> "" And  values <> "*" And values <> "NULL" Then
                SSFunc.ScanString values, ",", Fvalues, FvaluesCount
                GHXKZH = Fvalues(1)
                DTMC = Fvalues(0)
                JZWGUID = Fvalues(2)
                GHXKZGUID = Fvalues(3)
                'GHSPJDMJ=Fvalues(2)
                If DT = "" Then
                    DT = GHXKZH & "|" & DTMC & "|" & JZWGUID & "|" & GHXKZGUID
                Else
                    DT = DT & "," & GHXKZH & "|" & DTMC & "|" & JZWGUID & "|" & GHXKZGUID
                End If
            End If
            SSProcess.AccessMoveNext projectName, sql
        WEnd
    End If
    
    SSProcess.CloseAccessRecordset projectName, sql
    SSProcess.CloseAccessMdb projectName
End Function

'==============================================================================================================================================================================================
'传值函数
Function SearchNear1(id)
    x1 = PointArr1(0,0)
    y1 = PointArr1(0,1)
    x2 = PointArr1(1,0)
    y2 = PointArr1(1,1)
    SetLinepoiname1 x1,y1,x2,y2,id
    SetProp1 x1,y1,x2,y2,id
End Function' SearchNear

'获取线上的空间点信息
Function GetOnlinePoint1(id)
    Dim x, y, z, pointtype, name
    pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
    'MsgBox pointcount
    pointcount = transform(pointcount)
    For j = 0 To pointcount - 1
        SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name
        x = transform(x)
        y = transform(y)
        z = transform(z)
        PointArr1(j,0) = x
        PointArr1(j,1) = y
        PointArr1(j,2) = z
        PointArr1(j,3) = name
    Next
    'MsgBox PointArr(1,0)
End Function' GetOnlinePoint

'设置线的方向值和水平距离(方向值暂留)
Function SetProp1(x1,y1,x2,y2,id)
    longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
    longtitude = transform(longtitude)
    longtitude = FormatNumber(longtitude,3)
    If x1 < x2 And y1 < y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 270 + SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 > x2 And y1 < y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 90 - SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 < x2 And y1 > y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 90 + SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 > x2 And y1 > y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 180 + SSProcess.RadianToDms(Atn(Abs(y / x)))
    End If
    angarr = Split(angles,".", - 1,1)
    If UBound(angarr) > 0 Then
        str = angarr(1)
        dd = ""
        ss = ""
        If Len(str) > 4 Then
            dd = Mid(str,1,2)
            ss = Mid(str,3,2)
        End If
        If Len(str) = 3 Then
            dd = Mid(str,1,2)
            ss = Mid(str,3,1) & "0"
        End If
        If Len(str) = 2 Then
            dd = Mid(str,1,2)
            ss = "00"
        End If
        If Len(str) = 1 Then
            dd = Mid(str,1,1) & "0"
            ss = "00"
        End If
        If Len(str) = 0 Then
            dd = "00"
            ss = "00"
        End If
    ElseIf UBound(angarr) = 0 Then
        dd = "00"
        ss = "00"
    End If
    SSProcess.SetObjectAttr id,"[ShuiPJL]",longtitude
    SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "°" & dd & "′" & ss & "″"
End Function' SetProp

'搜索理论控制点名称
Function SetLinepoiname1(x1,y1,x2,y2,id)
    SSProcess.RemoveCheckRecord strGroupName1, strCheckName1
    idstring = SSProcess.SearchNearObjIDs(x1,y1,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo1 x1,y1,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        'MsgBox id
        SSProcess.SetObjectAttr id,"[CeZDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            'MsgBox id
            ExportInfo1 x1,y1,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
        End If
    End If
    
    idstring = SSProcess.SearchNearObjIDs(x2,y2,0.001,0,"9130311,9130312,9130217",0)
    idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo1 x2,y2,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        SSProcess.SetObjectAttr id,"[JianCDH]",pointname
        code = SSProcess.GetObjectAttr(idarr(0),"SSObj_Code")
        If code = "9130217" Then
            DiffXY id,"9130216"
        ElseIf code = "9130311" Then
            DiffXY id,"9130211"
        ElseIf code = "9130312" Then
            DiffXY id,"9130212"
        End If
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            ExportInfo1 x2,y2,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[ZhiZDH]",Firstname
        End If
    End If
End Function' SetLinepoiname

'设置X,Y差值
Function DiffXY(id,Code)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", Code
    SSProcess.SetSelectCondition "SSObj_PointName", "==",PointArr1(1,3)
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    'MsgBox PointArr(1,3)
    If SelCount > 0 Then
        X = SSProcess.GetSelGeoValue(0, "SSObj_X")
        X = transform(X)
        Y = SSProcess.GetSelGeoValue(0, "SSObj_Y")
        Y = transform(Y)
        diffx = Abs(X - PointArr1(1,0))
        diffy = Abs(Y - PointArr1(1,1))
        diffx = FormatNumber(diffx,3)
        diffy = FormatNumber(diffy,3)
        SSProcess.SetObjectAttr id,"[XZuoBCZ]",diffx
        SSProcess.SetObjectAttr id,"[YZuoBCZ]",diffy
    Else
        'MsgBox "不存在同名点" 
        Exit Function
    End If
End Function' DiffXY

'输出检查集函数
Function ExportInfo1(x,y,z,id)
    SSProcess.AddCheckRecord strGroupName1, strCheckName1, "自定义脚本检查类->" & strCheckName1, strPromptMessage1, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function' ExportInfo

'数据类型转换
Function transform(content)
    If content <> "" Then
        content = CDbl(content)
    Else
        MsgBox "数据有误"
        Exit Function
    End If
    transform = content
End Function

'=============================================================================================================================
Function SearchNear2(id)
    x1 = PointArr2(0,0)
    y1 = PointArr2(0,1)
    x2 = PointArr2(1,0)
    y2 = PointArr2(1,1)
    SetLinepoiname2 x1,y1,x2,y2,id
    SetProp2 x1,y1,x2,y2,id
End Function' SearchNear

'获取线上的空间点信息
Function GetOnlinePoint2(id)
    Dim x, y, z, pointtype, name
    pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
    'MsgBox pointcount
    pointcount = transform(pointcount)
    For j = 0 To pointcount - 1
        SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name
        x = transform(x)
        y = transform(y)
        z = transform(z)
        PointArr2(j,0) = x
        PointArr2(j,1) = y
        PointArr2(j,2) = z
        PointArr2(j,3) = name
    Next
    'MsgBox PointArr(1,0)
End Function' GetOnlinePoint

'设置线的方向值和水平距离(方向值暂留)
Function SetProp2(x1,y1,x2,y2,id)
    longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
    longtitude = transform(longtitude)
    longtitude = FormatNumber(longtitude,3)
    If x1 < x2 And y1 < y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 270 + SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 > x2 And y1 < y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 90 - SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 < x2 And y1 > y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 90 + SSProcess.RadianToDms(Atn(Abs(x / y)))
    End If
    If x1 > x2 And y1 > y2 Then
        x = x2 - x1
        y = y2 - y1
        angles = 180 + SSProcess.RadianToDms(Atn(Abs(y / x)))
    End If
    angarr = Split(angles,".", - 1,1)
    If UBound(angarr) > 0 Then
        str = angarr(1)
        dd = ""
        ss = ""
        If Len(str) > 4 Then
            dd = Mid(str,1,2)
            ss = Mid(str,3,2)
        End If
        If Len(str) = 3 Then
            dd = Mid(str,1,2)
            ss = Mid(str,3,1) & "0"
        End If
        If Len(str) = 2 Then
            dd = Mid(str,1,2)
            ss = "00"
        End If
        If Len(str) = 1 Then
            dd = Mid(str,1,1) & "0"
            ss = "00"
        End If
        If Len(str) = 0 Then
            dd = "00"
            ss = "00"
        End If
    ElseIf UBound(angarr) = 0 Then
        dd = "00"
        ss = "00"
    End If
    SSProcess.SetObjectAttr id,"[ShuiPJL]",longtitude
    SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "°" & dd & "′" & ss & "″"
End Function' SetProp

'搜索理论控制点名称
Function SetLinepoiname2(x1,y1,x2,y2,id)
    SSProcess.RemoveCheckRecord strGroupName2, strCheckName2
    idstring = SSProcess.SearchNearObjIDs(x1,y1,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo2 x1,y1,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        'MsgBox id
        SSProcess.SetObjectAttr id,"[CeZDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            'MsgBox id
            ExportInfo2 x1,y1,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
        End If
    End If
    
    idstring = SSProcess.SearchNearObjIDs(x2,y2,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo2 x2,y2,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        SSProcess.SetObjectAttr id,"[FangXDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            ExportInfo2 x2,y2,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[FangXDH]",Firstname
        End If
    End If
End Function' SetLinepoiname

'输出检查集函数
Function ExportInfo2(x,y,z,id)
    SSProcess.AddCheckRecord strGroupName2, strCheckName2, "自定义脚本检查类->" & strCheckName2, strPromptMessage2, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function' ExportInfo


'===========================================================================================================================================

'传值函数
Function SearchNear3(id)
    x1 = PointArr3(0,0)
    y1 = PointArr3(0,1)
    x2 = PointArr3(1,0)
    y2 = PointArr3(1,1)
    SetLinepoiname3 x1,y1,x2,y2,id
    SetProp3 x1,y1,x2,y2,id
End Function' SearchNear

'获取线上的空间点信息
Function GetOnlinePoint3(id)
    Dim x, y, z, pointtype, name
    pointcount = SSProcess.GetObjectAttr(id,"SSObj_PointCount")
    'MsgBox pointcount
    pointcount = transform(pointcount)
    For j = 0 To pointcount - 1
        SSProcess.GetObjectPoint id,j,x,y,z,pointtype,name
        x = transform(x)
        y = transform(y)
        z = transform(z)
        PointArr3(j,0) = x
        PointArr3(j,1) = y
        PointArr3(j,2) = z
        PointArr3(j,3) = name
    Next
    'MsgBox PointArr(1,0)
End Function' GetOnlinePoint

'设置线的方向值和水平距离
Function SetProp3(x1,y1,x2,y2,id)
    longtitude = SSProcess.GetObjectAttr(id,"SSObj_Length")
    longtitude = transform(longtitude)
    longtitude = FormatNumber(longtitude,3)
    SSProcess.SetObjectAttr id,"[JCBC]",longtitude
    'SSProcess.SetObjectAttr id,"[FangXZ]",angarr(0) & "°" & dd & "′" & ss & "″"
End Function' SetProp

'设置已知边长
Function SetYZBC(id)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9130211"
    SSProcess.SetSelectCondition "SSObj_PointName", "==",PointArr3(1,3)
    SSProcess.SelectFilter
    SelCount = SSProcess.GetSelGeoCount()
    'msgbox PointArr3(1,3)
    If SelCount > 0 Then
        X = SSProcess.GetSelGeoValue(0, "SSObj_X")
        X = transform(X)
        Y = SSProcess.GetSelGeoValue(0, "SSObj_Y")
        Y = transform(Y)
        yzbc = Sqr((PointArr3(0,0) - X) ^ 2 + (PointArr3(0,1) - Y) ^ 2)
        yzbc = FormatNumber(yzbc,3)
        SSProcess.SetObjectAttr id,"[YZBC]",yzbc
    End If
End Function' SetYZBC

'计算边长较差
Function comparelong(id)
    yzbc = SSProcess.GetObjectAttr(id,"[YZBC]")
    jcbc = SSProcess.GetObjectAttr(id,"[JCBC]")
    yzbc = transform(yzbc)
    jcbc = transform(jcbc)
    bcjc = Abs(yzbc - jcbc)
    SSProcess.SetObjectAttr id,"[BCJC]",bcjc
End Function' comparelong

'设置测站检查点号名称
Function SetLinepoiname3(x1,y1,x2,y2,id)
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
    idstring = SSProcess.SearchNearObjIDs(x1,y1,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo3 x1,y1,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        'MsgBox id
        SSProcess.SetObjectAttr id,"[CeZDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            'MsgBox id
            ExportInfo3 x1,y1,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[CeZDH]",Firstname
        End If
    End If
    
    idstring = SSProcess.SearchNearObjIDs(x2,y2,0.001,0,"",0)
    idarr = Split(idstring,",", - 1,1) '与线上点相近的点的ids
    IdCount = UBound(idarr) + 1
    'MsgBox IdCount
    If IdCount = 0 Then ExportInfo3 x2,y2,0,id
    If IdCount = 1 Then
        pointname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        SSProcess.SetObjectAttr id,"[JianCDH]",pointname
    ElseIf IdCount = 2 Then
        Firstname = SSProcess.GetObjectAttr(idarr(0),"SSObj_PointName")
        Secondname = SSProcess.GetObjectAttr(idarr(1),"SSObj_PointName")
        If Firstname <> Secondname Then
            ExportInfo3 x2,y2,0,id
            'Exit Function
        End If
        If Firstname = Secondname Then
            SSProcess.SetObjectAttr id,"[JianCDH]",Firstname
        End If
    End If
End Function' SetLinepoiname

'输出检查集函数
Function ExportInfo3(x,y,z,id)
    SSProcess.AddCheckRecord strGroupName3, strCheckName3, "自定义脚本检查类->" & strCheckName3, strPromptMessage3, x, y, z, 1, id, ""
    SSProcess.ShowCheckOutput
End Function' ExportInfo

