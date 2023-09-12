

DKMBM = "9130223"                                        '宗地面编码
DKJXBM = "9130242"                            '地块界线编码    6805122
DKJZDBM = "9130231"

Dim docObj
'ado 全局变量
Dim adoConnection
Dim JZDHHQFS
Dim YeWLX
#include ".\function\Encryption.vbs"
Dim  Registrationkeyidlist,HardID,usbkeyidlist,usbkeyid

Sub OnClick()
    RegistrationMode = SSProcess.ReadEpsGlobalIni("SoftRegister", "Mode" , "")
    
    'DKBJZDHHQFS =SSProcess.ReadEpsIni ("不动产确权", "地块表界址点号方式" , "顺序号" )                
    'SSProcess.ClearInputParameter
    
    'SSProcess.AddInputParameter "业务类型", "房开项目",0, "房开项目,民房项目,单一产权", ""
    
    JZDHHQFS = "1"
    
    'res = SSProcess.ShowInputParameterDlg ("业务类型" )
    'If res = 0  Then
    '    Exit Sub
    'End If
    'YeWLX = SSProcess.GetInputParameter ("业务类型" )
    
    mdbName = SSProcess.GetProjectFileName
    RPTPahtName = Left(mdbName,Len(mdbName) - 4)
    IsFolderExists(RPTPahtName)                            '判断文件夹是否存在，没有则创建
    'RPTPahtName =left(mdbName,len(mdbName)-4) & "\不动产权籍调查表\"
    'IsFolderExists(RPTPahtName)                            '判断文件夹是否存在，没有则创建
    
    
    
    Dim strKeyFiledName                                                '关键字段名，宗地代码字段
    strKeyFiledName = "ZDDM"
    
    Dim strTableZD
    
    strTableZD = "ZD_宗地基本信息属性表"            '宗地面属性表
    
    
    Dim strTableJZQZ
    strTableJZQZ = "ZD_界址签章信息表"                        '界址签章信息表
    Dim strTableGYMJFT
    strTableGYMJFT = "ZD_共有共用宗地面积分摊信息表"            '共有共用宗地面积分摊信息表
    
    Dim strTableHu
    strTableHu = "FC_户信息属性表"                            '房产 户信息表
    
    
    Dim PDBfileName
    Dim arRecordRPT(),nRecordRPTCount                                                '表单记录数组   3列
    Dim arRecordZD(),nRecordZDCount                                                    '宗地基本信息记录数组
    Dim arRecordQLR(),nRecordQLRCount                                                '权利人记录数组
    Dim arRecordJZBSB(),nRecordJZBSBCount                                            '界址标示表记录数组
    
    Dim arRecordJZQZB(),nRecordJZQZBCount                                            '界址签章表记录数组
    Dim arRecordGYMJFT(),nRecordGYMJFTCount                                            '共有共用宗地面积分摊信息表记录数组
    
    Dim arRecordZRZ(),nRecordZRZCount                                                '自然幢表记录数组
    
    Dim arRecordHU(), nRecordHUCount  '户信息
    Dim huxx(100), qlr(100), zrz(2)
    
    Dim arRecordCBDKMJHZ(),RecordCBDKMJHZCount                                        '承包地块面积汇总记录数组
    Dim arRecordCBHT(),RecordCBHTCount                                                '承包合同记录数组
    Dim arRecordFBF(),RecordFBFCount                                                '发包方记录数组
    Dim arRecordJTCY(),RecordJTCYCount                                                '家庭成员记录数组
    
    Dim arRecordCBFCheck(),RecordCBFCheckCount                                        '错误检查记录数组
    Dim arRecordDKCheck(),RecordDKCheckCount                                        '错误检查记录数组
    Dim strRecordSurplus,RecordSurplusCount,arRecordCheckTemp()                        '多余记录数组
    
    Dim arRecordShortList(),RecordShortListCount                                    '候选列表记录数组
    
    Dim strSelectedOutRPT                                                            '待输出信息
    Dim strSelectedGUIDOutRPT                                                        '待输出GUID信息
    
    'Dim arZD_SC(),nZDOutCount                                                        '宗地输出记录数组
    
    Dim arOutRecSelected(),nOutSelectedCount                                        '输出选项记录数组
    Dim arOutRecGUIDSelected(),nOutGUIDSelectedCount                                '输出选项记录GUID数组
    Dim arRecordZRZYT(),nRecordZRZYTCount
    '标记定义
    Dim Mark_ShareTable_Del                                                '删除共有/共用表的标记
    Dim Mark_ZD_BDCDYH                                                                '宗地不动产单元号
    
    Dim Mark_test
    Mark_test = 0
    SSProcess.OpenAccessMdb  mdbName
    
    '2.------------------选取待输出项---------------------
    '获取输出项记录
    sqltexts = "SELECT [" & strTableZD & "].[" & strKeyFiledName & "],[" & strTableZD & "].[ZDGUID] FROM " & strTableZD & " INNER JOIN GeoAreaTB ON " & strTableZD & ".ID = GeoAreaTB.ID "
    condition = " WHERE ([GeoAreaTB].[Mark] Mod 2)<>0"
    strOrderBy = " ORDER BY " & strTableZD & "." & strKeyFiledName
    sql = sqltexts & " " & condition & " " & strOrderBy
    MsgBox  sql
    GetSQLRecordAll mdbName,sql,arRecordShortList,RecordShortListCount
    
    '指定输出承包方编码
    'ResVal_Dlg =SSFunc.SelectListAttr("选择列表", "待选数据列表", "选中数据列表", arRecordShortList, RecordShortListCount)
    
    
    'if ResVal_Dlg = 1 then
    If RecordShortListCount <> 1 Then
        '关闭数据库
        SSProcess.CloseAccessMdb mdbName
        MsgBox "工程内无宗地或多个宗地，退出输出！"
        Exit Sub
    Else
        ReDim arOutRecSelected(RecordShortListCount)                '输出特征记录数组
        nOutSelectedCount = RecordShortListCount                        '输出记录数量
        ReDim arOutRecGUIDSelected(RecordShortListCount)            '输出特征记录GUID数组
        nOutSelectedCount = RecordShortListCount                        '输出记录GUID数量
        For i = 0 To RecordShortListCount - 1
            arOutTempCur = Split(arRecordShortList(i),",")    ':nCount_Temp =bound(arCBFBMCur)
            
            arOutRecSelected(i) = Trim(arOutTempCur(0))
            arOutRecGUIDSelected(i) = Trim(arOutTempCur(1))    'GUID
            
            If strSelectedOutRPT = "" Then
                strSelectedOutRPT = "'" & Trim(arOutTempCur(0)) & "'"
            Else
                strSelectedOutRPT = strSelectedOutRPT & ",'" & Trim(arOutTempCur(0)) & "'"
            End If
            If strSelectedGUIDOutRPT = "" Then
                strSelectedGUIDOutRPT = "{guid " & Trim(arOutTempCur(1)) & "}"
            Else
                strSelectedGUIDOutRPT = strSelectedGUIDOutRPT & ",{guid " & Trim(arOutTempCur(1)) & "}"
            End If
        Next
    End If
    'else
    '关闭数据库
    '    SSProcess.CloseAccessMdb mdbName
    '    exit sub
    'end if
    
    '3.---------------获取输出表的内容（GetOutContent）-------------------------------------
    '-----GetOutContent.3.1---获取宗地基本信息---------------------------------------
    sqltexts = "SELECT " & strTableZD & "." & strKeyFiledName & "," & strTableZD & ".ZDGUID, " & strTableZD & ".YBZDDM, " & strTableZD & ".QLLX, " & strTableZD & ".QLXZ, " _
     & strTableZD & ".TDQSLY, " & strTableZD & ".ZL, " & strTableZD & ".QLSDFS, " & strTableZD & ".GMJJHYFLDM, " & strTableZD & ".BDCDYH, " & strTableZD & ".TFH, " _
     & strTableZD & ".ZDSZB, " & strTableZD & ".ZDSZD, " & strTableZD & ".ZDSZN, " & strTableZD & ".ZDSZX, " & strTableZD & ".DJ, " & strTableZD & ".JG, " _
     & strTableZD & ".PZYT, " & strTableZD & ".YT, " & strTableZD & ".PZMJ, " & strTableZD & ".ZDMJ, " & strTableZD & ".JianZZDMJ, " & strTableZD & ".JianZZMJ, " _
     & strTableZD & ".ShiYQX, " & strTableZD & ".GYQLRQK, " & strTableZD & ".ZDJBXXSM, " & strTableZD & ".JZDWSM, " & strTableZD & ".ZYQSJXZXSM, " _
     & strTableZD & ".QSDCJS, " & strTableZD & ".DJCLJS, " & strTableZD & ".DJDCJGSHYJ, " & strTableZD & ".CeLDW, " & strTableZD & ".ZDBLC," & strTableZD & ".DiaoCRQ," _
     & strTableZD & ".CeLRQ," & strTableZD & ".XiangMMC," & strTableZD & ".XiangMFZR," & strTableZD & ".QDSJ," & strTableZD & ".ZZSJ," & strTableZD & ".DJZQMC," & strTableZD & ".ID FROM " & strTableZD & " INNER JOIN GeoAreaTB ON " & strTableZD & ".ID = GeoAreaTB.ID "
    condition = "WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And " & strTableZD & "." & strKeyFiledName & " In(" & strSelectedOutRPT & ")"
    strSQL = sqltexts & " " & condition
    GetSQLRecordAll mdbName,strSQL,arRecordZD,nRecordZDCount
    
    '-----GetOutContent.3.2---读取权利人信息---------------------------------------
    sqltexts = "SELECT QLR_权利人信息表.GLGUID, " & strTableZD & ".ZDDM, QLR_权利人信息表.SXH, QLR_权利人信息表.QLRMC, QLR_权利人信息表.QLRLX, QLR_权利人信息表.ZJZL, QLR_权利人信息表.ZJH, QLR_权利人信息表.DZ, QLR_权利人信息表.FaRDB, QLR_权利人信息表.FaRDBZJZL, QLR_权利人信息表.FaRDBZJH, QLR_权利人信息表.FaRDBDH, QLR_权利人信息表.DaiLRXM, QLR_权利人信息表.DaiLRZJZL, QLR_权利人信息表.DaiLRZJH, QLR_权利人信息表.DaiLRDH FROM GeoAreaTB INNER JOIN (" & strTableZD & " INNER JOIN QLR_权利人信息表 ON " & strTableZD & ".ZDGUID = QLR_权利人信息表.GLGUID) ON GeoAreaTB.ID = " & strTableZD & ".ID "
    condition = "WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And QLR_权利人信息表.GLGUID In (" & strSelectedGUIDOutRPT & ")"        '(((QLR_权利人信息表.GLGUID) In ({guid {780E4627-10E8-40F8-897A-DA71483AF6D8}},{guid {F05AC15E-CA2A-490D-AC46-30DD9DE4380C}}))
    strOrderBy = "ORDER BY QLR_权利人信息表.SXH"
    strSQL = sqltexts & " " & condition & " " & strOrderBy
    'addloginfo strSQL
    'GetSQLRecordAll mdbName,strSQL,arRecordQLR,arRecordQLR                '向函数中传递带GUID的参数提示出错，未找到原因--变量错误
    
    '读取权利人信息
    SSProcess.OpenAccessRecordset mdbName, strSQL
    '获取记录总数
    RecordCount = SSProcess.GetAccessRecordCount (mdbName, strSQL)
    If RecordCount > 0 Then
        nRecordQLRCount = RecordCount
        iRecordCount = 0
        ReDim arRecordQLR(RecordCount)
        '将记录游标移到第一行
        SSProcess.AccessMoveFirst mdbName, strSQL
        iRecordCount = 0
        '浏览记录
        While SSProcess.AccessIsEOF (mdbName, strSQL) = 0
            fields = ""
            values = ""
            '获取当前记录内容
            SSProcess.GetAccessRecord mdbName, strSQL, fields, values
            arRecordQLR(iRecordCount) = values                                        '查询记录
            iRecordCount = iRecordCount + 1                                                    '查询记录数
            '移动记录游标
            SSProcess.AccessMoveNext mdbName, strSQL
        WEnd
    End If
    '关闭记录集
    SSProcess.CloseAccessRecordset mdbName, strSQL
    '-------------------------------------------------------------
    Rem MsgBox     iRecordCount
    '-----GetOutContent.3.3---界址签章表记录数组---------------------------------------
    sqltexts = "SELECT " & strTableJZQZ & ".ZDGUID, " & strTableJZQZ & ".JZDHQ, " & strTableJZQZ & ".JZDHZJ, " & strTableJZQZ & ".JZDHZ, " & strTableJZQZ & ".LZQLRMC, " & strTableJZQZ & ".LZZJR, " & strTableJZQZ & ".BZZJR, " & strTableJZQZ & ".JZQZRQ FROM " & strTableJZQZ
    condition = "WHERE " & strTableJZQZ & ".ZDGUID In(" & strSelectedGUIDOutRPT & ")"
    strOrderBy = "ORDER BY " & strTableJZQZ & ".ZDGUID, " & strTableJZQZ & ".ID"
    strSQL = sqltexts & " " & condition & " " & strOrderBy
    GetSQLRecordAll mdbName,strSQL,arRecordJZQZB,nRecordJZQZBCount
    
    
    '-----GetOutContent.3.4---宗地共用分摊面积记录数组---------------------------------------
    sqltexts = "SELECT " & strTableGYMJFT & ".ZDGUID, " & strTableGYMJFT & ".DZWDM, " & strTableGYMJFT & ".QLMJ, " & strTableGYMJFT & ".DYTDMJ, " & strTableGYMJFT & ".FTTDMJ FROM " & strTableGYMJFT
    condition = "WHERE " & strTableGYMJFT & ".ID<>0 And " & strTableGYMJFT & ".ZDGUID In(" & strSelectedGUIDOutRPT & ")"
    strOrderBy = "ORDER BY " & strTableGYMJFT & ".ZDGUID, " & strTableGYMJFT & ".ID"
    strSQL = sqltexts & " " & condition & " " & strOrderBy
    Rem MsgBox strSQL
    GetSQLRecordAll mdbName,strSQL,arRecordGYMJFT,nRecordGYMJFTCount
    Rem MsgBox nRecordGYMJFTCount
    '当共用分摊表没有内容时，从户表提取分摊内容
    If nRecordGYMJFTCount =  - 1 Then
        sqltexts = "SELECT " & strTableHu & ".ZDGUID, Right([BDCDYH],9) , " & strTableHu & ".DYTDMJ, " & strTableHu & ".FTTDMJ FROM GeoAreaTB INNER JOIN " & strTableHu & " ON GeoAreaTB.ID = " & strTableHu & ".ID"
        condition = "WHERE " & strTableHu & ".ZDGUID In(" & strSelectedGUIDOutRPT & ")" & " AND (" & strTableHu & ".HXH like '*复式1层*'  or  " & strTableHu & ".HXH not like '*复式*') And (([GeoAreaTB].[Mark] Mod 2)<>0)"
        strOrderBy = "ORDER BY " & strTableHu & ".ZDGUID, Right([BDCDYH],9)"
        strSQL = sqltexts & " " & condition & " " & strOrderBy
        GetSQLRecordAll mdbName,strSQL,arRecordGYMJFT,nRecordGYMJFTCount
        Rem MsgBox "B：" & nRecordGYMJFTCount
    End If
    
    '-----GetOutContent.3.5---获取户信息---------------------------------------
    sqltexts = "SELECT FC_户信息属性表.HGUID, FC_户信息属性表.ZRZGUID, FC_户信息属性表.BDCDYH, FC_户信息属性表.ZL, FC_户信息属性表.FWXZ, FC_户信息属性表.ChanB, FC_户信息属性表.FWYT1, FC_户信息属性表.FWYT2," _
     & " FC_户信息属性表.FWYT3, FC_户信息属性表.GHYT, FC_户信息属性表.ZRZH, FC_户信息属性表.HH, FC_户信息属性表.ZCS, FC_户信息属性表.CH, FC_户信息属性表.FWJG, FC_户信息属性表.JGSJ, FC_户信息属性表.JZMJ, " _
     & " FC_户信息属性表.TNJZMJ, FC_户信息属性表.FTJZMJ, FC_户信息属性表.CQLY, FC_户信息属性表.FTTDMJ, FC_户信息属性表.DQ, FC_户信息属性表.NQ, FC_户信息属性表.XQ, FC_户信息属性表.BQ, " _
     & "FC_户信息属性表.YCJZMJ, FC_户信息属性表.YCTNJZMJ, FC_户信息属性表.YCFTJZMJ, FC_户信息属性表.SCJZMJ, FC_户信息属性表.SCTNJZMJ, FC_户信息属性表.SCFTJZMJ, FC_户信息属性表.ZDGUID " _
     & " FROM GeoAreaTB INNER JOIN FC_户信息属性表 ON GeoAreaTB.ID = FC_户信息属性表.ID "
    condition = "WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And FC_户信息属性表.ZDGUID In (" & strSelectedGUIDOutRPT & ")"        '(((QLR_权利人信息表.GLGUID) In ({guid {780E4627-10E8-40F8-897A-DA71483AF6D8}},{guid {F05AC15E-CA2A-490D-AC46-30DD9DE4380C}}))
    'strOrderBy ="ORDER BY FC_户信息属性表.SWBH"
    strSQL = sqltexts & " " & condition
    ' addloginfo "FC_户信息属性表-strSQL=" & strSQL
    GetSQLRecordAll mdbName,strSQL,arRecordHU,nRecordHUCount
    
    '-----GetOutContent.3.6---获取自然幢信息---------------------------------------
    'sqltexts ="SELECT FC_自然幢信息属性表.ZRZGUID, FC_自然幢信息属性表.XMMC, FC_自然幢信息属性表.ZTS, FC_自然幢信息属性表.ZZDMJ, FC_自然幢信息属性表.DCRY, FC_自然幢信息属性表.DCRQ, FC_自然幢信息属性表.CHZT FROM FC_自然幢信息属性表 INNER JOIN GeoAreaTB ON FC_自然幢信息属性表.ID = GeoAreaTB.ID "
    'condition ="WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And FC_自然幢信息属性表.ZDGUID In (" & strSelectedGUIDOutRPT & ")"
    sqltexts = "SELECT FC_自然幢信息属性表.ZDDM, FC_自然幢信息属性表.XMMC, FC_自然幢信息属性表.ZTS, FC_自然幢信息属性表.ZZDMJ, FC_自然幢信息属性表.DCRY,  FC_自然幢信息属性表.DCRQ, FC_自然幢信息属性表.CHZT," _
     & " FC_自然幢信息属性表.ZRZH,FC_自然幢信息属性表.JGRQ, FC_自然幢信息属性表.SCJZMJ, FC_自然幢信息属性表.YCJZMJ, FC_自然幢信息属性表.FWJGNAME, FC_自然幢信息属性表.ZCS, FC_自然幢信息属性表.ZTS," _
     & " FC_自然幢信息属性表.JZWJBYT, FC_自然幢信息属性表.GHYTNAME, FC_自然幢信息属性表.ZRZGUID, FC_自然幢信息属性表.BDCDYH,FC_自然幢信息属性表.DQTGS, " _
     & "FC_自然幢信息属性表.NQTGS,FC_自然幢信息属性表.XQTGS,FC_自然幢信息属性表.BQTGS,FC_自然幢信息属性表.CQLY,FC_自然幢信息属性表.FTMJ ,FC_自然幢信息属性表.FWXZ,FC_自然幢信息属性表.ZRZH,FC_自然幢信息属性表.JZMJ,FC_自然幢信息属性表.TNJZMJ,FC_自然幢信息属性表.GYJZMJ,FC_自然幢信息属性表.ChanB,FC_自然幢信息属性表.GHYT,FC_自然幢信息属性表.FWYT,FC_自然幢信息属性表.FWXZ,FC_自然幢信息属性表.SFZYYT,FC_自然幢信息属性表.ZCS,FC_自然幢信息属性表.DSCS,FC_自然幢信息属性表.DXCS FROM FC_自然幢信息属性表 INNER JOIN GeoAreaTB ON FC_自然幢信息属性表.ID = GeoAreaTB.ID "
    condition = "WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 And FC_自然幢信息属性表.ZDGUID In (" & strSelectedGUIDOutRPT & ")"
    strOrderBy = "ORDER BY FC_自然幢信息属性表.ZRZH"
    strSQL = sqltexts & " " & condition & " " & strOrderBy
    'addloginfo "strSQL=" & strSQL
    GetSQLRecordAll mdbName,strSQL,arRecordZRZ,nRecordZRZCount
    
    Dim nZDMID,strZDMCode,nZDMDS            '宗地面ID，宗地面要素编码，宗地面点数
    Dim strQLRBJ                         '权利人信息标记 1 有 0 无
    Dim nJZXCD                                 '界址线长度
    For i_rpt = 0 To nOutSelectedCount - 1
        strQLRBJ = "0"                                     '权利人信息标记 1 有 0 无
        Mark_ZD_BDCDYH = 0                                '宗地不动产单元号
        Mark_ShareTable_Del =  - 1                '设置共有/共用表的删除标记的初值
        '-----BG.1.1-----宗地基本信息表-----------------
        '从宗地记录数组中查宗地基本信息
        GetIndexOnly_StringInclude arOutRecSelected(i_rpt),arRecordZD,nRecordZDCount,Index_ZDXX
        arZDJBXX_Temp = Split(arRecordZD(Index_ZDXX),",")
        nZDJBXXColCount_Temp = UBound(arZDJBXX_Temp)
        '获取宗地面的图形基本属性
        nZDMID = arZDJBXX_Temp(nZDJBXXColCount_Temp)                                                  '宗地面ID
        strZDMCode = SSProcess.GetObjectAttr(nZDMID, "SSObj_Code")                   '宗地面编码
        nZDMDS = SSProcess.GetObjectAttr( nZDMID, "SSObj_PointCount")            '宗地面点数
        nDZWDYS = SSProcess.GetObjectAttr( nZDMID, "[GYMJFTDZWDYS]")                '定着物单元数据
        strZDGYTFSM = SSProcess.GetObjectAttr( nZDMID, "[GYGYZDMJFTBZ]")    '共用面积分摊说明
        strXMMC = SSProcess.GetObjectAttr( nZDMID, "[XiangMMC]")
        YeWLX = SSProcess.GetObjectAttr( nZDMID, "[YeWLX]")
        If YeWLX = "房开项目"  Then
            templateDocName = SSProcess.GetSysPathName(8) & "输出模板\不动产测量报告模板_房开.docx"
        ElseIf YeWLX = "民房项目"  Then
            templateDocName = SSProcess.GetSysPathName(8) & "输出模板\不动产测量报告模板_民房.docx"
        ElseIf  YeWLX = "单一产权"  Then
            templateDocName = SSProcess.GetSysPathName(8) & "输出模板\不动产测量报告模板_厂房.docx"
        Else
            MsgBox "业务类型为空，请在项目信息录入中填写！"
            Exit Sub
        End If
        
        'msgbox templateDocName
        If ReportFileStatus(templateDocName) = "不存在" Then
            MsgBox "报表模板不存在，请核实！" & Chr(13) & Chr(13) & templateDocName,48
            Exit Sub
        End If
        initDocCom  templateDocName
        docObj.CreateDocumentByTemplate templateDocName
        
        resultFileName = CreateSavePath()
        initDB()
        
        'strDCY =SSProcess.GetObjectAttr( nZDMID, "[DiaoCR]")                '调查员
        'strCLY =SSProcess.GetObjectAttr( nZDMID, "[CeLY]")                    '测量员
        'strSHR =SSProcess.GetObjectAttr( nZDMID, "[ShenHR]")                '审核人
        'strSHRQ =SSProcess.GetObjectAttr( nZDMID, "[ShenHRQ]")        '审核日期
        'strZHITY =SSProcess.GetObjectAttr( nZDMID, "[ZhiTY]")        '制图员
        
        Rem arRecordJZBSB(),nRecordJZBSBCount
        '-----BG.2-----界址标示表-----------------
        ReDim arRecordJZBSB(nZDMDS,22)            '重定义界址表数组大小
        GetJZDinfo nZDMID,nZDMDS,arRecordJZBSB,nRecordJZBSBCount
        
        '-----BG.3-----获取界址签章表-----------------
        GetIndexMul arOutRecGUIDSelected(i_rpt),arRecordJZQZB,nRecordJZQZBCount,index_JZQZB
        
        '-----BG.4-----共有共用宗地面积分摊信息表-----------------
        GetIndexMul arOutRecGUIDSelected(i_rpt),arRecordGYMJFT,nRecordGYMJFTCount,index_GYMJFT
        
        '---------------------------------开始输出报告----------------------------------------------
        
        '-------WriteRPT.3.2-宗地基本信息表-------------------------
        docObj.MoveToSection 2
        WriteZDJBXXB   arZDJBXX_Temp, arRecordQLR,nRecordQLRCount ,arOutRecGUIDSelected(i_rpt),strZDMCode
        
        'GetIndexMul_SpecifyColumn arOutRecGUIDSelected(i_rpt),arRecordHU,31,nRecordHUCount,index_HU
        
        '-------WriteRPT.3.2-房屋基本信息调查表----------------------------------
        docObj.MoveToSection 3
        TableNum = 0
        
        WriteFWJBXXDCB nZDMID,arRecordZRZ, nRecordZRZCount, index_ZRZ
        
        '-------WriteRPT.3.3-填写界址标示表----------------------------------
        docObj.MoveToSection 4
        WriteJZDZBB arRecordJZBSB,nRecordJZBSBCount,nResPageNum
        
        '-------WriteRPT.3.4-获取界址签章表----------------------------------
        nResPageNum_JZQZ = nResPageNum + 1    '界址签章表号
        If index_JZQZB <> "" Then
            arIndexJZQZTemp = Split(index_JZQZB,",")
            nJZQZCount_Temp = UBound(arIndexJZQZTemp)
            If nJZQZCount_Temp > 19 Then  docObj.CloneTableRow nResPageNum_JZQZ, 6, 1, nJZQZCount_Temp - 19
            For i_JZQZ = 0 To nJZQZCount_Temp
                arJZQZTemp = Split(arRecordJZQZB(i_JZQZ),",")
                SetCellValue nResPageNum_JZQZ,i_JZQZ + 4, 0, arJZQZTemp(1)        '填写单元格内容
                SetCellValue nResPageNum_JZQZ,i_JZQZ + 4, 1, arJZQZTemp(2)
                SetCellValue nResPageNum_JZQZ,i_JZQZ + 4, 2, arJZQZTemp(3)
                SetCellValue nResPageNum_JZQZ,i_JZQZ + 4, 3, arJZQZTemp(4)
                SetCellValue nResPageNum_JZQZ,i_JZQZ + 4, 4, arJZQZTemp(5)
                SetCellValue nResPageNum_JZQZ,i_JZQZ + 4, 5, arJZQZTemp(6)
                SetCellValue nResPageNum_JZQZ,i_JZQZ + 4, 6, arJZQZTemp(7)
            Next
        End If
        
        docObj.MoveToSection 7
        '地籍调查界址点成果表
        WriteDJDCJZDCGB arRecordJZBSB,nRecordJZBSBCount,nResPageNum,nZDMID
        'msgbox "6"
        '宗地测绘成果面积明细表
        If YeWLX = "房开项目"  Or YeWLX = "民房项目"  Then
            docObj.MoveToSection 8
            TableNum = 0
            strzdmj = SSProcess.GetObjectAttr(nZDMID, "[ZDMJ]")
            strJZZDMJ = SSProcess.GetObjectAttr(nZDMID, "[JianZZDMJ]")
            If strXMMC <> "" And strXMMC <> "*" Then SetCellValue TableNum,1, 0,  strXMMC
            If strZDMJ <> "" And strZDMJ <> "*" Then SetCellValue TableNum,1, 1,  FormatNumber(strZDMJ,2, - 1,0,0)             '宗地面积
            If strJZZDMJ <> "" And strJZZDMJ <> "*" Then SetCellValue TableNum,1, 2,  FormatNumber(strJZZDMJ,2, - 1,0,0) '建筑占地面积
            SetCellValue TableNum,1, 3,  FormatNumber((strZDMJ - strJZZDMJ),2, - 1,0,0) '宗地内绿化及道路面积=宗地面积-建筑占地面积
            SetCellValue TableNum,15, 3,  FormatNumber((strZDMJ - strJZZDMJ),2, - 1,0,0)
            SetCellValue TableNum,15, 1,  FormatNumber(strZDMJ,2, - 1,0,0)
            SetCellValue TableNum,15, 2,  FormatNumber(strJZZDMJ,2, - 1,0,0)
            
            '分幢建筑占地面积明细表
            docObj.MoveToSection 9
            WriteFZJZZDMJMXB nZDMID, arRecordZRZ,nRecordZRZCount,index_ZRZ
        End If
        'msgbox "7"
        If YeWLX = "房开项目"  Then
            '建筑物区分所有权业主共有部分登记信息
            docObj.MoveToSection 10
            WriteGYDJXXB  nZDMID, arOutRecGUIDSelected(i_rpt)
            '各基本单元不动产面积分摊表
            docObj.MoveToSection 11
            WriteJBDYBDCMJFTB nZDMID,arRecordZRZ,nRecordZRZCount,index_ZRZ
        End If
        
        If YeWLX = "民房项目"  Then
            '各基本单元不动产面积分摊表
            docObj.MoveToSection 10
            WriteJBDYBDCMJFTB nZDMID,arRecordZRZ,nRecordZRZCount,index_ZRZ
        End If
        'msgbox "8"
        '分摊系数 幢占地面积/幢建筑面积  土地用途-使用功能
        ' savedoc()
        ' exit sub         
        SSProcess.CloseAccessMdb mdbName
        docObj.UpdateFields
        strOutputPath = resultFileName & arOutRecSelected(i_rpt) & "_不动产测量报告.doc"
        'addloginfo "strOutputPath=" & strOutputPath
        docObj.SaveEx strOutputPath
        ReleaseDB()
    Next
    
    MsgBox "输出完成!"
    
End Sub

'宗地基本信息表
Function  WriteZDJBXXB (ByVal arZDJBXX_Temp,ByVal arRecordQLR,ByVal nRecordQLRCount,ByVal strZDGUID,strZDMCode)
    
    TableNum = 0
    nZDMID = arZDJBXX_Temp(UBound(arZDJBXX_Temp))
    strZDDM = arZDJBXX_Temp(0)            '宗地代码
    strYuBZDDM = arZDJBXX_Temp(2)    '宗地预编代码
    strQLLX = arZDJBXX_Temp(3)            '权利类型
    strQLXZ = arZDJBXX_Temp(4)            '权利性质
    strTDQSLY = arZDJBXX_Temp(5)        '土地权属来源证明材料
    strZL = arZDJBXX_Temp(6)                '土地坐落
    strQLSDFS = arZDJBXX_Temp(7)            '权利设定方式
    strGMJJHYFLDM = arZDJBXX_Temp(8)        '国民经济行业分类代码
    strBDCDYH = arZDJBXX_Temp(9)            '不动产单元号
    strTFH = arZDJBXX_Temp(10)            '图幅号
    strZDSZB = arZDJBXX_Temp(11)            '宗地四至北
    strZDSZD = arZDJBXX_Temp(12)            '宗地四至东
    strZDSZN = arZDJBXX_Temp(13)            '宗地四至南
    strZDSZX = arZDJBXX_Temp(14)            '宗地四至西
    strDJ = arZDJBXX_Temp(15)            '等级
    strJG = arZDJBXX_Temp(16)            '价格
    strPZYT = arZDJBXX_Temp(17)            '批准用途
    strSJYT = arZDJBXX_Temp(18)            '实际用途
    strPZMJ = arZDJBXX_Temp(19)            '批准面积
    strZDMJ = arZDJBXX_Temp(20)            '宗地面积
    strJZZDMJ = arZDJBXX_Temp(21)        '建筑占地总面积
    strJZZMJ = arZDJBXX_Temp(22)            '建筑总面积。小数点后保留2位。宗地内若有地下建筑物，地上建筑物与地下建筑物应分别填写建筑物总面积。用“/”作为分隔符。
    '如“1000.00/300.00”，其中，“1000.00”表示宗地地上建筑物总面积，“300.00”表示地下建筑物总面积。
    strShiYQX = arZDJBXX_Temp(23)        '使用期限
    strGYQLRQK = arZDJBXX_Temp(24)        '共有权利人情况
    strZDJBXXSM = arZDJBXX_Temp(25)        '宗地基本信息说明
    strJZDWSM = arZDJBXX_Temp(26)        '界址点位置说明
    strZYQSJXZXSM = arZDJBXX_Temp(27)    '主要权属界线走向说明
    strQSDCJS = arZDJBXX_Temp(28)        '权属调查记事
    strDJCLJS = arZDJBXX_Temp(29)        '地籍测量记事
    strDJDCJGSHYJ = arZDJBXX_Temp(30)    '地籍调查结果审核意见
    strCLDW = arZDJBXX_Temp(31)            '测量单位
    strZDBLC = arZDJBXX_Temp(32)            '宗地比例尺
    strDCRQ = arZDJBXX_Temp(33)            '宗地调查日期
    strCLRQ = arZDJBXX_Temp(34)            '宗地测量日期
    strXMMC = arZDJBXX_Temp(35)            '项目名称
    strXMFZR = arZDJBXX_Temp(36)            '项目负责人
    strQSSJ = arZDJBXX_Temp(37)            '取得时间
    strZZSJ = arZDJBXX_Temp(38)            '终止时间
    strDJZQ = arZDJBXX_Temp(40)            '制图员 
    
    
    If strZDBLC = "" And strZDBLC = "*" Then
        strZDBLC = "500"
    End If
    '国民经济行业分类名称
    strTemplateFileName = SSProcess.GetTemplateFileName ()
    If strGMJJHYFLDM <> "" And strGMJJHYFLDM <> "*" Then
        strFilename = Left(strTemplateFileName,Len(strTemplateFileName) - 4) & "\GMJJHYFLDM.DIC"
        GetMatchAttr2  strFilename , strGMJJHYFLDM,  strGMJJHYFLMC
    End If
    
    '-----BG.1.2-----宗地基本表：权利人-----------------
    '获取权利人信息
    GetIndexMul strZDGUID ,arRecordQLR,nRecordQLRCount,index_QLR
    If index_QLR <> "" Then
        arQLRIndexTemp = Split(index_QLR,",")
        nCountQLR = UBound(arQLRIndexTemp)
        arQLRInfoTemp = Split(arRecordQLR(arQLRIndexTemp(0)),",")
        nCount_Temp = UBound(arQLRInfoTemp)
        '权利人查询内容：GLGUID, " & strTableZD & ".ZDDM, SXH, QLRMC, QLRLX, ZJZL, ZJH, DZ, FaRDB, FaRDBZJZL, FaRDBZJH, FaRDBDH, DaiLRXM, DaiLRZJZL, DaiLRZJH, DaiLRDH
        If nCountQLR > 0 Then
            strQLRMC = arQLRInfoTemp(3) & "等"                '权利人名称
        Else
            strQLRMC = arQLRInfoTemp(3)                '权利人名称
        End If
        strQLRLX = arQLRInfoTemp(4)                '权利人类型
        strQLRZJZL = arQLRInfoTemp(5)            '权利人证件类型
        strQLRZJH = arQLRInfoTemp(6)                '权利人证件号
        strQLRDZ = arQLRInfoTemp(7)                '权利人通讯地址
        strFaRDB = arQLRInfoTemp(8)                '法人代表名称
        strFaRDBZJZL = arQLRInfoTemp(9)            '法人代表证件类型
        strFaRDBZJH = arQLRInfoTemp(10)            '法人代表证件号
        strFaRDBDH = arQLRInfoTemp(11)            '法人代表电话
        strDaiLRXM = arQLRInfoTemp(12)            '代理人名称
        strDaiLRZJZL = arQLRInfoTemp(13)        '代理人证件类型
        strDaiLRZJH = arQLRInfoTemp(14)            '代理人证件号
        strDaiLRDH = arQLRInfoTemp(15)            '代理人电话
        strQLRBJ = "1"                            '权利人信息标记 1 有 0 无
    End If
    
    'addloginfo GetQLLXMC(strQLLX) & ","& strQLLX&"," &GetQLXZM(strQLXZ)& ","& strQLLX
    If strQLLX <> "" And strQLLX <> "*" Then             SetCellValue TableNum,6, 1, GetQLLXMC(strQLLX)                '权利类型
    If strQLXZ <> "" And strQLXZ <> "*" Then             SetCellValue TableNum,6, 3, GetQLXZM(strQLXZ)                '权利性质 
    
    If strTDQSLY <> "" And strTDQSLY <> "*" Then         SetCellValue TableNum,6,  5, strTDQSLY                '土地权属来源证明材料
    If strZL <> "" And strZL <> "*" Then
        SetCellValue TableNum,7,  1,  strZL                    '土地坐落
        ReplaceOneStr "{土地坐落}",  strZL
    Else
        ReplaceOneStr "{土地坐落}",  ""
    End If
    If strQLSDFS <> "" And strQLSDFS <> "*" Then             SetCellValue TableNum,12, 1, GetQLSDFSM(strQLSDFS)               '权利设定方式
    If strGMJJHYFLDM <> "" And strGMJJHYFLDM <> "*" Then     SetCellValue TableNum,13, 1, strGMJJHYFLMC                       '国民经济行业分类代码
    If strZDDM <> "" And strZDDM <> "*" Then                 SetCellValue TableNum,14, 3, strZDDM                            '宗地代码
    If strYuBZDDM <> "" And strYuBZDDM <> "*" Then        SetCellValue TableNum,14, 1, strYuBZDDM                            '宗地预编代码
    'if strBDCDYH <>"" And strBDCDYH <>"*" then             SetCellValue TableNum,15, 1, strBDCDYH                           '不动产单元号
    '宗地比例尺
    If strZDBLC <> "" And strZDBLC <> "*" Then
        If InStr(strZDBLC,":") Or InStr(strZDBLC,"：") Then
            SetCellValue TableNum,15, 2, strZDBLC
        Else
            SetCellValue TableNum,15, 2, "1:" & strZDBLC
        End If
    Else
        SetCellValue TableNum,15, 2, "1:500"
    End If
    If strTFH <> "" And strTFH <> "*" Then                 SetCellValue TableNum,16, 2, Replace(strTFH,"|","、")                               '图幅号
    If strZDSZB <> "" And strZDSZB <> "*" Then             SetCellValue TableNum,17, 2, strZDSZB                   '宗地四至北
    If strZDSZD <> "" And strZDSZD <> "*" Then             SetCellValue TableNum,18, 2, strZDSZD                   '宗地四至东
    If strZDSZN <> "" And strZDSZN <> "*" Then             SetCellValue TableNum,19, 2, strZDSZN                   '宗地四至南
    If strZDSZX <> "" And strZDSZX <> "*" Then             SetCellValue TableNum,20, 2, strZDSZX                   '宗地四至西
    
    'if strDJ <>"" And strDJ <>"*" then                     SetCellValue TableNum,23, 2, strDJ                               '等级
    'if strJG <>"" And strJG >0.0 then                     SetCellValue TableNum,23, 4, strJG                               '价格
    
    If strPZYT <> "" And strPZYT <> "*" Then
        strPZYTMC = GetTDLYLXMC(strPZYT)
        
        SetCellValue TableNum,21, 1, strPZYTMC                                       '批准用途名
        SetCellValue TableNum,22, 2, strPZYT                                           '批准用途码
    End If
    If strSJYT <> "" And strSJYT <> "*" Then
        strSJJYTMC = GetTDLYLXMC(strSJYT)
        SetCellValue TableNum,21, 3, strSJJYTMC                                   '实际用途名
        SetCellValue TableNum,22, 5, strSJYT                                           '实际用途码
    End If
    If strPZMJ <> "" Then
        If strPZMJ > 0.0 Then
            SetCellValue TableNum,23, 1, FormatNumber(strPZMJ,2, - 1,0,0)                                       '批准面积
        End If
    End If
    If strZDMJ <> "" And strZDMJ > 0.0 Then
        SetCellValue TableNum,23, 3, FormatNumber(strZDMJ,2, - 1,0,0)                                           '宗地面积
        ReplaceOneStr "{ZDMJ}",  FormatNumber(strZDMJ,2, - 1,0,0)
    End If
    'QDSJ ZZSJ
    
    If strJZZDMJ <> "" And strJZZDMJ > 0.0 Then
        strJZZDMJ = FormatNumber(strJZZDMJ,2, - 1,0,0)
    Else
        strJZZDMJ = 0
        SetCellValue TableNum,23, 5, strJZZDMJ                                                           '建筑占地总面积
    End If
    If strJZZMJ <> "" And strJZZMJ > 0.0 Then
        strJZZMJ = FormatNumber(strJZZMJ,2, - 1,0,0)
        SetCellValue TableNum,24, 5, strJZZMJ                                                           '建筑总面积
        ReplaceOneStr "{JZMJ}",  FormatNumber(strJZZMJ,2, - 1,0,0)
    End If
    If strQSSJ <> "" And strQSSJ <> "*" Then         SetCellValue TableNum,25, 1,  FormatDateTime(strQSSJ,1)          '使用期限-起始时间
    If strZZSJ <> "" And strZZSJ <> "*" Then         SetCellValue TableNum,25, 3, FormatDateTime(strZZSJ,1)          '使用期限 -终点
    If strGYQLRQK <> "" And strGYQLRQK <> "*" Then         SetCellValue TableNum,26, 1, strGYQLRQK           '共有权利人情况
    If strZDJBXXSM <> "" And strZDJBXXSM <> "*" Then         SetCellValue TableNum,27, 1, strZDJBXXSM           '宗地基本信息说明
    
    If strJZDWSM <> "" And strJZDWSM <> "*" Then
        ReplaceOneStr "{界址点位说明}", strJZDWSM                 '界址点位置说明
    Else
        ReplaceOneStr "{界址点位说明}", ""
    End If
    If strZYQSJXZXSM <> "" And strZYQSJXZXSM <> "*" Then
        ReplaceOneStr "{主要权属界线走向说明}", strZYQSJXZXSM     '主要权属界线走向说明
    Else
        ReplaceOneStr "{主要权属界线走向说明}", "无"
    End If
    'if strQSDCJS <>"" And strQSDCJS <>"*" then
    'ReplaceOneStr "{权属调查记事}", strQSDCJS                 '权属调查记事
    'else
    'ReplaceOneStr "{权属调查记事}", ""
    'end if
    'if strDJCLJS <>"" And strDJCLJS <>"*" then
    'ReplaceOneStr "{地籍测量记事}", strDJCLJS                 '地籍测量记事
    'else
    'ReplaceOneStr "{地籍测量记事}", ""
    'end if
    If strDJDCJGSHYJ <> "" And strDJDCJGSHYJ <> "*" Then
        ReplaceOneStr "{地籍调查结果审核意见}", strDJDCJGSHYJ     '地籍调查结果审核意见
    Else
        ReplaceOneStr "{地籍调查结果审核意见}", "无"
    End If
    If strQLRMC <> "" And strQLRMC <> "*" Then ReplaceOneStr "{QLR}", strQLRMC              '权利人
    
    '-------WriteRPT.3.2.1-填写权利人----------------------
    If strQLRBJ = "1" Then
        If strQLRMC <> "" And strQLRMC <> "*" Then
            
            If strZDMCode = "9130223" Then
                strQLLX = GetQLLXMC(strQLLX)
                If InStr(strQLLX,"国家") > 0 Or  InStr(strQLLX,"国有") > 0 Then
                    SetCellValue TableNum,1, 2, "国家"
                Else
                    SetCellValue TableNum,1, 2, strDJZQ & "村民集体"
                End If
                SetCellValue TableNum,2, 2, strQLRMC                   '权利人名称
            Else
                SetCellValue TableNum,1, 2, strQLRMC                   '权利人名称
            End If
        End If
        If strQLRLX <> "" And strQLRLX <> "*" Then         SetCellValue TableNum,2, 4, GetQLRLXM(strQLRLX)                    '权利人类型
        If strQLRZJZL <> "" And strQLRZJZL <> "*" Then          SetCellValue TableNum,3, 4, GetZJZLM(strQLRZJZL)            '权利人证件类型
        If strQLRZJH <> "" And strQLRZJH <> "*" Then         SetCellValue TableNum,4, 4, strQLRZJH                            '权利人证件号
        If strQLRDZ <> "" And strQLRDZ <> "*" Then         SetCellValue TableNum,5, 4, strQLRDZ                            '权利人通讯地址
        
        If strFaRDB <> "" And strFaRDB <> "*" Then                 SetCellValue TableNum,8, 1, strFaRDB                    '法人代表名称
        If strFaRDBZJZL <> "" And strFaRDBZJZL <> "*" Then         SetCellValue TableNum,8, 3, GetZJZLM(strFaRDBZJZL)        '法人代表证件类型
        If strFaRDBZJH <> "" And strFaRDBZJH <> "*" Then             SetCellValue TableNum,9, 3, strFaRDBZJH                '法人代表证件号
        If strFaRDBDH <> "" And strFaRDBDH <> "*" Then             SetCellValue TableNum,8, 5, strFaRDBDH                    '法人代表电话
        
        If strDaiLRXM <> "" And strDaiLRXM <> "*" Then             SetCellValue TableNum,10, 1, strDaiLRXM                    '代理人名称
        If strDaiLRZJZL <> "" And strDaiLRZJZL <> "*" Then         SetCellValue TableNum,10, 3, GetZJZLM(strDaiLRZJZL)        '代理人证件类型
        If strDaiLRZJH <> "" And strDaiLRZJH <> "*" Then             SetCellValue TableNum,11, 3, strDaiLRZJH                '代理人证件号
        If strDaiLRDH <> "" And strDaiLRDH <> "*" Then             SetCellValue TableNum,10, 5, strDaiLRDH                    '代理人电话
    End If
    
    strDiaoCR = SSProcess.GetObjectAttr( nZDMID, "[DiaoCR]")            '调查员
    strDCRQ = SSProcess.GetObjectAttr( nZDMID, "[DiaoCRQ]")            '调查员
    strCLY = SSProcess.GetObjectAttr( nZDMID, "[CeLY]")                '测量员
    strSHR = SSProcess.GetObjectAttr( nZDMID, "[ShenHR]")            '审核人
    strSHRQ = SSProcess.GetObjectAttr( nZDMID, "[ShenHRQ]")    '审核日期
    strZHITY = SSProcess.GetObjectAttr( nZDMID, "[ZhiTY]")        '制图员
    strLBBZSZH = SSProcess.GetObjectAttr( nZDMID, "[LBBZSZH]")        '土地证号
    
    ReplaceOneStr "{ZDDM}", strZDDM                 '宗地代码
    ReplaceOneStr "{XiangMMC}", strXMMC         '项目名称
    ReplaceOneStr "{XiangMFZR}", strXMFZR             '项目负责人
    
    'if strDCRQ <>"" then     strDCRQ =FormatDateTime(replace(strDCRQ,"-","/"),1)   
    ReplaceOneStr "{DiaoCRQ}", strDCRQ         '调查日期
    
    'if strCLRQ <>"" then  strCLRQ =FormatDateTime(replace(strCLRQ,"-","/"),1)   
    ReplaceOneStr "{CeLRQY}", Left(strCLRQ,InStr(strCLRQ,"月"))                 '测量日期   
    ReplaceOneStr "{CeLRQ}",     strCLRQ     '测量日期   
    ReplaceOneStr "{ZhiTY}", strZHITY        '制图员
    ReplaceOneStr "{ShenHY}",strSHR              '审核员
    ReplaceOneStr "{ZDMJ}",FormatNumber(strZDMJ,2, - 1,0,0)              '宗地面积
    ReplaceOneStr "{JZMJ}",FormatNumber(strJZZMJ,2, - 1,0,0) '建筑面积
    ReplaceOneStr "{CeLY}", strCLY          '调查员
    ReplaceOneStr "{DiaoCY}", strDiaoCR          '调查员
    ReplaceOneStr "{QLR}", strQLRMC              '权利人
    ReplaceOneStr "{TDZH}",      strTDQSLY         '土地证号--等确认      
    
End Function

'获取宗地界址点信息
Function GetJZDinfo(ByVal nZDMID,ByVal nZDMDS,ByRef arRecordJZBSB,ByRef nRecordJZBSBCount)
    nRecordJZBSBCount = 0
    nJZXCD = 0                                    '界址线长度
    '获取界址标示表内容
    
    For j = 0 To nZDMDS - 1
        SSProcess.GetObjectPoint nZDMID, j, x1,  y1,  z1,  ptype1,  name1
        'msgbox arRecordJZBSB(j,1)
        idsP = SSProcess.SearchNearObjIDs(x1, y1, 0.001, 0, DKJZDBM, 0 )
        
        '--JSBSB.1---计算界址间距------------
        If j > 0 And j <= nZDMDS - 1 Then
            SSProcess.GetObjectPoint nZDMID, j - 1, x0,  y0,  z0,  ptype0,  name0
            '计算宗地边的距离和方位角
            flag = 0
            SSProcess.XYSA x0,y0,x1,y1,dist,angleA,flag
            dist = FormatNumber(dist,2, - 1,0,0)
            '当前界址线段长度
            nJZXCD = nJZXCD + dist
        End If
        '--JSBSB.2---获取界址点内容------------
        If idsP <> "" Then
            arID = Split(idsP, ",")
            '获取界址点号
            strJZDH = ""
            If JZDHHQFS = "1" Then
                If j <> nZDMDS - 1 Then
                    strJZDH = "J" & nRecordJZBSBCount + 1
                Else
                    strJZDH = "J1"
                End If
                Rem MsgBox j & Chr(13) & strJZDH
                arRecordJZBSB(nRecordJZBSBCount,1) = strJZDH
            ElseIf JZDHHQFS = "2" Then
                PointName_JZD = SSProcess.GetObjectAttr(arID(0), "[JZDH]" )
                arRecordJZBSB(nRecordJZBSBCount,1) = "J" & PointName_JZD
            End If
            JBLX = SSProcess.GetObjectAttr(arID(0), "[JBLX]" )
            If JBLX <> "" Then  '钢钉 水泥钉 石灰桩 喷涂 瓷标志 无标记 其它
                If JBLX = "1" Then arRecordJZBSB(nRecordJZBSBCount,2) = "√"            '钢钉
                If JBLX = "2" Then arRecordJZBSB(nRecordJZBSBCount,3) = "√"            '水泥桩
                If JBLX = "3" Then arRecordJZBSB(nRecordJZBSBCount,4) = "√"            '石灰桩
                If JBLX = "4" Then arRecordJZBSB(nRecordJZBSBCount,5) = "√"            '喷涂
                If JBLX = "5" Then arRecordJZBSB(nRecordJZBSBCount,6) = "√"            '瓷标志
                If JBLX = "6" Then arRecordJZBSB(nRecordJZBSBCount,7) = "√"            '无标记             
            Else
                arRecordJZBSB(nRecordJZBSBCount,8) = "√"  '其它
            End If
            
            If j > 0 Then
                arRecordJZBSB(j - 1 ,9) = FormatNumber(nJZXCD,2, - 1,0,0)            '界址间距
                'addloginfo "j=" &j & ",arRecordJZBSB(j-1 ,1)=" & arRecordJZBSB(j-1 ,1) & ",arRecordJZBSB(j-1 ,9)=" & arRecordJZBSB(j-1 ,9)
                nJZXCD = 0
            End If
            '--JSBSB.3---获取界址线内容------------
            If j < nZDMDS - 1 Then
                '计算当前宗地边(段)的中点坐标
                SSProcess.GetObjectPoint nZDMID, j + 1, x2,  y2,  z2,  ptype2,  name2
                x = (x1 + x2) / 2
                y = (y1 + y2) / 2
                '使用界址边的中点查找界址线
                idsL = SSProcess.SearchNearObjIDs(x, y, 0.001, 1, DKJXBM, 0 )
                If idsL <> "" Then
                    arJZXIDTemp = Split(idsL, ",")
                    JZXLB = SSProcess.GetObjectAttr(arJZXIDTemp(0), "[JZXLB]" )        '界址线类别
                    JZXWZ = SSProcess.GetObjectAttr(arJZXIDTemp(0), "[JZXWZ]" )        '界址线位置
                    arRecordJZBSB(nRecordJZBSBCount + 1,21) = SSProcess.GetObjectAttr(arJZXIDTemp(0), "[JZXSM]" )            '界址线说明
                    '界址线类别
                    Rem 1,围墙 2,墙壁  3,栅栏  4,铁丝网  5,滴水线  6,路涯线  7,两点连线  9,其它
                    If JZXLB = "1"  Then arRecordJZBSB(nRecordJZBSBCount,10) = "√"            '围墙
                    If JZXLB = "2"  Then arRecordJZBSB(nRecordJZBSBCount,11) = "√"            '墙壁
                    If JZXLB = "3"  Then arRecordJZBSB(nRecordJZBSBCount,12) = "√"            '栅栏
                    If JZXLB = "4"  Then arRecordJZBSB(nRecordJZBSBCount,13) = "√"            '铁丝网
                    If JZXLB = "5"  Then arRecordJZBSB(nRecordJZBSBCount,14) = "√"              '滴水线
                    If JZXLB = "6"  Then arRecordJZBSB(nRecordJZBSBCount,15) = "√"            '路涯线
                    If JZXLB = "7"  Then arRecordJZBSB(nRecordJZBSBCount,16) = "√"            '两点连线
                    If JZXLB = "9"  Then arRecordJZBSB(nRecordJZBSBCount,17) = "√"        '其它
                    '界址线位置
                    '1-内，2-中，3-外
                    If JZXWZ = "1" Then         arRecordJZBSB(nRecordJZBSBCount,18) = "√"
                    If JZXWZ = "2" Then         arRecordJZBSB(nRecordJZBSBCount,19) = "√"
                    If JZXWZ = "3" Then         arRecordJZBSB(nRecordJZBSBCount,20) = "√"
                End If
            End If
            nRecordJZBSBCount = nRecordJZBSBCount + 1                        '输出界址点总数
            Rem MsgBox "查找：" & nRecordJZBSBCount
        End If
    Next
    ReplaceOneStr "{JZDZS}", nRecordJZBSBCount - 1     '界址点总数  
End Function

'房屋基本信息调查表
Function WriteFWJBXXDCB (ByVal zdid,ByVal arRecordZRZ,ByVal nRecordZRZCount,ByVal index_ZRZ)
    tableNum = 0
    strzdguid = SSProcess.GetObjectAttr( zdid, "[ZDGUID]")
    GetZDinfo zdinfoAr, zdinfoCount,strzdguid
    '权利人信息
    GetQLRinfo qlrinfoAr,qlrinfoCount,strzdguid 'SXH, QLRMC, QLRLX, ZJZL, ZJH, DZ, DH, YB, GYFSNAME, HZGX
    strQLRMC = ""
    strQLRLX = ""
    strQLRZJZL = ""
    strQLRZJH = ""
    strQLRDZ = ""
    strQLRDH = ""
    strQLRYB = ""
    strQLRGYQK = ""
    qlrbs = 0
    For i = 0 To qlrinfoCount - 1
        'SXH, QLRMC, QLRLX, ZJZL, ZJH, DZ, DH, YB, GYFSNAME, HZGX
        If qlrinfoAr(i,9) = "是"  Then
            qlrbs = 1
            strQLRMC = qlrinfoAr(i,1)                        ' 权利人名称 
            strQLRLX = qlrinfoAr(i,2)                        ' 权利人类型
            strQLRZJZL = qlrinfoAr(i,3)                    ' 权利人证件类型
            strQLRZJH = qlrinfoAr(i,4)                    ' 权利人证件号
            strQLRDZ = qlrinfoAr(i,5)                        ' 权利人通讯地址
            strQLRDH = qlrinfoAr(i,6)                        ' 权利人联系方式
            strQLRYB = qlrinfoAr(i,7)                        ' 权利人邮编
            strQLRGYQK = qlrinfoAr(i,8)                    ' 权利人共有情况
            Exit For
        End If
    Next
    If qlrbs = 0  Then
        strQLRMC = qlrinfoAr(0,1)                        ' 权利人名称 
        strQLRLX = qlrinfoAr(0,2)                        ' 权利人类型
        strQLRZJZL = qlrinfoAr(0,3)                    ' 权利人证件类型
        strQLRZJH = qlrinfoAr(0,4)                    ' 权利人证件号
        strQLRDZ = qlrinfoAr(0,5)                        ' 权利人通讯地址
        strQLRDH = qlrinfoAr(0,6)                        ' 权利人联系方式
        strQLRYB = qlrinfoAr(0,7)                        ' 权利人邮编
        strQLRGYQK = qlrinfoAr(0,8)                    ' 权利人共有情况
    End If
    'strQLRLX = SSRETools.GetDictionaryCode("QLR_权利人信息表", "QLRLX", strQLRLX)            '权利人证件类型 GetQLRLXM(huxx(5))    
    'if strQLRLX = "" then strQLRLX = "/"
    'strQLRZJZL = SSRETools.GetDictionaryCode("QLR_权利人信息表", "ZJZL", strQLRZJZL)        '证件类型
    'if strQLRZJZL = "" then strQLRZJZL = "/"
    
    zddm = SSProcess.GetObjectAttr( zdid, "[ZDDM]")
    '获取自然幢的相关属性
    GetIndexMul zddm,arRecordZRZ,nRecordZRZCount,index_ZRZ
    'addloginfo  "index_ZRZ=" & index_ZRZ & ",nRecordZRZCount=" & nRecordZRZCount 
    If nRecordZRZCount > 9 Then docObj.CloneTableRow tableNum, 15, 1, nRecordZRZCount - 9
    strZDDM = zdinfoAr(0,0)                    ' 宗地代码
    strQXDM = zdinfoAr(0,1)                    ' 区县代码
    strDJQDM = zdinfoAr(0,2)                    ' 地籍区代码
    strDJZQDM = zdinfoAr(0,3)                    ' 地籍子区代码
    strJianZZDZMJ = zdinfoAr(0,4)                ' 建筑占地面积
    strJianZZMJ = zdinfoAr(0,5)                ' 建筑总面积
    strDiaoCRY = zdinfoAr(0,6)                ' 调查人员
    strDiaoCRQ = zdinfoAr(0,7)                ' 调查日期
    strZL = zdinfoAr(0,8)                            ' 座落
    strYT = zdinfoAr(0,9)                           '用途
    strXMMC = zdinfoAr(0,10)       '项目名称
    strGYQK = zdinfoAr(0,11)       '共有情况
    If Len(strZDDM) = 19 Then
        strQXDM = Left(strZDDM, 6)            ' 区县代码
        strDJQDM = Mid(strZDDM, 7, 3)        ' 地籍区代码
        strDJZQDM = Mid(strZDDM, 10, 3)        ' 地籍子区代码
    End If
    If IsNumeric(strJianZZDZMJ) Then strJianZZDZMJ = FormatNumber(strJianZZDZMJ, 2, - 1)
    If IsNumeric(strJianZZMJ) Then strJianZZMJ = FormatNumber(strJianZZMJ, 2, - 1)
    SetCellValue TableNum,3, 1, ""                            '不动产单元号
    SetCellValue TableNum,2, 1,strQXDM                     '市区名称
    SetCellValue TableNum,2, 3, strDJQDM         '地籍区
    SetCellValue TableNum,2, 5,strDJZQDM                      '地籍子区
    SetCellValue TableNum,2, 7,CStr(Right(strZDDM,7))                     '宗地号
    SetCellValue TableNum,2, 9, ""                     '定着物
    If strZL = "*" Then strZL = ""
    SetCellValue TableNum,4, 1, strZL                            '房地坐落  
    If strGYQK = "*" Then strGYQK = ""
    SetCellValue TableNum,7, 5, strGYQK                        '共有情况
    If strXMMC = "*" Then strXMMC = ""
    SetCellValue TableNum,8, 3, strXMMC                        '项目名称   
    If strQLRMC = "*" Then strQLRMC = ""
    SetCellValue TableNum,5, 1, strQLRMC                            '权利人
    If strQLRZJZL = "*" Then strQLRZJZL = ""
    SetCellValue TableNum,5, 3, GetZJZLM(strQLRZJZL)                        '证件种类
    If strQLRZJH = "*" Then strQLRZJH = ""
    SetCellValue TableNum,6, 3, strQLRZJH                            '证件号
    If strQLRDH = "*" Then strQLRDH = ""
    SetCellValue TableNum,7, 1, strQLRDH                            '电话
    If strQLRDZ = "*" Then strQLRDZ = ""
    SetCellValue TableNum,7, 3, strQLRDZ                        '住址
    If strQLRLX = "*" Then strQLRLX = ""
    SetCellValue TableNum,8, 1, GetQLRLXM(strQLRLX)                    '权利人类型               
    
    Dim zdzts,zdzdmj,zdjzmj,zdtnjzmj,zdgyjzmj
    zdzts = 0
    zdzdmj = 0
    zdjzmj = 0
    zdtnjzmj = 0
    zdgyjzmj = 0
    
    strcbs = ""
    strghyts = ""
    strfwyts = ""
    strfwxzs = ""
    strszc = ""
    If index_ZRZ <> "" Then
        ARRK = Split(index_ZRZ,"," )
        For k = 0 To UBound(ARRK)
            strzrzinfo = arRecordZRZ(ARRK(k))
            strzrzinfo = Replace(strzrzinfo,"*","")
            'addloginfo "arRecordZRZ(ARRK(k))=" & arRecordZRZ(ARRK(k)) 
            arZRZTemp = Split(arRecordZRZ(ARRK(k)),",")
            strZRZXMMC = arZRZTemp(1)    '自然幢.项目名称
            strZTS = arZRZTemp(13)            '自然幢.总套数
            strZZDMJ = arZRZTemp(3)        '自然幢.占地面积
            strDCRY = arZRZTemp(4)            '自然幢.调查人员
            strDCRQ = arZRZTemp(5)            '自然幢.调查日期
            strCHZT = arZRZTemp(6)            '自然幢.测绘状态
            strZRZH = arZRZTemp(25)            '自然幢.自然幢幢号
            strJGSJ = arZRZTemp(8)            '自然幢.竣工日期
            strcb = Replace(arZRZTemp(29) ,"*","")
            strghyt = Replace(arZRZTemp(30) ,"*","")
            strfwyt = Replace(arZRZTemp(31) ,"*","")
            strfwxz = Replace(arZRZTemp(32) ,"*","")
            strsfzyyt = Replace(arZRZTemp(33) ,"*","")
            If strZRZH <> "地下室"  Then
                strDSCS = CDbl(arZRZTemp(35))
                If strDSCS = 1 Then strszc = 1 Else strszc = "1-" & strDSCS
            Else
                strDXCS = CDbl(arZRZTemp(36))
                If strDXCS = 1 Then strszc =  - 1 Else strszc = "-1--" & strDXCS
            End If
            If strcb <> ""  Then
                If strcbs = ""  Then
                    strcbs = "|" & strcb & "|"
                Else
                    If InStr(strcbs,"|" & strcb & "|") = 0  Then strcbs = strcbs & ",|" & strcb & "|"
                End If
            End If
            If strfwxz <> ""  Then
                If strfwxzs = ""  Then
                    strfwxzs = "|" & strfwxz & "|"
                Else
                    If InStr(strfwxzs,"|" & strfwxz & "|") = 0  Then strfwxzs = strfwxzs & ",|" & strfwxz & "|"
                End If
            End If
            If strsfzyyt = "1"  Then
                If strghyts = ""  Then
                    strghyts = "|" & strghyt & "|"
                Else
                    If InStr(strghyts,"|" & strghyt & "|") = 0  Then strghyts = strghyts & ",|" & strghyt & "|"
                End If
                
                If strfwyts = ""  Then
                    strfwyts = "|" & strfwyt & "|"
                Else
                    If InStr(strfwyts,"|" & strfwyt & "|") = 0  Then strfwyts = strfwyts & ",|" & strfwyt & "|"
                End If
            End If
            
            '..........规范取值..................
            
            dbJZMJ = arZRZTemp(26)                '建筑面积
            dbTNJZMJ = arZRZTemp(27)
            dbGYJZMJ = arZRZTemp(28)                '专有建筑面积
            If strZTS = "" Or strZTS = "*" Then strZTS = 0
            zdzts = zdzts + CDbl(strZTS)
            
            'strZH =mid(huxx(2),21,4)        '幢号
            '幢占地面积
            If strZZDMJ = "" Or strZZDMJ = "*" Then strZZDMJ = 0
            If strZZDMJ <> 0.0 Then
                strZZDMJ = FormatNumber(strZZDMJ,2, - 1, - 1,0)
                zdzdmj = zdzdmj + strZZDMJ
            Else
                strZZDMJ = ""
            End If
            
            '建筑面积
            If dbJZMJ = "" Or dbJZMJ = "*" Then dbJZMJ = 0
            'if dbJZMJ <>0.0 then
            dbJZMJ = FormatNumber(dbJZMJ,2, - 1, - 1,0)
            zdjzmj = zdjzmj + dbJZMJ
            'else
            'dbJZMJ =0.00
            'end if
            '专有建筑面积
            If dbTNJZMJ = "" Or dbTNJZMJ = "*" Then dbTNJZMJ = 0
            'if dbTNJZMJ <>0.0 then
            dbTNJZMJ = FormatNumber(dbTNJZMJ,2, - 1, - 1,0)
            zdtnjzmj = zdtnjzmj + dbTNJZMJ
            'else
            '    dbTNJZMJ =0.00
            'end if
            '分摊建筑面积
            If dbGYJZMJ = "" Or dbGYJZMJ = "*" Then dbGYJZMJ = 0
            'if cdbl(dbGYJZMJ) <>0.0 then
            dbGYJZMJ = FormatNumber(dbGYJZMJ,2, - 1, - 1,0)
            zdgyjzmj = zdgyjzmj + dbGYJZMJ
            'else
            '    dbGYJZMJ =0.00
            'end if                        
            
            
            SetCellValue TableNum,13 + k, 2,  ""                            '户号
            'SetCellValue TableNum,14+k, 3,  RIGHT(strZRZH,2)                                    '编号
            
            SetCellValue TableNum,13 + k, 3, arZRZTemp(13)                                    '总套数
            
            SetCellValue TableNum,13 + k, 4, arZRZTemp(12)                                    '总层数
            SetCellValue TableNum,13 + k, 5, strszc                                    '所在层
            'addloginfo "arZRZTemp(12)=" & arZRZTemp(12) &",RIGHT(huxx(2),4)=" & RIGHT(huxx(2),4)& ",arZRZTemp(13)="& arZRZTemp(13)    
            
            
            SetCellValue TableNum,13 + k, 6,  arZRZTemp(11)                                       '房屋结构
            SetCellValue TableNum,13 + k, 7, Left(strJGSJ    ,4)                       '竣工时间
            SetCellValue TableNum,13 + k, 9, dbJZMJ                           '建筑面积
            If YeWLX = "房开项目"  Then
                SetCellValue TableNum,13 + k, 10, dbTNJZMJ                           '专有建筑面积
                SetCellValue TableNum,13 + k, 11, dbGYJZMJ                           '分摊建筑面积
            End If
            If arZRZTemp(22) = "*"  Then arZRZTemp(22) = ""
            SetCellValue TableNum,13 + k, 12, arZRZTemp(22)                     '产权来源
            
            SetCellValue TableNum,13 + k, 13,   arZRZTemp(18)                      '墙体归属东
            
            SetCellValue TableNum,13 + k, 14, arZRZTemp(19)                  '墙体归属南
            
            SetCellValue TableNum,13 + k, 15, arZRZTemp(20)                   '墙体归属西
            
            SetCellValue TableNum,13 + k, 16, arZRZTemp(21)                   '墙体归属北
            ' SaveDOC()
            'exit function
            '自然幢信息
            
            SetCellValue TableNum,13 + k, 1, strZRZH                         '幢号
            If YeWLX = "房开项目"  Or YeWLX = "单一产权"  Then
                SetCellValue TableNum,13 + k, 8, strZZDMJ                     '占地面积.自然幢
            End If
            'addloginfo "strZZDMJ=" & strZZDMJ
            '获取当前表的总行数
            
        Next
        strcbsmc = ""
        strghytsmc = ""
        strfwytsmc = ""
        strfwxzsmc = ""
        
        strTemplateFileName = SSProcess.GetTemplateFileName ()
        
        strcbs = Replace(strcbs,"|","")
        strghyts = Replace(strghyts,"|","")
        strfwyts = Replace(strfwyts,"|","")
        strfwxzs = Replace(strfwxzs,"|","")
        If strcbs <> "" Then
            artemp = Split(strcbs,",")
            For i = 0 To UBound(artemp)
                strFilename = Left(strTemplateFileName,Len(strTemplateFileName) - 4) & "\ChanB.DIC"
                GetMatchAttr2  strFilename , artemp(i),  strcbMC
                If strcbsmc = ""  Then
                    strcbsmc = strcbMC
                Else
                    strcbsmc = strcbsmc & "、" & strcbMC
                End If
            Next
            SetCellValue TableNum,9, 3, strcbsmc                    '产别
        End If
        
        If strghyts <> "" Then
            artemp = Split(strghyts,",")
            For i = 0 To UBound(artemp)
                strFilename = Left(strTemplateFileName,Len(strTemplateFileName) - 4) & "\GHYT.DIC"
                GetMatchAttr2  strFilename , artemp(i),  strghytMC
                If strghytsmc = ""  Then
                    strghytsmc = strghytMC
                Else
                    strghytsmc = strghytsmc & "、" & strghytMC
                End If
            Next
            SetCellValue TableNum,10, 3, strghytsmc          '规划用途
        End If
        
        If strfwyts <> "" Then
            artemp = Split(strfwyts,",")
            For i = 0 To UBound(artemp)
                strFilename = Left(strTemplateFileName,Len(strTemplateFileName) - 4) & "\FWYT.DIC"
                GetMatchAttr2  strFilename , artemp(i),  strfwytMC
                If strfwytsmc = ""  Then
                    strfwytsmc = strfwytMC
                Else
                    strfwytsmc = strfwytsmc & "、" & strfwytMC
                End If
            Next
            SetCellValue TableNum,10, 1, strfwytsmc            '用途
        End If
        
        If strfwxzs <> "" Then
            artemp = Split(strfwxzs,",")
            For i = 0 To UBound(artemp)
                strFilename = Left(strTemplateFileName,Len(strTemplateFileName) - 4) & "\FWXZ.DIC"
                GetMatchAttr2  strFilename , artemp(i),  strfwxzMC
                If strfwxzsmc = ""  Then
                    strfwxzsmc = strfwxzMC
                Else
                    strfwxzsmc = strfwxzsmc & "、" & strfwxzMC
                End If
            Next
            
            SetCellValue TableNum,9, 1,strfwxzsmc            '房屋性质'
        End If
        
        
        
        nTableRowCount = docobj.GetTableRowCount(TableNum)
        '合计
        SetCellValue TableNum,nTableRowCount - 4, 1, "合计"
        SetCellValue TableNum,nTableRowCount - 4, 3, zdzts
        If YeWLX = "房开项目"  Or YeWLX = "单一产权"  Then
            SetCellValue TableNum,nTableRowCount - 4, 8, zdzdmj
        End If
        SetCellValue TableNum,nTableRowCount - 4, 9, zdjzmj
        If YeWLX = "房开项目"  Then
            SetCellValue TableNum,nTableRowCount - 4, 10, zdtnjzmj
            SetCellValue TableNum,nTableRowCount - 4, 11, zdgyjzmj
        End If
        ' msgbox "c：" & nTableRowCount
        'SetCellValue TableNum,nTableRowCount -3, 3, "/"               '附加说明                待处理
        'SetCellValue TableNum,nTableRowCount -2,3, "/"                   '调查意见
        
        'if strDCRY ="" or strDCRY ="*" then strDCRY ="/"
        'SetCellValue TableNum,nTableRowCount-1, 1, strDCRY               '自然幢.调查人员
        
        'if strDCRQ <>"" then
        '    strDCRQ =FormatDateTime(strDCRQ,1)
        '    SetCellValue TableNum,nTableRowCount-1, 4, strDCRQ               '自然幢.调查日期
        'else
        '    SetCellValue TableNum,nTableRowCount-1, 4, "  / 年 / 月 / 日"               '自然幢.调查日期
        'end if
        
    End If
    
    'SaveDOC()
End Function

'界址点坐标表
Function WriteJZDZBB(ByVal arRecordJZBSB,ByVal nRecordJZBSBCount,ByRef nResPageNum)
    TableNum = 0                '表号
    nStartRowNum = 3        '开始行号      
    '-------WriteRPT.3.3.1-计算页（表）数--------------------------------
    CalculatePagesCount nResPageCount,nRecordJZBSBCount,17,17,17
    '-------WriteRPT.3.3.2-复制表-----------------------------------------
    'addloginfo "nRecordJZBSBCount=" & nRecordJZBSBCount & ",nResPageCount=" & nResPageCount
    
    If nResPageCount > 1 Then
        For p = 1 To nResPageCount - 1
            docobj.CloneTable tableNum,0
        Next
    End If
    'exit function
    nStartPageNum = TableNum
    nRowFirstPageCount = 17
    nStartRowNumFirstPage = 3
    nRowPageCount = 17
    nStartRowNumPage = 3
    nRowEndPageCount = 17
    
    For i_WriteJZBS = 0 To nRecordJZBSBCount - 1
        nRowRecordNum = i_WriteJZBS + 1
        '-------WriteRPT.3.3.3-计算表号和行号----------------------------
        '计算记录对应行号，参数：返回行号，返回页号，记录行号，起始页(表)号，表格首页行数，表格首页起始行号，中间页行数，中间页起始行号，末页行数
        CalculatePageRowNum nResRowNum, nResPageNum, nRowRecordNum, nStartPageNum, nRowFirstPageCount, nStartRowNumFirstPage, nRowPageCount, nStartRowNumPage, nRowEndPageCount
        'addloginfo "i_WriteJZBS=" & i_WriteJZBS& ",nResRowNum=" & nResRowNum 
        For j_WriteJZBS = 1 To 21
            strCELLValCur = arRecordJZBSB(i_WriteJZBS,j_WriteJZBS)                    '当前单元格内容
            'addloginfo "i_WriteJZBS=" & i_WriteJZBS & ",j_WriteJZBS=" &j_WriteJZBS & ",strCELLValCur=" &strCELLValCur
            If strCELLValCur <> "" Then
                '-------WriteRPT.3.3.4-计算行号-----跳行填写行号计算-----------------------
                If nResRowNum = nStartRowNum Then            '计算行号
                    nRowNumCur = nStartRowNum
                Else
                    If j_WriteJZBS <= 8 Then
                        nRowNumCur = (nStartRowNum - 1) + ((nResRowNum - nStartRowNum) * 2)
                    Else
                        nRowNumCur = (nStartRowNum - 0) + ((nResRowNum - nStartRowNum) * 2)
                    End If
                End If
                'addloginfo "nResPageNum=" & nResPageNum & ",nRowNumCur=" &nRowNumCur & ",j_WriteJZBS=" & j_WriteJZBS& ",strCELLValCur=" & strCELLValCur
                SetCellValue nResPageNum,nRowNumCur, j_WriteJZBS - 1, strCELLValCur        '填写单元格内容  
                'if nRowNumCur<>(nRowEndPageCount+nStartRowNumPage)  then             
            End If
        Next
    Next
    If nResPageCount = 0 Then nResPageNum = nStartPageNum                                '如界址标示表记录为0，页码的累加
End Function


'地籍调查界址点成果表
Function WriteDJDCJZDCGB(ByVal arRecordJZBSB,ByVal nRecordJZBSBCount,ByRef nResPageNum,ByVal nZDMID)
    nResPageNum = 0
    TableNum = 0                '表号
    
    '-------WriteRPT.3.3.1-计算页（表）数--------------------------------
    CalculatePagesCount nResPageCount,nRecordJZBSBCount,14,14,14
    '-------WriteRPT.3.3.2-复制表-----------------------------------------
    'addloginfo "nRecordJZBSBCount=" & nRecordJZBSBCount & ",nResPageCount=" & nResPageCount
    
    If nResPageCount > 1 Then
        For p = 1 To nResPageCount - 1
            'docobj. CloneTableEx(TableNum)
            docobj.CloneTable tableNum,0
        Next
    End If
    nStartPageNum = TableNum
    nStartRowNum = 4    '开始行号 
    nRowFirstPageCount = 14
    nStartRowNumFirstPage = 4
    nRowPageCount = 14
    nStartRowNumPage = 4
    nRowEndPageCount = 14
    'SetCellValue nResPageNum,4, 1, "ff"
    'SetCellValue nResPageNum,5, 11, "5-11"
    ' SetCellValue nResPageNum,7, 11, "7-11"
    For i_WriteJZBS = 0 To nRecordJZBSBCount - 1
        
        SSProcess.GetObjectPoint nZDMID, i_WriteJZBS, x1,  y1,  z1,  ptype1,  name1
        'SetCellValue nResPageNum,nRowNumCur, j_WriteJZBS-1, strCELLValCur         
        nRowRecordNum = i_WriteJZBS + 1
        '-------WriteRPT.3.3.3-计算表号和行号----------------------------
        '计算记录对应行号，参数：返回行号，返回页号，记录行号，起始页(表)号，表格首页行数，表格首页起始行号，中间页行数，中间页起始行号，末页行数
        CalculatePageRowNum nResRowNum, nResPageNum, nRowRecordNum, nStartPageNum, nRowFirstPageCount, nStartRowNumFirstPage, nRowPageCount, nStartRowNumPage, nRowEndPageCount
        'addloginfo "i_WriteJZBS=" & i_WriteJZBS& ",nResRowNum=" & nResRowNum 
        For j_WriteJZBS = 1 To 21
            strCELLValCur = arRecordJZBSB(i_WriteJZBS,j_WriteJZBS)                    '当前单元格内容
            'addloginfo "i_WriteJZBS=" & i_WriteJZBS & ",j_WriteJZBS=" &j_WriteJZBS & ",strCELLValCur=" &strCELLValCur
            If strCELLValCur <> "" Then
                '-------WriteRPT.3.3.4-计算行号-----跳行填写行号计算-----------------------
                If nResRowNum = nStartRowNum Then            '计算行号                 
                    nRowNumCur = nStartRowNum
                Else
                    If j_WriteJZBS <= 8 Then
                        nRowNumCur = (nStartRowNum - 0) + ((nResRowNum - nStartRowNum) * 2)
                    Else
                        nRowNumCur = (nStartRowNum + 1) + ((nResRowNum - nStartRowNum) * 2)
                    End If
                End If
                'addloginfo "nResRowNum=" & nResRowNum & ",nRowNumCur=" &nRowNumCur & ",j_WriteJZBS=" & j_WriteJZBS& ",strCELLValCur=" & strCELLValCur
                If j_WriteJZBS = 1 Then
                    SetCellValue nResPageNum,nRowNumCur, 0, strCELLValCur
                    SetCellValue nResPageNum,nRowNumCur, 1, FormatNumber(y1,3, - 1, - 1,0)
                    SetCellValue nResPageNum,nRowNumCur, 2, FormatNumber(x1,3, - 1, - 1,0)
                Else
                    SetCellValue nResPageNum,nRowNumCur, j_WriteJZBS + 1, strCELLValCur        '填写单元格内容
                End If
            End If
        Next
    Next
    If nResPageCount = 0 Then nResPageNum = nStartPageNum                                '如界址标示表记录为0，页码的累加
    
    
End Function

'各基本单元不动产面积分摊表
Function WriteJBDYBDCMJFTB (ByVal nZDMID,ByVal arRecordZRZ,ByVal nRecordZRZCount,ByRef index_ZRZ)
    zddm = SSProcess.GetObjectAttr( nZDMID, "[ZDDM]")
    strZDGUID = SSProcess.GetObjectAttr( nZDMID, "[ZDGUID]")
    tableNum = 0
    nStartRowNum = 3
    '获取自然幢的相关属性
    GetIndexMul zddm,arRecordZRZ,nRecordZRZCount,index_ZRZ
    If index_ZRZ <> "" Then
        ARRK = Split(index_ZRZ,"," )
        totalrows = 33
        If UBound(ARRK) > 1 Then
            ' msgbox "UBOUND(ARRK)=" & UBOUND(ARRK)
            For k = UBound(ARRK) To 0 step - 1
                'addloginfo "CloneTableEx-k=" & k
                docObj.CloneTable tableNum,0
                strzrzinfo = arRecordZRZ(ARRK(k))
                strzrzinfo = Replace(strzrzinfo,"*","")
                arZRZTemp = Split(strzrzinfo,",")
                strJZWMC = arZRZTemp(UBound(arZRZTemp) - 1)            '自然幢.自然幢幢号
                ' addloginfo "strJZWMC=" & strJZWMC
                ' SetCellValue tablenum,0, 0, strJZWMC& "各基本单元不动产面积分摊表"
            Next
            docobj.DeleteTable UBound(ARRK)
        End If
        
        Dim totalzzdmj
        For k = 0 To  UBound(ARRK)
            ' if k>0 then docObj.CloneTable  tableNum,0
            strzrzinfo = arRecordZRZ(ARRK(k))
            strzrzinfo = Replace(strzrzinfo,"*","")
            'addloginfo "arRecordZRZ(ARRK(k))=" & arRecordZRZ(ARRK(k)) 
            arZRZTemp = Split(strzrzinfo,",")
            strZRZH = arZRZTemp(7)            '自然幢.自然幢幢号
            strJZWMC = arZRZTemp(25)
            SetCellValue tablenum,0, 0, strJZWMC & "各基本单元不动产面积分摊表"
            
            strZZDMJ = arZRZTemp(3)       '幢占地面积
            strZJZMJ = arZRZTemp(26)'幢建筑面积    
            If strZZDMJ = ""  Or strZZDMJ = "*"  Then strZZDMJ = 0
            If strZJZMJ = ""  Or strZJZMJ = "*"  Then strZJZMJ = 0
            strzrzguid = arZRZTemp(16)    'ZRZGUID
            If strZJZMJ <> "" And strZJZMJ <> "*" And strZJZMJ <> 0 Then strFTXS = CDbl(strZZDMJ) / CDbl(strZJZMJ)      '幢占地面积/幢建筑面积
            newrow = 0
            GetHuinfo  huinfoAr, huinfoCount,strZDGUID, strZRZGUID 'FH,TNJZMJ,SYFGN,FTTDMJ
            If huinfoCount > 66 Then   newrow = (huinfoCount / 2) - 33
            If huinfoCount > 66 Then docObj.CloneTableRow tableNum, 3, 1, Int(newrow * ( - 1)) / ( - 1) '户过多一页写不完时,插入行 　　　　　　
            totalrows = 32 + Int(newrow * ( - 1)) / ( - 1) + nStartRowNum
            'addloginfo "strJZWMC=" &strJZWMC & ",huinfoCount=" &huinfoCount& ",totalrows=" & totalrows
            Dim lefttotaljzmj,lefttotalftxs,lefttotalfttdmj,rightttotaljzmj,righttotalftxs,righttotalfttdmj
            lefttotaljzmj = 0
            lefttotalftxs = 0
            lefttotalfttdmj = 0
            rightttotaljzmj = 0
            righttotalftxs = 0
            righttotalfttdmj = 0
            For i = 0 To huinfoCount - 1
                strFH = huinfoAr(i,0)
                strTNMJ = huinfoAr(i,1)
                strSYFGN = huinfoAr(i,2)
                strFTTDMJ = huinfoAr(i,3)
                If strTNMJ = "*"  Then strTNMJ = 0
                If strFTXS = "*"  Then strFTXS = 0
                If strFTTDMJ = "*"  Then strFTTDMJ = 0
                
                If i <= (totalrows - nStartRowNum)Then
                    'addloginfo "i=" & i & ",currow=" & (i+nStartRowNum)& ",strFH=" & strFH
                    If strFH <> "" And strFH <> "*" Then SetCellValue tablenum,i + nStartRowNum, 0, strFH
                    If strTNMJ <> "" And strTNMJ <> "*" Then SetCellValue tablenum,i + nStartRowNum, 1, FormatNumber(CDbl(strTNMJ),2, - 1,0,0)
                    If strFTXS <> "" And strFTXS <> "*" Then SetCellValue tablenum,i + nStartRowNum, 2, FormatNumber(CDbl(strFTXS ),6, - 1,0,0)
                    If strFTTDMJ <> "" And strFTTDMJ <> "*" Then SetCellValue tablenum,i + nStartRowNum, 3, FormatNumber(CDbl(strFTTDMJ),2, - 1,0,0)
                    If strSYFGN <> "" And strSYFGN <> "*" Then SetCellValue tablenum,i + nStartRowNum, 4, strSYFGN
                    
                    lefttotaljzmj = CDbl(lefttotaljzmj) + CDbl(strTNMJ)
                    lefttotalftxs = CDbl(lefttotalftxs) + CDbl(strFTXS)
                    lefttotalfttdmj = CDbl(lefttotalfttdmj) + CDbl(strFTTDMJ)
                Else
                    'addloginfo "currow="& i-(totalrows- nStartRowNum) &",i=" & i & ",strFH=" & strFH
                    If strFH <> "" And strFH <> "*" Then SetCellValue tablenum,i - (totalrows - nStartRowNum) + nStartRowNum - 1, 6, strFH
                    If strTNMJ <> "" And strTNMJ <> "*" Then SetCellValue tablenum,i - (totalrows - nStartRowNum) + nStartRowNum - 1, 7, FormatNumber(CDbl(strTNMJ),2, - 1,0,0)
                    If strFTXS <> "" And strFTXS <> "*" Then SetCellValue tablenum,i - (totalrows - nStartRowNum) + nStartRowNum - 1, 8, FormatNumber(CDbl(strFTXS ),6, - 1,0,0)
                    If strFTTDMJ <> "" And strFTTDMJ <> "*" Then SetCellValue tablenum,i - (totalrows - nStartRowNum) + nStartRowNum - 1, 9, FormatNumber(CDbl(strFTTDMJ),2, - 1,0,0)
                    If strSYFGN <> "" And strSYFGN <> "*" Then SetCellValue tablenum,i - (totalrows - nStartRowNum) + nStartRowNum - 1, 10, strSYFGN
                    rightttotaljzmj = CDbl(rightttotaljzmj) + CDbl(strTNMJ)
                    righttotalftxs = CDbl(righttotalftxs) + CDbl(strFTXS)
                    righttotalfttdmj = CDbl(righttotalfttdmj) + CDbl(strFTTDMJ)
                    
                End If
            Next
            '合计
            
            ' SetCellValue tablenum,30, 1, "test"
            If lefttotaljzmj <> "" And lefttotaljzmj <> "*" And  lefttotaljzmj > 0 Then SetCellValue tablenum,totalrows + 1, 1, FormatNumber(CDbl(lefttotaljzmj),2, - 1,0,0)
            'if lefttotalftxs<>"" and lefttotalftxs<>"*" then SetCellValue tablenum,totalrows+1, 2, formatnumber(cdbl(lefttotalftxs ),2,-1,0,0) 
            If lefttotalfttdmj <> "" And lefttotalfttdmj <> "*" And  lefttotalfttdmj > 0 Then SetCellValue tablenum,totalrows + 1, 3, FormatNumber(CDbl(lefttotalfttdmj),2, - 1,0,0)
            If rightttotaljzmj <> "" And rightttotaljzmj <> "*" And  rightttotaljzmj > 0 Then SetCellValue tablenum,totalrows + 1, 7, FormatNumber(CDbl(rightttotaljzmj),2, - 1,0,0)
            'if righttotalftxs<>"" and righttotalftxs<>"*" then SetCellValue tablenum,totalrows+1, 8, formatnumber(cdbl(righttotalftxs ),2,-1,0,0) 
            If righttotalfttdmj <> "" And righttotalfttdmj <> "*" And  righttotalfttdmj > 0 Then SetCellValue tablenum,totalrows + 1, 9, FormatNumber(CDbl(righttotalfttdmj),2, - 1,0,0)
            'addloginfo "strZJZMJ=" & strZJZMJ & ",strZZDMJ=" & strZZDMJ  & ",ftxs=" & ftxs 
            
            If strZJZMJ <> "" And strZJZMJ <> "*"  And  strZJZMJ > 0 Then SetCellValue tablenum,totalrows + 2, 1, FormatNumber(CDbl(strZJZMJ),2, - 1,0,0)  '建筑面积S1  
            If strZZDMJ <> "" And strZZDMJ <> "*" And  strZZDMJ > 0 Then SetCellValue tablenum,totalrows + 2, 3, FormatNumber(CDbl(strZZDMJ),2, - 1,0,0)  '建筑占地面积S2
            If CDbl(strZJZMJ) <> 0 Then  ftxs = CDbl(strZZDMJ) / CDbl(strZJZMJ)
            If ftxs <> "" And ftxs <> "*" And  ftxs > 0 Then SetCellValue tablenum,totalrows + 2, 6, FormatNumber(CDbl(ftxs),6, - 1,0,0)  '分摊系数K＝S2/S1
            tableNum = tableNum + 1
        Next
        
    End If
    
End Function

'分幢建筑占地面积明细表
Function  WriteFZJZZDMJMXB (ByVal nZDMID, ByVal arRecordZRZ,ByVal nRecordZRZCount,ByVal index_ZRZ)
    tablenum = 0
    zddm = SSProcess.GetObjectAttr( nZDMID, "[ZDDM]")
    '获取自然幢的相关属性
    GetIndexMul zddm,arRecordZRZ,nRecordZRZCount,index_ZRZ
    If index_ZRZ <> "" Then
        ARRK = Split(index_ZRZ,"," )
        totalrows = 26
        If UBound(ARRK) > 26 Then
            docObj.CloneTableRow tableNum, 2, 1, UBound(ARRK) + 1 - 25
            totalrows = UBound(ARRK) + 1
        End If
        totalrows = totalrows + 3
        Dim totalzzdmj
        
        For k = 0 To UBound(ARRK)
            
            strzrzinfo = arRecordZRZ(ARRK(k))
            strzrzinfo = Replace(strzrzinfo,"*","")
            'addloginfo "arRecordZRZ(ARRK(k))=" & arRecordZRZ(ARRK(k)) 
            arZRZTemp = Split(strzrzinfo,",")
            strZRZH = arZRZTemp(25)            '自然幢.自然幢幢号
            strZCS = arZRZTemp(12 )
            strZZDMJ = arZRZTemp(3)       '幢占地面积
            
            If strZRZH <> "" And strZRZH <> "*"   Then SetCellValue tableNum,k + 1, 0,strZRZH
            If strZZDMJ <> "" And strZZDMJ <> "*"   Then SetCellValue tableNum,k + 1, 1,FormatNumber(strZZDMJ,2, - 1,0,0)
            If strZCS <> "" And strZCS <> "*"   Then SetCellValue tableNum,k + 1, 2,strZCS
            totalzzdmj = CDbl(totalzzdmj) + CDbl(strZZDMJ)
            '分摊系数 幢占地面积/幢建筑面积
            
            'if strZCS<>"" and strZCS<>"*" then SetCellValue tableNum,k+1, 0,
        Next
        SetCellValue tableNum,totalrows - 2, 1,FormatNumber(totalzzdmj,2, - 1,0,0)  '合计
        tablenum = tablenum + 1
    End If
End Function

'建筑物区分所有权业主共有部分登记信息
Function WriteGYDJXXB  (ByVal nZDMID, ByVal strZDGUID)
    
    strFields = "ZRZH,JZWMC,FJ,FTTDMJ,JZMJ,DJSJ"
    fieldsCount = 6
    sql = "select  ZRZH,JZWMC,FJ,FTTDMJ,JZMJ,DJSJ from ZD_建筑物区分所有权业主共有部分登记信息表 where ZDGUID IN(" & strZDGUID & ")"
    'addloginfo "GetHUinfo sql=" & sql
    GetMdbValues sql,strFields,fieldsCount,gyinfoAr,gyinfoCount
    tableNum = 0
    nStartRowNum = 4
    Dim totaljzmj,totalfttdmj,totalrows
    totalrows = 18 + 4 + 2
    If gyinfoCount > 18 Then
        docobj.CloneTableRow tableNum, 5, 1, gyinfoCount - 18
        totalrows = totalrows + gyinfoCount - 18
    End If
    'addloginfo "sql=" & sql & ",gyinfoCount=" & gyinfoCount
    
    For i = 0 To gyinfoCount - 1
        
        SetCellValue tableNum,nStartRowNum + i, 0,i + 1
        strZRZH = gyinfoAr(i,0)
        strJZWMC = gyinfoAr(i,1)
        strFJ = gyinfoAr(i,2)
        strFTTDMJ = gyinfoAr(i,3)
        strJZMJ = gyinfoAr(i,4)
        strDSSJ = gyinfoAr(i,5)
        If strZRZH <> "" And strZRZH <> "*"  Then SetCellValue tableNum,nStartRowNum + i, 1,strZRZH
        If strJZWMC <> "" And strJZWMC <> "*"  Then SetCellValue tableNum,nStartRowNum + i, 2,strJZWMC
        If strJZMJ <> "" And strJZMJ <> "*" And strJZMJ > 0 Then SetCellValue tableNum,nStartRowNum + i, 3,FormatNumber(strJZMJ,2, - 1,0,0)
        If strFTTDMJ <> "" And strFTTDMJ <> "*"  Then SetCellValue tableNum,nStartRowNum + i, 4,FormatNumber(strFTTDMJ,2, - 1,0,0)
        If strDSSJ <> "" And strDSSJ <> "*"  Then SetCellValue tableNum,nStartRowNum + i, 5,FormatDateTime(strDSSJ,1)
        If strFJ <> "" And strFJ <> "*"  Then SetCellValue tableNum,nStartRowNum + i,7,strFJ
        totaljzmj = CDbl(totaljzmj) + CDbl(strJZMJ)
        totalfttdmj = CDbl(totalfttdmj) + CDbl(strFTTDMJ)
    Next
    
    If totaljzmj <> "" And totaljzmj <> "*" And totaljzmj > 0 Then SetCellValue tableNum,totalrows - 2, 3,FormatNumber(totaljzmj,2, - 1,0,0)
    If totalfttdmj <> "" And totalfttdmj <> "*"  And totalfttdmj > 0 Then SetCellValue tableNum,totalrows - 2, 4,FormatNumber(totalfttdmj,2, - 1,0,0)
    '宗地内绿化及道路面积=宗地面积-建筑占地面积
    strzdmj = SSProcess.GetObjectAttr(nZDMID, "[ZDMJ]")
    strJZZDMJ = SSProcess.GetObjectAttr(nZDMID, "[JianZZDMJ]")
    strXMMC = SSProcess.GetObjectAttr(nZDMID, "[XiangMMC]")
    SetCellValue tableNum,totalrows - 1, 1,FormatNumber(totaljzmj,2, - 1,0,0)
    SetCellValue tableNum,1, 1,strXMMC & "全体业主共有"
End Function

Function SaveDOC()
    
    resultFileName = CreateSavePath()
    strOutputPath = resultFileName & arOutRecSelected(i_rpt) & "_不动产测量报告.doc"
    
    docObj.SaveEx strOutputPath
    mdbName = SSProcess.GetProjectFileName
    SSProcess.CloseAccessMdb mdbName
    ReleaseDB()
    MsgBox "save"
End Function

Function GetZDinfo(ByRef zdinfoAr,ByRef  zdinfoCount,ByVal strFeatureGUID)
    strFields = "ZDDM,QXDM,DJQDM,DJZQDM,JianZZDMJ,JianZZMJ,DiaoCR,DiaoCRQ,ZL,YTNAME,XiangMMC,GYQLRQK"
    fieldsCount = 12
    sql = "select ZDDM, QXDM, DJQDM, DJZQDM,JianZZDMJ,JianZZMJ,DiaoCR,DiaoCRQ,ZL,YTNAME,XiangMMC,GYQLRQK from ZD_宗地基本信息属性表 inner join GeoAreaTB on ZD_宗地基本信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 and ZDGUID IN(" & strFeatureGUID & ") ORDER BY ZDDM"
    'addloginfo "GetZDinfo sql=" & sql
    GetMdbValues sql,strFields,fieldsCount,zdinfoAr,zdinfoCount
End Function

Function GetHUinfo(ByRef huinfoAr,ByRef  huinfoCount,ByVal strZDGUID,ByVal strZRZGUID)
    strFields = "FH,SCJZMJ,SYGN,FTTDMJ"
    fieldsCount = 4
    sql = "select FH,SCJZMJ,SYGN,FTTDMJ from FC_户信息属性表 inner join GeoAreaTB on FC_户信息属性表.ID = GeoAreaTB.ID WHERE (GeoAreaTB.Mark Mod 2)<>0 and ZDGUID IN(" & strZDGUID & ") and ZRZGUID in (" & strZRZGUID & ") AND (FC_户信息属性表.HXH like '*复式1层*'  or  FC_户信息属性表.HXH not like '*复式*') AND (FC_户信息属性表.HXXXX is null or FC_户信息属性表.HXXXX<>'无分摊房产') AND (FC_户信息属性表.sygn is not null and FC_户信息属性表.sygn<>'架空层') ORDER BY cdbl(CH),DYH,Len(SHBW),SHBW"
    'addloginfo "GetHUinfo sql=" & sql
    GetMdbValues sql,strFields,fieldsCount,huinfoAr,huinfoCount
End Function

Function GetQLRinfo(ByRef qlrinfoAr,ByRef  qlrinfoCount,ByVal strFeatureGUID)
    
    strFields = "SXH, QLRMC, QLRLX, ZJZL, ZJH, DZ, DH, YB, GYFSNAME, SFCZR"
    strCondition = "GLGUID IN(" & strFeatureGUID & ") ORDER BY SXH"
    fieldsCount = 10
    sql = "select SXH, QLRMC, QLRLX, ZJZL, ZJH, DZ, DH, YB, GYFSNAME, SFCZR from QLR_权利人信息表 WHERE GLGUID IN(" & strFeatureGUID & ") ORDER BY SXH"
    'addloginfo "GetQLRinfo sql=" & sql
    GetMdbValues sql,strFields,fieldsCount,qlrinfoAr,qlrinfoCount
End Function

'计算报表页数，参数：记录总行数，首页行数，页行数，末页行数
Function CalculatePagesCount(ByRef nResPageCount, ByRef nRowRecordCount, ByRef nRowFirstPagecount, ByRef nRowPageCount, ByRef nRowEndPageCount)
    nRPTRowCount = nRowRecordCount - nRowFirstPagecount - nRowEndPageCount                '减去首末页行数的记录数，排除首、末页有非整页的情况
    nResPageCount = (1 + Int((nRPTRowCount - 1) / nRowPageCount))                                           '计算中间页数，按整页行计算的页数    
    nResPageCount = nResPageCount + 1 + 1                                                                                                '总页数，包含首、末页和中间页
End Function

'计算记录对应行号，参数：返回行号，返回页号，记录行号，起始页(表)号，表格首页行数，表格首页起始行号，中间页行数，中间页起始行号，末页行数
Function CalculatePageRowNum(ByRef nResRowNum, ByRef nResPageNum, ByRef nRowRecordNum, ByRef nStartPageNum, ByRef nRowFirstPageCount, ByRef nStartRowNumFirstPage, ByRef nRowPageCount, ByRef nStartRowNumPage, ByRef nRowEndPageCount)
    If nRowRecordNum <= nRowFirstPageCount Then                        '计算首页的页码、行号
        nResPageNum = nStartPageNum
        nResRowNum = nStartRowNumFirstPage + (nRowRecordNum - 1)
    Else                                                            '计算中间页的页码、行号
        nResPageNum = nStartPageNum + (Int((nRowRecordNum - nRowFirstPageCount - 1) / nRowPageCount) + 1)
        nResRowNum = nStartRowNumPage + (((nRowRecordNum - nRowFirstPageCount) - 1) Mod nRowPageCount)
    End If
    'MsgBox nRowRecordNum & "_" & nResPageNum & Chr(13) & nResRowNum
End Function

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



'***********************************************************数据库操作函数***********************************************************
'//strTableName 表
'//strFields 字段
'//strAddCondition 条件 
'//strTableType "AttributeData（纯属性表） ,SpatialData（地物属性表）" 
'//strGeoType 地物类型 点、线、面、注记(0点，1线，2面，3注记)
'//rs 表记录二维数组rs(行,列)
'//fieldCount 字段个数
'//返回值 ：sql查询表记录个数
Function GetProjectTableList(ByVal strTableName,ByVal strFields,ByVal strAddCondition,ByVal strOrder,ByVal strTableType,ByVal strGeoType,ByRef rs(),ByRef fieldCount)
    GetProjectTableList = 0
    values = ""
    rsCount = 0
    fieldCount = 0
    If strTableName = "" Or strFields = "" Then Exit Function
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
        If strAddCondition <> "" Then      strCondition = " (" & GeoType & ".Mark Mod 2)<>0 and " & strAddCondition & ""
        sql = "select  " & strFields & " from " & strTableName & "  INNER JOIN " & GeoType & " ON " & strTableName & ".ID = " & GeoType & ".ID WHERE " & strCondition & ""
    Else
        If strAddCondition <> "" Then
            strCondition = strAddCondition
            sql = "select  " & strFields & " from " & strTableName & "  WHERE  " & strCondition & ""
        Else
            sql = "select  " & strFields & " from " & strTableName & ""
        End If
    End If
    If    strOrder <> "" Then sql = sql & " order by " & strOrder
    '获取当前工程edb表记录
    AccessName = SSProcess.GetProjectFileName
    '判断表是否存在
    'if  IsTableExits(AccessName,strTableName)=false then exit function 
    'set adoConnection=createobject("adodb.connection")
    'strcon="DBQ="& AccessName &";DRIVER={Microsoft Access Driver (*.mdb)};"  
    'adoConnection.Open strcon
    Set adoRs = CreateObject("ADODB.recordset")
    count = 0
    'addloginfo "sql=" & sql
    adoRs.open sql  ,adoConnection,3,3
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

Function AddLoginfo(msg)
    SSProcess.MapCallBackFunction "OutputMsg", "[" & Now & "] " & msg, 1
End Function

'//初始化doc
Function initDocCom(ByVal strDocFileName)
    On Error Resume Next
    initDocCom = False
    Set docObj = CreateObject ("asposewordscom.asposewordshelper")
    If  TypeName (docObj) = "AsposeWordsHelper" Then
        initDocCom = True
        docObj.CreateDocumentByTemplate strDocFileName
    Else
        batPath = SSProcess.GetSysPathName (0) & "\Aspose.Words\regAspose.bat"
        res = RunBat (batPath)
        If res = True Then docObj.CreateDocumentByTemplate strDocFileName
        initDocCom = True
    End If
    
End Function

'//运行Bat 注册Aspose.Words插件
Function RunBat(ByVal batPath)
    On Error Resume Next
    RunBat = False
    'run函数有三个参数:
    '第一个参数是你要执行的程序的路径;
    '第二个程序是窗口的形式，0是在后台运行；1表示正常运行；2表示激活程序并且显示为最小化；3表示激活程序并且显示为最大化；一共有10个这样的参数我只列出了4个最常用的;
    '第三个参数是表示这个脚本是等待还是继续执行，如果设为了true,脚本就会等待调用的程序退出后再向后执行。其实，run做为函数，前面还有一个接受返回值的变量，一般来说如果返回为0，表示成功执行，如果不为0，则这个返回值就是错误代码，可以通过这个代码找出相应的错误。
    Set objShell = CreateObject("Wscript.Shell")
    res = objShell.Run (batPath,0,False)
    If res = 0 Then RunBat = True
    Set objShell = Nothing
End Function
Function ReleaseDB()
    adoConnection.Close
    Set adoConnection = Nothing
End Function

Function initDB
    accessName = SSProcess.GetProjectFileName
    Set adoConnection = CreateObject("adodb.connection")
    strcon = "DBQ=" & accessName & ";DRIVER={Microsoft Access Driver (*.mdb)};"
    adoConnection.Open strcon
End Function
'//关闭doc
Function ReleaseDOC(ByVal strDocFileName)
    docObj.SaveEx strDocFileName
End Function

'//填充表内单元格
Function SetCellValue(ByVal TableNum,ByVal row, ByVal col, ByVal value,ByVal mark)
    If IsNull(Value) = True  Or Value = "" Then  Value = " "
    TableNum = Int(TableNum)
    row = Int(row)
    col = Int(col)
    docObj.SetCellText TableNum,row,col,value,True
End Function

'//替换文字内容
Function ReplaceOneStr(ByVal Field, ByVal Value)
    If IsNull(Value) = True  Or Value = "" Then  Value = " "
    docObj.Replace Field,Value,0
End Function

Function  IsFolderExists(fldName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FolderExists(fldName)) = False Then fso.CreateFolder fldName
    Set fso = Nothing
End Function

Function CreateSavePath
    filePath = SSProcess.GetProjectFileName
    path = Left(filePath,InStrRev(filePath,"\"))
    IsFolderExists path
    
    spath = path & "\成果文件\不动产测量报告\"
    IsFolderExists spath
    
    CreateSavePath = spath
End Function

Function SetCellValue(ByVal TableNum,ByVal row, ByVal col, ByVal value)
    If IsNull(Value) = True  Or Value = "" Then  Value = " "
    TableNum = Int(TableNum)
    row = Int(row)
    col = Int(col)
    docObj.SetCellText TableNum,row,col,value,True
End Function

Function ReportFileStatus(filename)
    Dim fso, f, s
    Set fso = CreateObject("Scripting.FileSystemObject")
    ReportFileStatus = fso.FileExists(fileName)
End Function

#include ".\function\SQLOperateVbsFunc.vbs"
#include ".\function\DictionaryTranslateFunc.vbs"
#include ".\function\SortFindFunc.vbs"
#include ".\function\ScanString.vbs"
#include ".\function\FileFolderOperateFunc.vbs"
