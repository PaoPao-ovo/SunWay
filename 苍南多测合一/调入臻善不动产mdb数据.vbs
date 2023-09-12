Function BeforeSaveImportData()
    
    geoCount = SSProcess.GetSelGeoCount
    For i = 0 To geoCount - 1
        objType = SSProcess.GetSelGeoValue (i, "SSObj_Type")
        objLayerName = SSProcess.GetSelGeoValue (i, "SSObj_LayerName")
        strMemoData = SSProcess.GetSelGeoValue(i, "SSObj_MemoData")
        '   sysid = SSProcess.GetSelGeoValue(i, "[sysid]")
        strEpsCode = ""
        
        
        If objLayerName = "ST_ZD"  Then
            strEpsCode = "9130223"
            SSProcess.SetSelGeoValue i, "[ZDGUID]", GenerateGuid()
        ElseIf objLayerName = "ST_ZRZ"  Then
            strEpsCode = "9210123"
        ElseIf objLayerName = "ST_JZD"  Then
            strEpsCode = "9130231"
        ElseIf objLayerName = "ST_DJQ"  Then
            strEpsCode = "9130023"
        ElseIf objLayerName = "ST_DJZQ"  Then
            strEpsCode = "9130033"
        ElseIf objLayerName = "ST_XZQXJ"  Then
            strEpsCode = "9130013"
        End If
        
        
        If objLayerName = "ST_ZD"  Then
            strEpsCode = "9130223"  '使用权宗地
            SSProcess.SetSelGeoValue i, "[ZDGUID]", GenerateGuid()
        ElseIf objLayerName = "ST_DJQ"  Then
            strEpsCode = "9130023"  '地籍区
        ElseIf objLayerName = "ST_DJZQ"  Then
            strEpsCode = "9130033" '地籍子区
        ElseIf objLayerName = "ST_DSJZD"  Then
            strEpsCode = "9130231"
        ElseIf objLayerName = "ST_DSJZX"  Then
            strEpsCode = "9130242"
        ElseIf objLayerName = "ST_DSZD"  Then
            strEpsCode = "9130213"
        ElseIf objLayerName = "ST_DXJZD"  Then
            strEpsCode = "9130231" '宗地界址点
        ElseIf objLayerName = "ST_DXJZX"  Then
            strEpsCode = "9130242"
        ElseIf objLayerName = "ST_DXZD"  Then
            strEpsCode = "9130213"  '所有权宗地
        ElseIf objLayerName = "ST_GZW"  Then
            strEpsCode = "9210113" '构筑物
        ElseIf objLayerName = "ST_JD"  Then
            strEpsCode = "6601003"
        ElseIf objLayerName = "ST_JF"  Then
            strEpsCode = "6705013"
        ElseIf objLayerName = "ST_JZD"  Then
            strEpsCode = "9130231" '宗地界址点
        ElseIf objLayerName = "ST_JZDSUOYQ"  Then
            strEpsCode = "9130231" '宗地界址点
        ElseIf objLayerName = "ST_JZDZH"  Then
            strEpsCode = "9130231" '宗地界址点
        ElseIf objLayerName = "ST_JZX"  Then
            strEpsCode = "9130242"
        ElseIf objLayerName = "ST_JZXSUOYQ"  Then
            strEpsCode = "9130242"
        ElseIf objLayerName = "ST_JZXZH"  Then
            strEpsCode = "9130242"
        ElseIf objLayerName = "ST_SYQJZD"  Then
            strEpsCode = "9130231"
        ElseIf objLayerName = "ST_SYQJZX"  Then
            strEpsCode = "9130242"
        ElseIf objLayerName = "ST_SYQZD"  Then
            strEpsCode = "9130213"
        ElseIf objLayerName = "ST_SYQZDZJ"  Then
            strEpsCode = "9135027"
        ElseIf objLayerName = "ST_XZQJX"  Then
            strEpsCode = "6502012"
        ElseIf objLayerName = "ST_XZQXJ"  Then
            strEpsCode = "6501003"
        ElseIf objLayerName = "ST_XZQZJ"  Then
            strEpsCode = "9035015"
        ElseIf objLayerName = "ST_ZDNYD"  Then '未对
            strEpsCode = "9130223"
        ElseIf objLayerName = "ST_ZDSUOYQ"  Then
            strEpsCode = "9130213"  '所有权宗地
        ElseIf objLayerName = "ST_ZDZJ"  Then
            strEpsCode = "9135026"
        ElseIf objLayerName = "ST_ZRZ"  Then
            strEpsCode = "9210123" '自然幢
        ElseIf objLayerName = "YCFW_C"  Then
            strEpsCode = "9210313"
        ElseIf objLayerName = "YCFW_H"  Then
            strEpsCode = "9210513"
        ElseIf objLayerName = "YCZRZ"  Then
            strEpsCode = "9210123"
        ElseIf objLayerName = "ST_DLTB"  Then
            strEpsCode = "7320"
        Else
            FCODE = SSProcess.GetSelGeoValue (i,  "[YSBM]")
            EpsCode = ""
            CallBackFunc_FindGeoCode FCODE,EPSCode,statue
            If EpsCode <> ""  Then
                strEpsCode = EPSCODE
            End If
        End If
        
        
        If strEpsCode <> "" Then
            SSProcess.SetSelGeoValue  i,  "SSObj_ID","0"
            SSProcess.ResetSelGeoByCode i, strEpsCode
        End If
    Next
    
    
End Function

#include ".\function\Encryption.vbs"
Dim  Registrationkeyidlist,HardID,usbkeyidlist,usbkeyid
Dim NewCodes(1000000),EPSCodes(1000000),OldCodes(1000000),statues(1000000),CodeCount
Sub OnClick()
    RegistrationMode = SSProcess.ReadEpsGlobalIni("SoftRegister", "Mode" , "")
    
    'If   RegistrationMode = 1 Then
    ' RegistrationMode1
    ' If Registrationkeyidlist = Replace(Registrationkeyidlist,HardID,"") Or HardID = "0"  Then  MsgBox "软件未正常授权，请确认注册是否正确！"
    ' Exit Sub
    'ElseIf RegistrationMode = 2 Then
    ' RegistrationMode2
    ' If usbkeyidlist = Replace(usbkeyidlist,usbkeyid,"") Or usbkeyid = 0  Then    MsgBox "软件未正常授权，请确认注册是否正确！"
    'Exit Sub
    ' Else
    ' MsgBox "软件未正常授权，请确认注册是否正确！"
    'Exit Sub
    'End If
    
    mapHandle = SSProject.GetActiveMap
    mapType = SSProject.GetMapInfo(mapHandle, "MapType")
    If mapType <> 2 Then
        MsgBox "本功能只支持在地形图窗口执行！"
        Exit Sub
    End If
    If 0 Then
        SSProcess.ClearInputParameter
        SSProcess.AddInputParameter "选择调入类型", "实测",0, "预测,实测", ""
        result = SSProcess.ShowInputParameterDlg( "选择调入类型" )
        If result = 0 Then Exit Sub
        
        SSProcess.UpdateScriptDlgParameter 1
        DiaoRLX = SSProcess.GetInputParameter ("选择调入类型" )
    End If
    
    fileName = SSProcess.SelectFileName(1,"",0,"ArcGis Pdb/Fgdb MDB Files(*.mdb)|*.mdb|All Files (*.*)|*.*||")
    If fileName = ""  Then  Exit Sub
    ReadCodeTable
    importmdb fileName
    SetZDDM
    'DiaoRDX
    CREATECH fileName
    
    CreateCeng fileName
    
    CreateYCeng fileName
    
    MsgBox "ok!"
    
End Sub

Function SetZDDM
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=", "9130223"
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        polygonID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
        BDCDYH = SSProcess.GetSelGeoValue( i, "[BDCDYH]" )
        If Len(BDCDYH) > 19 Then
            GetFirstNineteenChars = Left(BDCDYH, 19)
        Else
            GetFirstNineteenChars = BDCDYH
        End If
        SSProcess.SetObjectAttr polygonID, "[ZDDM]", GetFirstNineteenChars
    Next
End Function


Function DELXSD(SYSID)
    WZ = InStr(SYSID,".")
    CHANGDU = Len(SYSID)
    If WZ <> 0 Then
        XSDWS = CHANGDU - WZ
        SYSID = Left(SYSID,XSDWS)
    End If
End Function
Function importmdb(filename)
    '清空转换参数
    SSProcess.ClearDataXParameter
    '设置导入文件格式为ArcGIS PDB
    SSProcess.SetDataXParameter "DataType", "22"
    SSProcess.SetDataXParameter "SaveAttrToMemoData", "0"
    SSProcess.SetDataXParameter "ImportPathName",filename
    SSProcess.SetDataXParameter "ImportNotMatchAttrToMemoData", "1"
    
    
    startIndex = 0
    SSProcess.SetDataXParameter "ExportLayerCount","54"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_ZD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_JZD"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_JZDX"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_DJQ"
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_DJZQ"
    
    
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_DXJZD" '地下界址点
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_DXJZX" '地下界址线
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_DXZD" '地下宗地
    
    'if DiaoRLX="实测" then   
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"ST_ZRZ"
    '    else
    SSProcess.SetDataXParameter "ExportLayer" & CStr(AddOne(startIndex)),"YCZRZ"
    '    end if
    '属性对照
    '关联表输出对照个数//未在属性表中挂接的属性表
    SSProcess.SetDataXParameter "RelateTableRelationCount", "500"
    startIndex = 1
    'EPS表名,SDE表名,SDE表别名
    'SSProcess.SetDataXParameter "RelateTableRelation"&CStr(AddOne(startIndex)), "ZD_宗地基本信息属性表,T_ZD,土地所有权"
    SSProcess.SetDataXParameter "RelateTableRelation" & CStr(AddOne(startIndex)), "QLR_权利人信息表,FW_QLR,权利人信息"
    SSProcess.SetDataXParameter "RelateTableRelation" & CStr(AddOne(startIndex)), "QLR_权利人信息表,ZD_SHIYQQLR,权利人信息"
    startIndex = 10
    SSProcess.SetDataXParameter "TableFieldDefCount","20000"
    
    'ZDDM    BDCDYH    SJLX    SYQLX    ZDTZM    QLLX    QLXZ    QLSDFS    PZMJ    ZDMJ    MJDW    ZL    ZDSZD    ZDSZN    ZDSZX    ZDSZB    PZYT    TDYT    DJ    JG    RJL    JZMD    JZXG    JZZDMJ    JZMJ    ZDT    TFH    DJH    JZDWSM    JZXZXSM    DCJS    DCR    DCRQ    CLJS    CLR    CLRQ    SHYJ    SHR    SHRQ    BZ    ZT
    '层名,类型(0点,1线,2面,3注记,10点线面共用),EPS字段名,客户字段名,[客户字段别名,]系统字段名,缺省值,字段类型,字段长度,小数位
    '使用权宗地
    
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,CHXMBH,CHXMBH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,DJH,DJH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,DJQDM,DJQDM,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,DJQZH,DJQZH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,HQ_RQ,HQ_RQ,,,dbDate,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,LGID,LGID,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,QLXZNAME,QSXZ,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,SJ_YT,SJ_YT,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,state,state,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,sysid,sysid,,,dbtext,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,SZ_JF,SZ_JF,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,ZL,TDZL,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,YT,XSJ_YT,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,ZD_FH,ZD_FH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,BSM,ZD_ID,,,dbText,19,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,TDSYQXZ,ZD_SYQLX,,,dbText,19,0"
    ' SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,YSDM,YS_DM,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,ZDTZM,ZD_TZM,,,dbText,19,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,ZD_ZH,ZD_ZH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,ZDMJ,ZDMJ,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,ZDSZ,ZDSZ,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"使用权宗地,2,ZDTYBM,ZDTYBM,,,dbText,255,0"
    '自然幢
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,JianSDW,KFDW,开发单位,,,dbText,50,0"  '==============CeLDW
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,CeLDW,CHDW,测绘单位,,,dbText,50,0"  '==============CeLDW
    'SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,GHYT,FWYT,规划用途,,,dbText,2,0"  '==============
    'SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,FWJG,FWLX,房屋结构,,,dbText,2,0" '==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,ZT,JZWZT,状态,,,dbText,2,0" '============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,YCDXJZMJ,YCDXMJ,预测地下面积,,,dbDouble,16,6"  '==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,YCQTJZMJ,YCQTMJ,预测其它面积,,,dbDouble,16,6"  '==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,SCDXJZMJ,SCDXMJ,实测地下面积,,,dbDouble,16,6"  '==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,SCQTJZMJ,SCQTMJ,实测其它面积,,,dbDouble,16,6"  '==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,TNJZMJ,TNMJ,套内建筑面积,,,dbDouble,16,6"'==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,GYJZMJ,GYMJ,共有建筑面积,,,dbDouble,16,6"'==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,DiaoCR,DCR,DCR,,,dbText,200,0"'==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,ShenHR,SHR,SHR,,,dbText,200,0"'==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,ZhiTY,HTY,HTY,,,dbText,200,0"'==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,GuiHXKZBH,GHXKZH,GHXKZH,,,dbText,100,0"'==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,DCJSJYJ,DCYJ,DCYJ,,,dbText,255,0"'==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,ShenHRQ,SHRQ,SHRQ,,,dbDATE,255,0"'==============
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"自然幢,2,CSHYJ,SHYJ,SHYJ,,,dbTEXT,100,0"'==============
    
    '界址点
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,BSM,JZD_ID,,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,SZ_JF,SZ_JF,,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,HYP_JD,HYP_JD,,,,dbLONG,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,HQ_RQ,HQ_RQ,,,,dbdate,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,ISNEW,ISNEW,,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,DJZQDM,DJQDM,,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,CHXMBH,CHXMBH,,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,LGID,LGID,,,,dbText,255,0"
    'SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,SYSID,SYSID,,,,dblong,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"宗地界址点,0,state,state,,,,dbText,255,0"
    
    '权利人信息待加
    '==================================================================================表名,EPS字段名,客户字段名,[客户字段别名,]系统字段名,缺省值,字段类型,字段长度,小数位"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,BDCDYH,BDCDYH,,,dbText,28,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,BDCQFJ,BDCQFJ,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,BDCQZH,BDCQZH,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,BSM,BSM,,,dbLong,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,BZ,BZ,,,dbText,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,CHXMBH,CHXMBH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DH,DH,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DJSJ,DJSJ,,,dbDate,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DLRDZYX,DLRDZYX,,,dbText,30,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DLRGDDH,DLRGDDH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DLRJGMC,DLRJGMC,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DLRTXDZ,DLRTXDZ,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DLRXB,DLRXB,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DaiLRXM,DLRXM,,,dbText,200,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DLRYB,DLRYB,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DaiLRDH,DLRYDDH,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DLRYYZZHM,DLRYYZZHM,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DaiLRZJH,DLRZJH,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DaiLRZJZL,DLRZJZL,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DLRZYZGZH,DLRZYZGZH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DZ,DZ,,,dbText,200,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,DZYJ,DZYJ,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,FRDZYX,FRDZYX,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,FRGDDH,FRGDDH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,FaRDB,FRXM,,,dbText,200,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,FaRDBDH,FRYDDH,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,FaRDBZJZL,FRZJZL,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,FRZW,FRZW,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,FZJG,FZJG,,,dbText,200,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,GJ,GJ,,,dbText,6,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,GYFS,GYFS,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,GYQK,GYQK,,,dbText,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,GZDW,GZDW,,,dbText,100,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,HJSZSS,HJSZSS,,,dbText,6,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,JYJG,JYJG,,,dbDouble,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,LXRDZYX,LXRDZYX,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,LXRGDDH,LXRGDDH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,LXRTXDZ,LXRTXDZ,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,LXRXM,LXRXM,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,LXRYB,LXRYB,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,LXRYDDH,LXRYDDH,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,MJDW,MJDW,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,QLBL,QLBL,,,dbText,100,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,QLLX,QLLX,,,dbText,5,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,QLMJ,QLMJ,,,dbDouble,0,2"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,QLRLX,QLRLX,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,QLRMC,QLRMC,,,dbText,100,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,QXDM,QXDM,,,dbLong,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,SSHY,SSHY,,,dbText,6,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,state,state,,,dbLong,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,SXH,SXH,,,dbLong,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,sysid,sysid,,,dbLong,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,XB,XB,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,YB,YB,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,YSDM,YSDM,,,dbText,10,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,YXTBM,YXTBM,,,dbText,255,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,YXTBSM,YXTBSM,,,dbLong,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,ZDBSM,ZDBSM,,,dbLong,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,ZJH,ZJH,,,dbText,50,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,ZJZL,ZJZL,,,dbText,2,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,ZRZBSM,ZRZBSM,,,dbLong,0,0"
    SSProcess.SetDataXParameter "TableFieldDef" & CStr(AddOne(startIndex)),"QLR_权利人信息表,ZSGBH,ZSGBH,,,dbText,255,0"
    
    SSProcess.ImportData
    
    
    Dim arRecordZD(),nRecordZDCount
    Dim arRecordJZD(),nRecordJZDCount
    
    
    ZDMDBATTR = "BDCDJBSM,BDCDYH,CC_MJ,CC_TXMJ,CHXMBH,CLDW_ID,CODE,CR_ZZRQ,DC_RQ,DCBH,DCDWMC,DCJS,DCR,DJ,DJH,DJQDM,DJQZH,DKBM,FHWH,GDHT,GDHTH,HC_YJ,HQ_RQ,HTLHL,HTY,HYFLDM,JXZXSM,JZ_MD,JZ_MJ,JZ_RJL,JZ_XG,JZDWSM,JZJD_MJ,JZSMB,KZ_RQ,KZJS,KZR,LGID,LHL,MJDW,PZ_YT,PZMJ,PZYT_MC,QDJG,QLLX,QLSDFS,QLXZ,QS_XZ,QSLYZM,QXDM,RKSJ,SH_RQ,SHR,SHYJ,SJ_BZ,SJ_YT,SJYT_MC,state,SYQ_LX,SYQJSSJ2,SYQQSSJ,SYQX,sysid,SZ_B,SZ_D,SZ_JF,SZ_N,SZ_X"
    ZDATTR = "BDCDJBSM,BDCDYH,CC_MJ,CC_TXMJ,CHXMBH,CLDW_ID,CODE,ZZSJ,DiaoCRQ,DCBH,DCDWMC,QSDCJS,DiaoCR,DJ,DJH,DJQDM,DJQZH,DKBM,FHWH,GDHT,GDHTH,HC_YJ,HQ_RQ,GuiHSPLHL,ZhiTY,GMJJHYFLDM,ZYQSJXZXSM,JunGCLJZMD,JunGCLZJZMJ,JunGCLRJL,JZXG,JZDWSM,JunGCLZDMJ,JZSMB,CeLRQ,DJCLJS,CeLY,LGID,JunGCLLHL,MJDW,PZ_YT,PZMJ,PZYTMC,JG,QLLX,QLSDFS,QLXZ,QS_XZ,TDQSLY,QXDM,RKSJ,ShenHRQ,ShenHR,DJDCJGSHYJ,BZ,SJ_YT,YTNAME,State,SYQ_LX,SYQJSSJ2,QDSJ,ShiYQX,sysid,ZDSZB,ZDSZD,SZ_JF,ZDSZN,ZDSZX"
    
    JZDMDBATTR = "BZ,CHXMBH,DJQDM,HH,HQRQ,JBLX,JZ_XH,JZD_ID,JZD_X,JZD_Y,JZDH,JZDLX,JZJJ,LGID,RKSJ,SZ_JF,XGJZDBSM,YSDM,ZD_ID,ZDJZD_ID,ZDJZXBSM"
    JZDATTR = "BZ,CHXMBH,DJZQDM,HH,HQRQ,JBLX,SXH,BSM,JZD_X,JZD_Y,JZDH,JZDLX,JZJJ,LGID,RKSJ,SZ_JF,XGJZDBSM,YSDM,ZDBSM,ZDJZD_ID,ZDJZXBSM"
    
    '1.-------------打开数据库-----------------------
    SSProcess.OpenAccessMdb filename
    
    ' 宗地信息处理
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=", "9130223"
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        polygonID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
        featureguid = SSProcess.GetSelGeoValue( i, "[FeatureGUID]" )
        ZDGUID = SSProcess.GetSelGeoValue( i, "[ZDGUID]" )
        YDHXGUID = SSProcess.GetSelGeoValue( i, "[YDHXGUID]" )
        SSProcess.SetObjectAttr polygonID,"[YSDM]","DJ1213"
        If ZDGUID = "" Or ZDGUID = "{00000000-0000-0000-0000-000000000000}"  Then  SSProcess.SetObjectAttr polygonID,"[ZDGUID]",featureguid
        If YDHXGUID = "" Or YDHXGUID = "{00000000-0000-0000-0000-000000000000}"  Then  SSProcess.SetObjectAttr polygonID,"[YDHXGUID]",featureguid
        ZDID = CDbl( SSProcess.GetObjectAttr (polygonID, "[BSM]"))
        TDSYQXZ = SSProcess.GetObjectAttr (polygonID, "[TDSYQXZ]")
        If TDSYQXZ = "G"  Then
            TDSYQXZ = "1"
        ElseIf TDSYQXZ = "J"  Then
            TDSYQXZ = "2"
        End If
        
        SSProcess.SetObjectAttr polygonID, "[TDSYQXZ]", TDSYQXZ
        strSQL = "SELECT " & ZDMDBATTR & " FROM T_ZD WHERE ZD_ID =" & ZDID & ""
        
        GetSQLRecordAll filename,strSQL,arRecordZD,nRecordZDCount
        SSProcess.AccessIsEOF mdbName, sql
        
        If nRecordZDCount = 1  Then
            arRecordZD(0) = Replace(arRecordZD(0), "*", "")
            arTempZongD = Split( arRecordZD(0), ",")
            arZDATTR = Split(ZDATTR, ",")
            For j = 1 To UBound(arZDATTR)
                If arTempZongD(j) <> "" And arTempZongD(j) <> "*" Then
                    If arZDATTR(j) = "SysID" Then  DELXSD   arTempZongD(j)
                    SSProcess.SetObjectAttr polygonID, "[" & arZDATTR(j) & "]", arTempZongD(j)
                End If
            Next
        End If
    Next
    
    ' 界址点信息处理
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=", "9130231"
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        polygonID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
        JZDH = CDbl( SSProcess.GetObjectAttr (polygonID, "[BSM]"))
        
        strSQL = "SELECT " & JZDMDBATTR & " FROM T_ZDJZD WHERE JZD_ID =" & JZDH & ""
        
        GetSQLRecordAll filename,strSQL,arRecordJZD,nRecordJZDCount
        If nRecordJZDCount >= 1  Then
            arRecordJZD(0) = Replace(arRecordJZD(0), "*", "")
            arTempJZD = Split( arRecordJZD(0), ",")
            arJZDATTR = Split( JZDATTR, ",")
            For j = 1 To UBound(arJZDATTR)
                'jzdsx = SSProcess.GetObjectAttr polygonID, "[" & arJZDATTR(j) & "]"
                If arTempJZD(j) <> "" And arTempJZD(j) <> "*" Then
                    SSProcess.SetObjectAttr polygonID, "[" & arJZDATTR(j) & "]", arTempJZD(j)
                End If
            Next
        Else
            '待添加 多条记录属性加逗号
            
        End If
    Next
    
    ' 自然幢信息处理
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "=", "9210123"
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    For i = 0 To geoCount - 1
        CHZT = ""
        polygonID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
        CHLB = SSProcess.GetObjectAttr (polygonID, "[CHLB]")
        If CHLB = "预测绘"  Then CHZT = "1"
        If CHLB = "实测绘"  Then CHZT = "2"
        SSProcess.SetObjectAttr polygonID, "[CHZT]", CHZT
    Next
    
    
    '关闭数据库
    SSProcess.CloseAccessMdb filename
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

Function AddOne( ByRef startIndex )
    startIndex = startIndex + 1
    AddOne = startIndex
End Function

'----------------------- 生成GUID（正确函数）-------------------------------------------
Function GenerateGuid()
    Dim TypeLib
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GenerateGuid = Left(TypeLib.Guid ,38)
End Function

Function DiaoRDX
    '地物调入
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_LayerName", "==", "ST_DWD,ST_DMD,ST_DMX,ST_DWX,ST_DMM,ST_FWDWM,ST_LHDWM,ST_STDWM,ST_DLDWM,ST_OTDWM"
    SSProcess.SelectFilter
    count = SSProcess.GetSelgeoCount
    For i = 0 To count - 1
        FCODE = SSProcess.GetSelGeoValue (i,  "[YSBM]")
        lanyername = SSProcess.GetSelGeoValue (i, "SSObj_Layername")
        lanyername = Left(lanyername,6)
        ' sxtq I,YSDM,FWJG,FWCS,JZMJ,ZDMJ,HQRQ,KZDDH,KZDZT,KZDDJ,BSGC,KD,FHMC
        EpsCode = ""
        CallBackFunc_FindGeoCode FCODE,EPSCode,statue
        
        If EpsCode <> ""  Then
            'MSGBOX  EpsCode
            SSObj_ID = SSProcess.GetSelGeoValue( i, "SSObj_ID")
            SSProcess.SetSelGeoValue  i, "SSObj_ID", "0"
            'sxfz I,YSDM,FWJG,FWCS,JZMJ,ZDMJ,HQRQ,KZDDH,KZDZT,KZDDJ,BSGC,KD,FHMC
            SSProcess.ResetSelGeoByCode i, EpsCode
            SSProcess.SetSelGeoValue  i, "SSObj_ID", SSObj_ID
        End If
    Next
    '注记调入
    If 0 Then
        SSProcess.ClearSelection
        SSProcess.ClearSelectCondition
        SSProcess.SetSelectCondition "SSObj_LayerName", "==", "ST_DGXZJ,ST_DMZJ,ST_JFFWZJ,ST_GCDZJ,ST_JFJTZJ,ST_JFSXZJ,ST_JFZBZJ,ST_JFQTZJ"
        noteCount = SSProcess.GetSelNoteCount
        For i = 0 To noteCount - 1
            
            'Pointcount= SSProcess.GetSelNotePointcount(i)
            ' SSProcess.LockSelGeoPoint i,1
            'For j=0 to Pointcount-1
            '    SSProcess.GetSelNotePoint i, j, x, y, z, ptype, pname
            'if EllipseName="CGCS-2000椭球" then 说明：需要转坐标这句不用放开
            '         SSProcess.LongiLatiToxyCGCS2000 120, y, x, x0, y0 
            
            '    SSProcess.SetSelNotePoint i, j, x0, y0, z, ptype, pname
            '     next
            'SSProcess.UpdateSelGeoPoint i
            
            FCODE1 = SSProcess.GetSelNoteValue (i, "[YSBM]")
            MDBFontClass = ""
            'condition = "NoteTemplate.ByName = '" & FCODE &"'"
            'fontclass=SSProcess.FindNoteClass ("NoteTemplateTB_500", condition) 
            CallBackFunc_FindFontClass  FCODE1,MDBFontClass
            If MDBFontClass <> "" Then
                SSObj_ID = SSProcess.GetSelNoteValue( i, "SSObj_ID")
                SSProcess.SetSelNoteValue  i, "SSObj_ID", "0"
                SSProcess.ResetSelNoteByFontClass i, MDBFontClass
                SSProcess.SetSelNoteValue  i, "SSObj_ID", SSObj_ID
            End If
        Next
    End If
End Function


Function ReadCodeTable()
    Dim fso, ts, chLine
    Dim strs(10000),strs0(10000),count
    IsNote = 0
    fileName = SSProcess.GetSysPathName (0) & "\权调库对照.txt"
    If fileName = "" Then
        Exit Function
    End If
    If IsExistentFile( fileName ) = 0  Then
        MsgBox  "工作台面目录下没有 " & fileName & "\权调库对照"
        Exit Function
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filename, 1)
    Do While Not ts.AtEndOfStream
        chLine = ts.ReadLine & NewLine
        chLine = Trim(chLine)
        If chLine <> "" Then
            count = 0
            If  Left(chLine, 2) <> "//" Then
                'if  chLine = "NoteBegain"  then   IsNote = 1
                'if IsNote = 0 then '地物
                'SSFunc.ScanString chLine, "//", strs0, count0
                SSFunc.ScanString chLine, ",", strs, count
                If count = 2 Then '地物
                    NewCodes(CodeCount) = Trim(strs(0))   '临海编码
                    OldCodes(CodeCount) = Trim(strs(1))   'EPS编码
                    'MSGBOX  NewCodes(CodeCount)&"||"& OldCodes(CodeCount)
                    CodeCount = CodeCount + 1
                End If
                'if count=3 then'注记
                '    NewFontClasses(FontClassCount)=Trim(strs(0))   '临海注记编码
                '    OldFontClasses(FontClassCount)=Trim(strs(1))   'eps注记编码 
                '    FontClassCount=FontClassCount+1
                '    end if
            End If
        End If
    Loop
    ts.Close
End Function
'查找对应的eps地物编码
Function CallBackFunc_FindGeoCode( FCODE, ByRef EPSCode, ByRef statue)
    EPSCode = ""
    statue = ""
    For i = 0 To CodeCount - 1
        
        If   FCODE = NewCodes(i)  Then
            EPSCode = OldCodes(i)
            statue = statues(i)
            'MSGBOX  EPSCode
            Exit For
        End If
    Next
End Function
'查找对应EPS注记分类号
Function CallBackFunc_FindFontClass(FCODE1,ByRef MDBFontClass )
    MDBFontClass = ""
    For i = 0 To FontClassCount - 1
        If  NewFontClasses(i) = FCODE1 Then
            MDBFontClass = OldFontClasses(i)
            Exit For
        End If
    Next
End Function
Function IsExistentFile( fileName )
    Dim fso, f, s
    Set fso = CreateObject("Scripting.FileSystemObject")
    IsExistentFile = fso.FileExists(fileName)
End Function


Function CREATECH(filename)
    SSProcess.OpenAccessMdb filename
    '获取层对应的自然幢BSM
    strSQL = "SELECT CHXMBH,QXDM,ZRZH,CJZMJ,CFTJZMJ,SJC,CBQMJ,BZ,MYC,CG,ZDBSM,SPTYMJ,BSM,YXTBM,CTNJZMJ,YXTBSM,ZRZBSM,GXSJ,CYTMJ,CH,CGYJZMJ,SysID,State FROM FW_C WHERE OBJECTID IS NOT NULL"
    GetSQLRecordAll filename,strSQL,arRecordZD,nRecordZDCount
    If  nRecordZDCount > 0 Then
        'MSGBOX  nRecordZDCount
        For I = 0 To  nRecordZDCount - 1
            
            '获取自然幢的中心点，创建层
            CJL = Split (arRecordZD(I),",")
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_Code", "=", "9210123"
            SSProcess.SetSelectCondition "[BSM]", "==", CJL(16)
            SSProcess.SelectFilter
            geoCount = SSProcess.GetSelGeoCount()
            If GEOCOUNT = 1 Then
                
                id = SSProcess.GetSelGeoValue(0, "SSObj_ID")
                SSProcess.GetObjectFocusPoint ID, x, y
                makeAreaC x,y,arRecordZD(I)
            Else
                'qsbsm= qsbsm&"," &CJL(16)
            End If
        Next
        'msgbox  qsbsm
    End If
    '创建户
    
    MDBHZD = "BDCDJBSM,BDCDYH,BQTGS,BSM,BZ,CBSM,CG,CHXMBH,CQLY,DCFJSM,DCR,DCRQ,DCYJ,DJZT,DQTGS,DYH,DYTDMJ,DZWSXH,DZWTZM,FCFHT,FJSXH,FTTDMJ,FW_DH,FW_LDH,FWBM,FWCB,FWJG,FWLX,FWQJXSYT,FWXZ,FWYT,GHYT,GNQBH,GROUPINDEX,GYTDMJ,HH,HX,HXJG,JC,JGSJ,JHQHH,JSXMBSM,LJZH,MPH,MYC,NQTGS,PZJZMJ,QHH,QXDM,RKSJ,SCDXBFJZMJ,SCFTJZMJ,SCFTXS,SCJZMJ,SCQTJZMJ,SCTNJZMJ,SFFF,SHBW,SJC,SJCS,state,sysid,TDQLXZ,TDSYJSSJ,TDSYQMJ,TDSYQSSJ,TDYT,TDYTMC,XMMC,XQMC,XQTGS,YCDXBFJZMJ,YCFTJZMJ,YCFTXS,YCJZMJ,YCQTJZMJ,YCTNJZMJ,YFWJG,YFWLX,YFWXZ,YFWYT,YFWYTMC,YGHYT,YSDM,YXTBH,YXTBM,YXTBSM,YZBM,ZCS,ZDBSM,ZFBDCDYH,ZFBSM,ZFFWBM,ZL,ZRZBSM,ZRZH,ZRZSXH,ZSCJZMJ,ZT,ZYCJZMJ"
    HSQL = "SELECT BDCDJBSM,BDCDYH,BQTGS,BSM,BZ,CBSM,CG,CHXMBH,CQLY,DCFJSM,DCR,DCRQ,DCYJ,DJZT,DQTGS,DYH,DYTDMJ,DZWSXH,DZWTZM,FCFHT,FJSXH,FTTDMJ,FW_DH,FW_LDH,FWBM,FWCB,FWJG,FWLX,FWQJXSYT,FWXZ,FWYT,GHYT,GNQBH,GROUPINDEX,GYTDMJ,HH,HX,HXJG,JC,JGSJ,JHQHH,JSXMBSM,LJZH,MPH,MYC,NQTGS,PZJZMJ,QHH,QXDM,RKSJ,SCDXBFJZMJ,SCFTJZMJ,SCFTXS,SCJZMJ,SCQTJZMJ,SCTNJZMJ,SFFF,SHBW,SJC,SJCS,state,sysid,TDQLXZ,TDSYJSSJ,TDSYQMJ,TDSYQSSJ,TDYT,TDYTMC,XMMC,XQMC,XQTGS,YCDXBFJZMJ,YCFTJZMJ,YCFTXS,YCJZMJ,YCQTJZMJ,YCTNJZMJ,YFWJG,YFWLX,YFWXZ,YFWYT,YFWYTMC,YGHYT,YSDM,YXTBH,YXTBM,YXTBSM,YZBM,ZCS,ZDBSM,ZFBDCDYH,ZFBSM,ZFFWBM,ZL,ZRZBSM,ZRZH,ZRZSXH,ZSCJZMJ,ZT,ZYCJZMJ FROM FW_H WHERE OBJECTID IS NOT NULL"
    GetSQLRecordAll filename,HSQL,HRecord,RecordHCount
    
    If  RecordHCount > 0 Then
        'MSGBOX  RecordHCount
        For I = 0 To  RecordHCount - 1
            '获取自然幢的中心点，创建户
            HJL = Split (HRecord(I),",")
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_Code", "=", "9210123"
            SSProcess.SetSelectCondition "[BSM]", "==", HJL(94)
            SSProcess.SelectFilter
            geoCount = SSProcess.GetSelGeoCount()
            'MSGBOX GEOCOUNT
            If GEOCOUNT = 1 Then
                id = SSProcess.GetSelGeoValue(0, "SSObj_ID")
                SSProcess.GetObjectFocusPoint ID, x, y
                makeAreaH x,y,HRecord(I)
                HTZ HJL(3),HRecord(I)
            Else
                qsbsm = Replace(qsbsm,HJL(94),"")
                qsbsm = qsbsm & "," & HJL(94)
            End If
        Next
    End If
    
    
    SSProcess.CloseAccessMdb filename
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


Function makeAreaC(x1,y1,sxzd)
    zdmcsz = "CHXMBH,QXDM,ZRZH,CJZMJ,CFTJZMJ,SJC,CBQMJ,BZ,MYC,CG,ZDBSM,SPTYMJ,BSM,YXTBM,CTNJZMJ,YXTBSM,ZRZBSM,GXSJ,CYTMJ,CH,CGYJZMJ,SysID,State"
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", 9210313
    'SSProcess.SetNewObjValue "SSObj_Color", color
    ' SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "楼层"
    sxz = Split(sxzd,",")
    zdmc = Split(zdmcsz,",")
    For i = 0 To UBound(zdmc)
        SSProcess.SetNewObjValue "[" & zdmc(i) & "]", sxz(i)
    Next
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function makeAreaH(x1,y1,sxzd)
    EPSHZD = "BDCDJBSM,BDCDYH,BQ,BSM,BZ,CBSM,CG,CHXMBH,CQLY,DCFJSM,DCR,DCRQ,DCYJ,DJZT,DQ,DYH,DYTDMJ,DZWSXH,DZWTZM,FWFHT,HXH,FTTDMJ,FW_DH,FW_LDH,FWBM,ChanB,FWJG,FWLX,FWQJXSYT,FWXZ,FWYT1,GHYT,GNQBH,GROUPINDEX,GYTDMJ,HH,HX,HXJG,JC,JGSJ,JHQHH,JSXMBSM,LJZH,ZBMPH,MYC,NQ,PZJZMJ,QHH,QXDM,RKSJ,SCDXBFJZMJ,SCFTJZMJ,SCFTXS,SCJZMJ,SCQTJZMJ,SCTNJZMJ,SFFF,SHBW,SJC,SJCS,State,SysID,QLXZNAME,TDSYZZRQ,TDSYQMJ,TDSYKSRQ,TDYT,TDYTNAME,XMMC,XQMC,XQ,YCDXBFJZMJ,YCFTJZMJ,YCFTXS,YCJZMJ,YCQTJZMJ,YCTNJZMJ,YFWJG,YFWLX,YFWXZ,YFWYT,YFWYTMC,YGHYT,YSDM,YXTBH,YXTBM,YXTBSM,YZBM,ZCS,ZDBSM,ZFBDCDYH,ZFBSM,ZFBSM,ZL,ZRZBSM,ZRZH,ZRZSXH,ZSCJZMJ,ZT,ZYCJZMJ"
    'EPSHZD="BDCDJBSM,BDCDYH,BGQBSM,BQ,BSM,BZ,CBSM,CHXMBH,CQLY,DCFJSM,DCR,DCRQ,DCYJ,DJZT,DQ,DYH,DYTDMJ,DZWSXH,DZWTZM,FWFHT,FCPMTBSM,HXH,FTTDMJ,FW_DH,FW_LDH,FW_STAT,FWBM,ChanB,FWJG,FWLX,FWXZ,FWYT1,GHYT,GNQBH,GROUPINDEX,GYTDMJ,HH,HX,HXJG,JC,JGSJ,JHQHH,JSXMBSM,LJZBSM,LJZH,ZBMPH,MYC,NQ,QHH,QS_FLOOR,QXDM,RKSJ,SCCHXMBH,SCDXBFJZMJ,SCFTJZMJ,SCFTXS,SCJZMJ,SCQTJZMJ,SCSJ,SCTNJZMJ,SFFF,SHBW,SJC,SJCS,SJZH,State,SysID,QLXZNAME,TDSYZZRQ,TDSYQMJ,TDSYKSRQ,TDYT,TDYTNAME,XMMC,XQMC,XQ,YCDXBFJZMJ,YCFTJZMJ,YCFTXS,YCJZMJ,YCQTJZMJ,YCTNJZMJ,YFWJG,YFWLX,YFWXZ,YFWYT,YFWYTMC,YGHYT,YSDM,YXTBH,YXTBM,YXTBSM,YZBM,ZCS,ZDBSM,ZDDM,ZDQLR_ID,ZFBDCBH,ZFBDCDYH,ZFBSM,ZFBSM,ZL,ZRZBSM,ZRZH,ZRZSXH,ZSCJZMJ,ZT,ZYCJZMJ"
    
    SSProcess.CreateNewObj 2
    SSProcess.SetNewObjValue "SSObj_Code", 9210513
    'SSProcess.SetNewObjValue "SSObj_Color", color
    ' SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "户"
    sxz = Split(sxzd,",")
    '    zdmc=split(EPSHZD,",")
    'for i=0 to ubound(zdmc)    
    'SSProcess.SetNewObjValue "["&zdmc(i)&"]", sxz(i)
    
    'next
    SSProcess.SetNewObjValue "[BSM]", sxz(3)
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjPoint x1, y1, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function HTZ(BSM,sxzd)
    
    EPSHZD = "BDCDJBSM,BDCDYH,BQ,BSM,BZ,CBSM,CG,CHXMBH,CQLY,DCFJSM,DCR,DCRQ,DCYJ,DJZT,DQ,DYH,DYTDMJ,DZWSXH,DZWTZM,FWFHT,HXH,FTTDMJ,FW_DH,FW_LDH,FWBM,ChanB,FWJG,FWLX,FWQJXSYT,FWXZ,FWYT1,GHYT,GNQBH,GROUPINDEX,GYTDMJ,HH,HX,HXJG,JC,JGSJ,JHQHH,JSXMBSM,LJZH,ZBMPH,MYC,NQ,PZJZMJ,QHH,QXDM,RKSJ,SCDXBFJZMJ,SCFTJZMJ,SCFTXS,SCJZMJ,SCQTJZMJ,SCTNJZMJ,SFFF,SHBW,SJC,SJCS,State,SysID,QLXZNAME,TDSYZZRQ,TDSYQMJ,TDSYKSRQ,TDYT,TDYTNAME,XMMC,XQMC,XQ,YCDXBFJZMJ,YCFTJZMJ,YCFTXS,YCJZMJ,YCQTJZMJ,YCTNJZMJ,YFWJG,YFWLX,YFWXZ,YFWYT,YFWYTMC,YGHYT,YSDM,YXTBH,YXTBM,YXTBSM,YZBM,ZCS,ZDBSM,ZFBDCDYH,ZFBSM,ZFBSM,ZL,ZRZBSM,ZRZH,ZRZSXH,ZSCJZMJ,ZT,ZYCJZMJ"
    sxz = Split(sxzd,",")
    zdmc = Split(EPSHZD,",")
    For i = 0 To UBound(zdmc)
        If i = 0 Then
            zfc = zdmc(i) & "=" & sxz(i)
        Else
            If  zdmc(i) <> "BSM" Then
                zfc = zfc & "," & zdmc(i) & "=" & sxz(i)
            End If
        End If
    Next
    mdbName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb mdbName
    sql = "update  FC_户信息属性表 set " & ZFC & " where BSM = " & BSM
    SSProcess.ExecuteAccessSql  mdbName,sql
    SSProcess.CloseAccessMdb mdbName
    SSProcess.MapMethod "clearattrbuffer",  "FC_户信息属性表"
End Function

'创建层
Function CreateCeng(ByVal FileName)
    
    SSProcess.OpenAccessMdb FileName
    
    SqlStr = "Select ZDBSM,ZRZBSM,ZRZH,SJC,MYC,QXDM From FW_H Where OBJECTID IS NOT NULL "
    GetSQLRecordAll EdbName,SqlStr,HInfoArr,HCount
    
    FildsStr = "ZDBSM,ZRZBSM,ZRZH,SJC,MYC,QXDM"
    FildArr = Split(FildsStr,",", - 1,1)
    
    If HCount > 0 Then
        SqlStr = "Select ZDBSM,ZRZBSM,ZRZH,SJC,MYC,QXDM From FW_C Where OBJECTID IS NOT NULL "
        GetSQLRecordAll EdbName,SqlStr,CInfoArr,CCount
        
        If CCount > 0 Then
            For i = 0 To CCount - 1
                ValArr = Split(CInfoArr(i),",", - 1,1)
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", "9210123"
                SSProcess.SetSelectCondition "[BSM]", "==", ValArr(1)
                SSProcess.SelectFilter
                geoCount = SSProcess.GetSelGeoCount()
                
                If geoCount > 0 Then
                    SSProcess.CreateNewObj 2
                    SSProcess.SetNewObjValue "SSObj_Code",9210313
                    SSProcess.SetNewObjValue "SSObj_LayerName", "楼层"
                    SSProcess.SetNewObjValue "[BZ]", "新增"
                    SSProcess.SetNewObjValue "[CHZT]", 2
                    For j = 0 To UBound(FildArr)
                        SSProcess.SetNewObjValue "[" & FildArr(j) & "]" , ValArr(j)
                    Next 'j
                    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
                    SSProcess.GetObjectFocusPoint ID, x, y
                    SSProcess.AddNewObjPoint x, y, 0, 0, ""
                    SSProcess.AddNewObjPoint x, y, 0, 0, ""
                    SSProcess.AddNewObjToSaveObjList
                    SSProcess.SaveBufferObjToDatabase
                End If
            Next 'i
            
            For i = 0 To CCount - 1
                ValArr = Split(CInfoArr(i),",", - 1,1)
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", 9210313
                SSProcess.SetSelectCondition "[ZRZBSM]", "==", ValArr(1)
                SSProcess.SetSelectCondition "[SJC]", "==", ValArr(3)
                SSProcess.SetSelectCondition "[CHZT]", "==", 2
                SSProcess.SelectFilter
                geoCount = SSProcess.GetSelGeoCount()
                For j = 0 To geoCount - 2
                    SSProcess.DelSelGeo j
                Next 'j
            Next 'i
            
        Else
            For i = 0 To HCount - 1
                ValHArr = Split(HInfoArr(i),",", - 1,1)
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", "9210123"
                SSProcess.SetSelectCondition "[BSM]", "==", ValHArr(1)
                SSProcess.SelectFilter
                geoCount = SSProcess.GetSelGeoCount()
                If geoCount > 0 Then
                    SSProcess.CreateNewObj 2
                    SSProcess.SetNewObjValue "SSObj_Code",9210313
                    SSProcess.SetNewObjValue "SSObj_LayerName", "楼层"
                    SSProcess.SetNewObjValue "[BZ]", "新增"
                    SSProcess.SetNewObjValue "[CHZT]", 2
                    
                    For j = 0 To UBound(FildArr)
                        SSProcess.SetNewObjValue "[" & FildArr(j) & "]" , ValHArr(j)
                    Next 'j
                    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
                    SSProcess.GetObjectFocusPoint ID, x, y
                    
                    SSProcess.AddNewObjPoint x, y, 0, 0, ""
                    SSProcess.AddNewObjPoint x, y, 0, 0, ""
                    SSProcess.AddNewObjToSaveObjList
                    SSProcess.SaveBufferObjToDatabase
                End If
            Next 'i
            
            For i = 0 To HCount - 1
                ValHArr = Split(HInfoArr(i),",", - 1,1)
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", "9210313"
                SSProcess.SetSelectCondition "[ZRZBSM]", "==", ValHArr(1)
                SSProcess.SetSelectCondition "[SJC]", "==", ValHArr(3)
                SSProcess.SetSelectCondition "[CHZT]", "==", 2
                SSProcess.SelectFilter
                geoCount = SSProcess.GetSelGeoCount()
                For j = 0 To geoCount - 2
                    SSProcess.DelSelGeo j
                Next 'j
            Next 'i
            
        End If
    End If
    
    
    
    SSProcess.CloseAccessMdb FileName
    
End Function' CreateCeng

'创建层
Function CreateYCeng(ByVal FileName)
    
    SSProcess.OpenAccessMdb FileName
    
    SqlStr = "Select ZDBSM,ZRZBSM,ZRZH,SJC,MYC,QXDM From YCFW_H Where OBJECTID IS NOT NULL "
    GetSQLRecordAll EdbName,SqlStr,HInfoArr,HCount
    
    FildsStr = "ZDBSM,ZRZBSM,ZRZH,SJC,MYC,QXDM"
    FildArr = Split(FildsStr,",", - 1,1)
    
    If HCount > 0 Then
        SqlStr = "Select ZDBSM,ZRZBSM,ZRZH,SJC,MYC,QXDM From YCFW_C Where OBJECTID IS NOT NULL "
        GetSQLRecordAll EdbName,SqlStr,CInfoArr,CCount
        
        If CCount > 0 Then
            For i = 0 To CCount - 1
                ValArr = Split(CInfoArr(i),",", - 1,1)
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", "9210123"
                SSProcess.SetSelectCondition "[BSM]", "==", ValArr(1)
                SSProcess.SelectFilter
                geoCount = SSProcess.GetSelGeoCount()
                
                If geoCount > 0 Then
                    SSProcess.CreateNewObj 2
                    SSProcess.SetNewObjValue "SSObj_Code",9210313
                    SSProcess.SetNewObjValue "SSObj_LayerName", "楼层"
                    SSProcess.SetNewObjValue "[BZ]", "新增"
                    SSProcess.SetNewObjValue "[CHZT]", 1
                    For j = 0 To UBound(FildArr)
                        SSProcess.SetNewObjValue "[" & FildArr(j) & "]" , ValArr(j)
                    Next 'j
                    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
                    SSProcess.GetObjectFocusPoint ID, x, y
                    SSProcess.AddNewObjPoint x, y, 0, 0, ""
                    SSProcess.AddNewObjPoint x, y, 0, 0, ""
                    SSProcess.AddNewObjToSaveObjList
                    SSProcess.SaveBufferObjToDatabase
                End If
            Next 'i
            
            For i = 0 To CCount - 1
                ValArr = Split(CInfoArr(i),",", - 1,1)
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", 9210313
                SSProcess.SetSelectCondition "[ZRZBSM]", "==", ValArr(1)
                SSProcess.SetSelectCondition "[SJC]", "==", ValArr(3)
                SSProcess.SetSelectCondition "[CHZT]", "==", 1
                SSProcess.SelectFilter
                geoCount = SSProcess.GetSelGeoCount()
                For j = 0 To geoCount - 2
                    SSProcess.DelSelGeo j
                Next 'j
            Next 'i
            
        Else
            For i = 0 To HCount - 1
                ValHArr = Split(HInfoArr(i),",", - 1,1)
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", "9210123"
                SSProcess.SetSelectCondition "[BSM]", "==", ValHArr(1)
                SSProcess.SelectFilter
                geoCount = SSProcess.GetSelGeoCount()
                If geoCount > 0 Then
                    SSProcess.CreateNewObj 2
                    SSProcess.SetNewObjValue "SSObj_Code",9210313
                    SSProcess.SetNewObjValue "SSObj_LayerName", "楼层"
                    SSProcess.SetNewObjValue "[BZ]", "新增"
                    SSProcess.SetNewObjValue "[CHZT]", 1
                    
                    For j = 0 To UBound(FildArr)
                        SSProcess.SetNewObjValue "[" & FildArr(j) & "]" , ValHArr(j)
                    Next 'j
                    ID = SSProcess.GetSelGeoValue(0,"SSObj_ID")
                    SSProcess.GetObjectFocusPoint ID, x, y
                    
                    SSProcess.AddNewObjPoint x, y, 0, 0, ""
                    SSProcess.AddNewObjPoint x, y, 0, 0, ""
                    SSProcess.AddNewObjToSaveObjList
                    SSProcess.SaveBufferObjToDatabase
                End If
            Next 'i
            
            For i = 0 To HCount - 1
                ValHArr = Split(HInfoArr(i),",", - 1,1)
                SSProcess.ClearSelection
                SSProcess.ClearSelectCondition
                SSProcess.SetSelectCondition "SSObj_Code", "=", "9210313"
                SSProcess.SetSelectCondition "[ZRZBSM]", "==", ValHArr(1)
                SSProcess.SetSelectCondition "[SJC]", "==", ValHArr(3)
                SSProcess.SetSelectCondition "[CHZT]", "==", 1
                SSProcess.SelectFilter
                geoCount = SSProcess.GetSelGeoCount()
                For j = 0 To geoCount - 2
                    SSProcess.DelSelGeo j
                Next 'j
            Next 'i
            
        End If
    End If
    
    SSProcess.CloseAccessMdb FileName
    
End Function' CreateYCeng