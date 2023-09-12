' /*
'  * @Description: 请填写简介
'  * @Author: LHY
'  * @Date: 2023-08-21 11:16:51
'  * @LastEditors: LHY
'  * @LastEditTime: 2023-09-12 16:42:27
'  */

Dim GZLTJcount
Dim arGZLTJ
Dim zdFields(10000),TableName(10000)
Dim arvalues(10000)


'--------------------------模板h信息表、jzqzb与标准不对应，jzbsb是否删掉


Sub OnClick()
        strProjectName=SSProcess.GetProjectFileName()
        fileName=replace(strProjectName,".edb",".mdb")

  SSProcess.ClearDataXParameter
  SSProcess.SetDataXParameter "DataType","22"
  SSProcess.SetDataXParameter "FeatureCodeTBName","FeatureCodeTB_500"
  SSProcess.SetDataXParameter "SymbolScriptTBName","SymbolScriptTB_500"
  SSProcess.SetDataXParameter "NoteTemplateTBName","NoteTemplateTB_500"
  SSProcess.SetDataXParameter "ExportPathName",fileName
  SSProcess.SetDataXParameter "DataBoundMode","0"
  SSProcess.SetDataXParameter "SymbolExplodeMode","1"
  SSProcess.SetDataXParameter "LayerUseStatus","0"
  SSProcess.SetDataXParameter "AddSystemFieldMode","0"
  SSProcess.SetDataXParameter "EXCHANGE_PDB_ExportEmptyLayer","1"


  startIndex = 0
  SSProcess.SetDataXParameter "ExportLayerCount","192"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"宗地基本信息表"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"建设用地使用权界址点表"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"幢表"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"房产矢量注记"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"矢量户"

  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"使用权宗地"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"宗地界址点"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"宗海界址点"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"自然幢"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"房产矢量注记"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"户"


  startIndex = 0
  SSProcess.SetDataXParameter "LayerRelationCount","192"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"宗地基本信息表:::ZDJBXX::"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"建设用地使用权界址点表:JZD::::"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"幢表:::ZRZ::"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"房产矢量注记::::FCSL_ZJ:FCSL_ZJ"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"矢量户:::FCSL_H::"

  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"使用权宗地:::ZDJBXX::"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"宗地界址点:JZD::::"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"宗海界址点:JZD::::"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"自然幢:::ZRZ::"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"房产矢量注记::::FCSL_ZJ:FCSL_ZJ"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"户:::FCSL_H::"

  startIndex = 0
  SSProcess.SetDataXParameter "TableFieldDefCount","100000"
'宗地基本信息表
'层名,类型(0点,1线,2面,3注记,10点线面共用),EPS字段名,客户字段名,[客户字段别名,]系统字段名,缺省值,字段类型,字段长度,小数位"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,BSM,BSM,标识码,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,YSDM,YSDM,要素代码,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZDDM,ZDDM,宗地代码,,,dbText:1,19,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,BDCDYH,BDCDYH,不动产单元号,,,dbText:1,28,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZDTZM,ZDTZM,宗地特征码,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZL,ZL,坐落,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZDMJ,ZDMJ,宗地面积,,,dbSingle:1,15,4"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,MJDW,MJDW,面积单位,,,dbText:1,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,YT,YT,用途,,,dbText:1,6,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,YTNAME,GHYTMC,用途名称,,,dbMemo:1,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,DJ,DJ,等级,,,dbText,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,JG,JG,价格,,,dbSingle,15,4"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,JGDW,JGDW,价格单位,,,dbText,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,QLLX,QLLX,权利类型,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,QLXZ,QLXZ,权利性质,,,dbText,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,QLSDFS,QLSDFS,权利设定方式,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,RJL,RJL,容积率,,,dbSingle,4,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,JZMD,JZMD,建筑密度,,,dbSingle,3,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,JZXG,RJJZXGL,建筑限高,,,dbSingle,5,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZDSZD,ZDSZD,宗地四至-东,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZDSZN,ZDSZN,宗地四至-南,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZDSZX,ZDSZX,宗地四至-西,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZDSZB,ZDSZB,宗地四至-北:1,,,dbText,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,BZ,BZ,备注,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,JZZMJ,JZMJ,建筑面积,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,JZZDZMJ,JZWZDMJ,建筑物占地面积,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,SYQQSSJ,SYQQSSJ,使用权起始时间,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,SYQJSSJ,SYQJSSJ,使用权结束时间,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,TDSYQX,TDSYQX,土地使用期限,,,dbInteger,5,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,DCDWMC,DCDWMC,调查测绘单位名称,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,KJCS,KJCS,空间层数,,,dbInteger,5,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,GDHTH,GDHTH,供地合同号,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,LHL,LHL,绿化率,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,QJKBSM,QJKBSM,系统标识码,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZT,BGZT,状态值,,,dbInteger:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,SUOYQR,SUOYQR,所有权人,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,SJYT,SJYT,实际用途,,,dbMemo,400,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,PZYT,PZYT,批准用途,,,dbText,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,DBR,DBR,代表人,,,dbText,255,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,DaiLR,DLR,代理人,,,dbText,255,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,GYQLR,GYQLR,共有权利人信息,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,PZMJ,PZMJ,批准面积,,,dbSingle,10,4"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,GMJJHYFLDM,GMJJHYDM,国民经济行业代码,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,SM,SM,说明,,,dbMemo,2000,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,QSDCJS,QSDCJS,权属调查记事,,,dbMemo,2000,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,DCR,DCY,调查员,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,DCRQ,DCRQ,调查日期,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,DJCLJS,DJCLJS,地籍测量记事,,,dbText,2000,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,CeLY,CLY,测量员,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,CeLRQ,CLRQ,测量日期,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,YBZDDM,YBZDDM,预编宗地代码,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZDBLC,BLC,比例尺,,,dbText,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,JZDWSM,JZDWSM,界址点位说明,,,dbMemo,500,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,ZYQSJXZXSM,JZXZXSM,界址线走向说明,,,dbMemo,500,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,DJDCJGSHYJ,SHYJ,地籍调查结果审核意见,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,SHR,SHR,审核人,,,dbText,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"使用权宗地,2,SHRQ,SHRQ,审核日期,,,dbDate,15,2"

'宗地界址点
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,BSM,BSM,标识码,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,ZDZHDM,ZDZHDM,宗地/宗海代码,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,YSDM,YSDM,要素代码,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,JZDH,JZDH,界址点号,,,dbText,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,SXH,SXH,顺序号,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,JBLX,JBLX,界标类型,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,JZDLX,JZDLX,界址点类型,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,XZBZ,XZBZ,X坐标值,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,YZBZ,YZBZ,Y坐标值,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,ZZBZ,ZZBZ,Z坐标值,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,BZ,BZ,备注,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,QLSDFS,QLSDFS,权利设定方式,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,QJKBSM,QJKBSM,权籍库标识码,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗地界址点,0,STATE,BGZT,状态值,,,dbInteger:1,2,0"

'宗海界址点
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,BSM,BSM,标识码,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,ZDZHDM,ZDZHDM,宗地/宗海代码,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,YSDM,YSDM,要素代码,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,JZDH,JZDH,界址点号,,,dbText,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,SXH,SXH,顺序号,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,JBLX,JBLX,界标类型,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,JZDLX,JZDLX,界址点类型,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,XZBZ,XZBZ,X坐标值,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,YZBZ,YZBZ,Y坐标值,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,ZZBZ,ZZBZ,Z坐标值,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,BZ,BZ,备注,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,QLSDFS,QLSDFS,权利设定方式,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,QJKBSM,QJKBSM,权籍库标识码,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"宗海界址点,0,STATE,BGZT,状态值,,,dbInteger:1,2,0"
'幢表
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,BSM,BSM,标识码,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,YSDM,YSDM,要素代码,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,BDCDYH,BDCDYH,不动产单元号,,,dbText,28,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZDDM,ZDDM,宗地代码,,,dbText:1,19,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZRZH,ZRZH,幢号,,,dbText:1,24,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZL,LDZL,楼幢坐落,,,dbMemo:1,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,XMMC,XMMC,项目名称,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,JZWMC,JZWMC,建筑物名称,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,JGRQ,JGRQ,竣工日期,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,JZWGD,JZWGD,建筑物高度,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZZDMJ,ZZDMJ,幢占地面积,,,dbSingle,15:1,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZYDMJ,ZYDMJ,幢用地面积,,,dbSingle:1,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,YCJZMJ,YCJZMJ,预测建筑面积,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,SCJZMJ,SCJZMJ,实测建筑面积,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZCS,ZCS,总层数,,,dbInteger:1,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,DSCS,DSCS,地上层数,,,dbMemo:1,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,DXCS,DXCS,地下层数,,,dbMemo:1,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,GHYT,GHYT,规划用途,,,dbText:1,3,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,GHYTMC,YTMC,规划用途名称,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,FWJG,FWJG,建筑结构,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZTS,ZTS,总套数,,,dbLong:1,6,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,BZ,BZ,备注,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,GHJZMJ,GHJZMJ,规划建筑面积,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,GHXKZH,GHXKZH,规划许可证号,,,dbText,30,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,YSXKZH,YSXKZH,预售许可证号,,,dbText,30,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,YSHZZH,YSHZZH,验收核准证号,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,YCDXJZMJ,YCDXMJ,预测地下面积,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,YCQTJZMJ,YCQTMJ,预测其他面积,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,SCDXJZMJ,SCDXMJ,实测地下面积,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,SCQTJZMJ,SCQTMJ,实测其他面积,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,XQMC,XQMC,小区名称,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,DYS,DYS,单元数,,,dbInteger,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,QLSDFS,QLSDFS,权利设定方式,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,JZWZT,JZWZT,建筑物状态,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZDBSM,ZDBSM,宗地标识码,,,dbLong,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,QJKBSM,QJKBSM,权籍库标识码,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,ZT,BGZT,状态值,,,dbInteger:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,DMZ,DMZ,施工幢编号,,,dbText,255,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,JZXS,JZXS,建筑形式,,,dbText:1,255,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,DYMJ,DYMJ,独用面积,,,dbDouble,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,DZWDM,DZWDM,定着物特征码,,,dbText,30,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,SYQMJ,SYQMJ,使用权面积,,,dbDouble,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"自然幢,2,FTTDMJ,FTTDMJ,分摊土地面积,,,dbDouble,15,3"

'房产矢量注记
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,BSM,BSM,标识码,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,ZJNR,ZJNR,注记内容,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,ZT,ZT,字体,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,YS,YS,颜色,,,dbText,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,BS,BS,磅数,,,dbInteger,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,XZ,XZ,形状,,,dbText,1,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,XHX,XHX,下划线,,,dbText,1,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,KD,KD,宽度,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,GD,GD,高度,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,ZJFX,ZJFX,注记方向,,,dbSingle,10,6"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,ZRZBSM,ZRZBSM,自然幢标识码,,,dbLong:1,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,SJC,SJC,实际层,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,BDCDYH,BDCDYH,户不动产单元号,,,dbText,28,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,DYZJ,DYZJ,打印注记,,,dbInteger:1,1,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,LXMC,LXMC,注记类型名称,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,BZ,BZ,备注,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,QJKBSM,QJKBSM,权籍库标识码,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"房产矢量注记,3,BGZT,BGZT,状态值,,,dbInteger:1,2,0"

'房产矢量户
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,BSM,BSM,标识码,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,ZRZBSM,ZBDCDYH,自然幢标识码,,,dbLong:1,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,BDCDYH,BDCDYH,户不动产单元号,,,dbText,28,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,SJC,SJC,实际层,,,dbInteger:1,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,HH,HH,户号,,,dbText,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,TNJZMJ,TNMJ,套内面积,,,dbSingle:1,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,FTJZMJ,FTMJ,分摊面积,,,dbSingle,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,JZMJ,MJ,面积,,,dbSingle,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,MJJSGS,MJJSGS,面积计算公式,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,GC1,GC1,(底部)高程,,,dbSingle,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,GC2,GC2,(顶部)高程,,,dbSingle,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,CG,CG,层高,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,FTLX,FTLX,分摊类型,,,dbInteger:1,5,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,MLXMC,MLXMC,面类型名称,,,dbText:1,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,QJKBSM,QJKBSM,权籍库标识码,,,dbInteger,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"户,2,ZT,BGZT,状态值,,,dbInteger:1,2,0"

  SSProcess.ExportData

EWBvalue

End Sub

function EWBvalue()
zdFields(0)="CH,ZRZBSM,YSDM,SJC,MYC,BSM,QJKBSM,BGZT"
zdFields(1)="BDCDYH,ZDDM,ZRZBSM,YSDM,ZL,MJDW,SJCS,HH,SHBW,FWYT1,FWYT2,FWYT3,YTMC,YCJZMJ,YCTNJZMJ,YCFTJZMJ,YCFTXS,SCJZMJ,SCTNJZMJ,SCFTJZMJ,SCFTXS,GYTDMJ,FTTDMJ,DYTDMJ,FWLX,FWLXMC,FWXZ,FWXZMC,DYH,SZC,SJC,MYC,SFJSMJ,ZFBDCDYH,FFLB,TDSYQSSJ,TDSYJSSJ,TDSYQX,DCDWMC,XMMC,BSM,QJKBSM,BGZT,SFRF,GHYT,GHJZMJ,MPH"
zdFields(2)="BSM,ZDDM,BDCDYH,XMMC,JSDWMC,KFDWMC,ZL,TDLYWJ,JSYDGHXKZ,JSGCXKZ,JSGCHSQRS,JGBAZMS,SPFYSXKZ,ZJZMJ,JGSJ,ZZS,BZ"
zdFields(3)="BDCDYH,DKMC,ZCS,DXCS,JZJG,ZZMJ,ZZTS,FZZMJ,FZZTS,SYMJ,SYTS,BGMJ,BGTS,SQYFMJ,SQYFTS,WYJYMJ,WYJYTS,WYGLMJ,WYGLTS,YLMJ,YLTS,QTFZZMJ,QTFZZTS,BZ"
zdFields(4)="BSM,BZJZDMC,YDYBSM,XDYBSM,DYBM,YBDCDYH,XBDCDYH,BZ"
zdFields(5)="BSM,YDYBSM,YBDCDYH,XBDCDYH,XDYBSM,BZ,QJKBSM,BGZT"
zdFields(6)="JZDDH,JZDLX,JBLX,JJ,JZXLB,JZXWZ,JZXXZ,JZXSM,LZDJH,LZDMC,QDH,ZJDH,ZDH,LZZJRQ,TDQLR,LZZJR,BZZJR,BZZJRQ,LZQLR,BZZDDM,LZZDDM,TBRQ,TBR,RQ"
zdFields(7)="QLR_QLR,QLR_QLRLX,QLR_ZJZL,QLR_ZJH,QLR_DZ,QLR_DH,QLR_HJSZSS,QLR_YZBM,QLR_XB,QLR_GJ,QLR_FZJG,QLR_GZDW,OWNER_CONTACT,QLR_DZYJ,QLR_SSHY,QLR_SFCZR,QLR_QLBL,QLR_GYFS,QLR_GYQK,OWNED_PORTION,QLR_YQLRMC,QLR_BZ,QLR_ZDDM,QLR_ID"
'zdFields(8)="JZDDH,JZDLX,JBLX,JJ,JZXLB,JZXWZ,JZXXZ,JZXSM,LZDJH,LZDMC,QDH,ZJDH,ZDH,LZZJRQ,TDQLR,LZZJR,BZZJR,BZZJRQ,LZQLR,BZZDDM,LZZDDM,TBRQ,TBR,RQ"
zdFields(8)="key,value"
zdFields(9)="序号,仪器名称,品牌型号,仪器编号,等级精度,仪器检定有效性"
zdFields(10)="序号,软件名称,软件用途"
zdFields(11)="序号,工作内容,工作量,工作量单位,备注"
zdFields(12)="序号,姓名,职称或职业资格,主要工作职责,备注"
zdFields(13)="FCXH,FCJTBW,FCJSGZJFF,FCBZ,GHXH,GHJTBW,GHJSGZJFF,GHBZ"
zdFields(14)="FCXH,FCBGBW,FCJTQK,FCBH,FCSFBA,FCFW,FCMJ,FCBZ,GHXH,GHBGBW,GHJTQK,GHBH,GHSFBA,GHFW,GHMJ,GHBZ"
TableName(0)="C"
TableName(1)="H"
TableName(2)="JSXMXX"
TableName(3)="ZRZYTHZ"
TableName(4)="BHGX"
TableName(5)="HYSCGX"
TableName(6)="JZQZB"
TableName(7)="QLRXX"
'TableName(8)="JZBSB"
TableName(8)="PROJECTINFO"
TableName(9)="INFO_YQSB"
TableName(10)="INFO_RJPZ"
TableName(11)="INFO_GZLTJ"
TableName(12)="INFO_RYXX"
TableName(13)="TSBWJSSMB"
TableName(14)="JZMJBG"

strFileName=SSProcess.GetProjectFileName()
CreateTableInfo strFileName

for i=0 to 14
Fieldsvalue=zdFields(i)
TableNamevalue=TableName(i) 
	sql = "SELECT " & Fieldsvalue & " FROM " & TableNamevalue
	GetTableInfo strFileName,sql,GZLTJvalues,GZLTJcount
   SetMDBTableInfo strFileName,GZLTJvalues,GZLTJcount,Fieldsvalue,TableNamevalue 
next

end function

'-----------------------------------------填写mdb字段值-------------------------------------------
function SetMDBTableInfo(strFileName,GZLTJvalues,GZLTJcount,Fieldsvalue,TableNamevalue)
mdbName=replace(strFileName,".edb",".mdb")
  SSProcess.OpenAccessMdb mdbName 
	sql= "SELECT " & Fieldsvalue & " FROM "& TableNamevalue 

	'打开记录集
	SSProcess.OpenAccessRecordset mdbName, sql
   SSFunc.ScanString GZLTJvalues, ";", arvalues, valueCount
	'添加记录
	for i =0 to GZLTJcount -1
		'strValues =replace(GZLTJvalues(i),"*","")
		strValues =arvalues(i+1)
'msgbox strValues
		SSProcess.AddAccessRecord mdbName, sql, Fieldsvalue,strValues
	next
	SSProcess.CloseAccessRecordset mdbname, sql
end function


'------------------------------------获取表字段值---------------------------------------------------
Function GetTableInfo(ByRef MdbFileName,ByVal strSQL,ByRef GZLTJvalues,ByRef GZLTJcount)

SSProcess.OpenAccessMdb MdbFileName
	sql =strSQL
	EdbName =MdbFileName
	'打开记录集
	SSProcess.OpenAccessRecordset EdbName, sql
	'获取记录总数
	RecordCount =SSProcess.GetAccessRecordCount (EdbName, sql)
	'ReDim arRecord(RecordCount)				'重新定义界址点数组
	if RecordCount >0 then
		'将记录游标移到第一行
		SSProcess.AccessMoveFirst EdbName, sql
      i=0
GZLTJvalues=" "
		'浏览记录
		While (SSProcess.AccessIsEOF (EdbName, sql ) = False)
			'获取当前记录内容
			SSProcess.GetAccessRecord EdbName, sql, fields, values
'msgbox sql
         'GZLTJvalues(i)=values
GZLTJvalues=GZLTJvalues+";"+values
'msgbox values
			'移动记录游标
			SSProcess.AccessMoveNext EdbName, sql
         i=i+1
		Wend
	end if
'msgbox GZLTJvalues
	'关闭记录集
GZLTJcount=i
	SSProcess.CloseAccessRecordset EdbName, sql
   SSProcess.CloseAccessMdb EdbName 
End Function

'-----------------------------------------mdb创建二维表-----------------------------------------------------------------
Function CreateTableInfo(strFileName)
mdbName=replace(strFileName,".edb",".mdb")
CFields = ""
CFields = "[CH] varchar(20) NULL," +"[ZRZBSM] int NULL," +"[YSDM] varchar(10) NULL," +"[SJC] int NULL," +"[MYC] varchar(50) NULL,"+"[BSM] int NULL," +"[QJKBSM] int NULL," +"[BGZT] int NULL"  
fieldStructs=  CFields
CreateTable mdbName, "C",fieldStructs

HFields = ""
HFields =  "[BDCDYH] varchar(28) NULL," +"[ZDDM] varchar(19) NULL," +"[ZRZBSM] int NULL," +"[YSDM] varchar(10) NULL," +"[ZL] varchar(200) NULL," +"[MJDW] varchar(2) NULL,"+"[SJCS] int NULL,"+"[HH] varchar(10) NULL,"+"[SHBW] varchar(20) NULL,"+"[FWYT1] varchar(3) NULL,"+"[FWYT2] varchar(3) NULL,"+"[FWYT3] varchar(3) NULL,"+"[YTMC] varchar(255) NULL,"+"[YCJZMJ] float NULL,"+"[YCTNJZMJ] float NULL,"+"[YCFTJZMJ] float NULL,"+"[YCFTXS] float NULL,"+"[SCJZMJ] float NULL,"+"[SCTNJZMJ] float NULL,"+"[SCFTJZMJ] float NULL,"+"[SCFTXS] float NULL,"+"[GYTDMJ] float NULL,"+"[FTTDMJ] float NULL,"+"[DYTDMJ] float NULL,"+"[FWLX] varchar(2) NULL,"+"[FWLXMC] varchar(255) NULL,"+"[FWXZ] varchar(2) NULL,"+"[FWXZMC] varchar(255) NULL,"+"[DYH] varchar(10) NULL,"+"[SZC] varchar(20) NULL,"+"[SJC] int NULL,"+"[MYC] varchar(20) NULL,"+"[SFJSMJ] int NULL,"+"[ZFBDCDYH] varchar(28) NULL,"+"[FFLB] int NULL,"+"[TDSYQSSJ] date NULL,"+"[TDSYJSSJ] date NULL,"+"[TDSYQX] int NULL,"+"[DCDWMC] varchar(100) NULL,"+"[XMMC] varchar(50) NULL,"+"[BSM] int NULL,"+"[QJKBSM] int NULL,"+"[BGZT] int NULL,"+"[SFRF] int NULL,"+"[GHYT] float NULL,"+"[GHJZMJ] double NULL,"+"[MPH] varchar(255) NULL"
fieldStructs=  HFields
CreateTable mdbName, "H",fieldStructs

JSXMXXFields = ""
JSXMXXFields =  "[BSM] int NULL," +"[ZDDM] varchar(19) NULL," +"[BDCDYH] varchar(200) NULL," +"[XMMC] varchar(100) NULL," +"[JSDWMC] varchar(100) NULL," +"[KFDWMC] varchar(100) NULL,"+"[ZL] varchar(100) NULL,"+"[TDLYWJ] varchar(50) NULL,"+"[JSYDGHXKZ] varchar(50) NULL,"+"[JSGCXKZ] varchar(50) NULL,"+"[JSGCHSQRS] varchar(50) NULL,"+"[JGBAZMS] varchar(50) NULL,"+"[SPFYSXKZ] varchar(50) NULL,"+"[ZJZMJ] double NULL,"+"[JGSJ] date NULL,"+"[ZZS] int NULL,"+"[BZ] varchar(100) NULL"
fieldStructs=  JSXMXXFields
CreateTable mdbName, "JSXMXX",fieldStructs

ZRZYTHZFields = ""
ZRZYTHZFields =  "[BDCDYH] varchar(24) NULL," +"[DKMC] varchar(100) NULL," +"[ZCS] int NULL," +"[DXCS] int NULL," +"[JZJG] varchar(100) NULL," +"[ZZMJ] double NULL," +"[ZZTS] int NULL," +"[FZZMJ] double NULL," +"[FZZTS] int NULL," +"[SYMJ] double NULL," +"[SYTS] int NULL," +"[BGMJ] double NULL," +"[BGTS] int NULL," +"[SQYFMJ] double NULL," +"[SQYFTS] int NULL," +"[WYJYMJ] double NULL," +"[WYJYTS] int NULL," +"[WYGLMJ] double NULL," +"[WYGLTS] int NULL," +"[YLMJ] double NULL," +"[YLTS] int NULL," +"[QTFZZMJ] double NULL," +"[QTFZZTS] int NULL," +"[BZ] varchar(100) NULL" 
fieldStructs=  ZRZYTHZFields
CreateTable mdbName, "ZRZYTHZ",fieldStructs

BHGXFields = ""
BHGXFields =  "[BSM] int NULL," +"[BZJZDMC] varchar(20) NULL," +"[YDYBSM] int NULL," +"[XDYBSM] int NULL," +"[DYBM] varchar(20) NULL," +"[YBDCDYH] varchar(30) NULL," +"[XBDCDYH] varchar(30) NULL," +"[BZ] varchar(20) NULL" 
fieldStructs=  BHGXFields
CreateTable mdbName, "BHGX",fieldStructs

HYSCGXFields = ""
HYSCGXFields =  "[BSM] int NULL," +"[YDYBSM] int NULL," +"[YBDCDYH] varchar(28) NULL," +"[XBDCDYH] varchar(28) NULL," +"[XDYBSM] int NULL,"+"[BZ] varchar(50) NULL," +"[QJKBSM] int NULL," +"[BGZT] int NULL" 
fieldStructs=  HYSCGXFields
CreateTable mdbName, "HYSCGX",fieldStructs

JZQZBFields = ""
JZQZBFields =  "[JZDDH] varchar(255) NULL," +"[JZDLX] varchar(255) NULL," +"[JBLX] varchar(255) NULL," +"[JJ] varchar(255) NULL," +"[JZXLB] varchar(255) NULL," +"[JZXWZ] varchar(255) NULL," +"[JZXXZ] varchar(255) NULL," +"[JZXSM] varchar(255) NULL," +"[LZDJH] varchar(255) NULL," +"[LZDMC] varchar(255) NULL," +"[QDH] varchar(255) NULL," +"[ZJDH] varchar(255) NULL," +"[ZDH] varchar(255) NULL," +"[LZZJRQ] varchar(255) NULL," +"[TDQLR] varchar(255) NULL," +"[LZZJR] varchar(255) NULL," +"[BZZJR] varchar(255) NULL," +"[BZZJRQ] varchar(255) NULL," +"[LZQLR] varchar(255) NULL," +"[BZZDDM] varchar(255) NULL," +"[LZZDDM] varchar(255) NULL," +"[TBRQ] varchar(255) NULL," +"[TBR] varchar(255) NULL," +"[RQ] varchar(255) NULL" 
fieldStructs=  JZQZBFields
CreateTable mdbName, "JZQZB",fieldStructs

QLRXXFields = ""
QLRXXFields =  "[QLR_QLR] varchar(20) NULL," +"[QLR_QLRLX] varchar(50) NULL," +"[QLR_ZJZL] varchar(100) NULL," +"[QLR_ZJH] varchar(255) NULL," +"[QLR_DZ] varchar(255) NULL," +"[QLR_DH] varchar(255) NULL," +"[QLR_HJSZSS] varchar(255) NULL," +"[QLR_YZBM] varchar(255) NULL," +"[QLR_XB] varchar(255) NULL," +"[QLR_GJ] varchar(255) NULL," +"[QLR_FZJG] varchar(255) NULL," +"[QLR_GZDW] varchar(255) NULL," +"[OWNER_CONTACT] varchar(255) NULL," +"[QLR_DZYJ] varchar(255) NULL," +"[QLR_SSHY] varchar(255) NULL," +"[QLR_SFCZR] varchar(255) NULL," +"[QLR_QLBL] varchar(255) NULL," +"[QLR_GYFS] varchar(255) NULL," +"[QLR_GYQK] varchar(255) NULL," +"[OWNED_PORTION] varchar(255) NULL," +"[QLR_YQLRMC] varchar(255) NULL," +"[QLR_BZ] varchar(255) NULL," +"[QLR_ZDDM] varchar(255) NULL," +"[QLR_ID] varchar(255) NULL" 
fieldStructs=  QLRXXFields
CreateTable mdbName, "QLRXX",fieldStructs

'JZBSBFields = ""
'JZBSBFields =  "[JZDDH] varchar(255) NULL," +"[JZDLX] varchar(255) NULL," +"[JBLX] varchar(255) NULL," +"[JJ] varchar(255) NULL," +"[JZXLB] varchar(255) NULL," +"[JZXWZ] varchar(255) NULL," +"[JZXXZ] varchar(255) NULL," +"[JZXSM] varchar(255) NULL," +"[LZDJH] varchar(255) NULL," +"[LZDMC] varchar(255) NULL," +"[QDH] varchar(255) NULL," +"[ZJDH] varchar(255) NULL," +"[ZDH] varchar(255) NULL," +"[LZZJRQ] varchar(255) NULL," +"[TDQLR] varchar(255) NULL," +"[LZZJR] varchar(255) NULL," +"[BZZJR] varchar(255) NULL," +"[BZZJRQ] varchar(255) NULL," +"[LZQLR] varchar(255) NULL," +"[BZZDDM] varchar(255) NULL," +"[LZZDDM] varchar(255) NULL" 
'fieldStructs=  JZBSBFields
'CreateTable mdbName, "JZBSB",fieldStructs

GZLTJFields = ""
GZLTJFields = "[序号] varchar(10) NULL," +"[工作内容] varchar(255) NULL," +"[工作量] varchar(200) NULL," +"[工作量单位] varchar(100) NULL," +"[备注] varchar(255) NULL" 
fieldStructs=  GZLTJFields
CreateTable mdbName, "INFO_GZLTJ",fieldStructs

RYXXFields = ""
RYXXFields =  "[序号] varchar(10) NULL," +"[姓名] varchar(100) NULL," +"[职称或职业资格] varchar(200) NULL," +"[主要工作职责] varchar(200) NULL," +"[备注] varchar(255) NULL" 
fieldStructs=  RYXXFields
CreateTable mdbName, "INFO_RYXX",fieldStructs

YQSBFields = ""
YQSBFields =  YQSBFields +"[序号] varchar(10) NULL," +"[仪器名称] varchar(200) NULL," +"[品牌型号] varchar(200) NULL," +"[仪器编号] varchar(100) NULL," +"[等级精度] varchar(100) NULL," +"[仪器检定有效性] varchar(100) NULL"
fieldStructs=  YQSBFields
CreateTable mdbName, "INFO_YQSB",fieldStructs

RJPZFields = ""
RJPZFields =  "[序号] varchar(10) NULL," +"[软件名称] varchar(200) NULL," +"[软件用途] varchar(255) NULL" 
fieldStructs=  RJPZFields
CreateTable mdbName, "INFO_RJPZ",fieldStructs

PROJECTINFOFields = ""
PROJECTINFOFields =  "[Key] varchar(255) NULL," +"[Value] Memo NULL" 
fieldStructs=  PROJECTINFOFields
CreateTable mdbName, "PROJECTINFO",fieldStructs

TSBWJSSMBFields = ""
TSBWJSSMBFields =  "[FCXH] varchar(10) NULL," +"[FCJTBW] varchar(255) NULL," +"[FCJSGZJFF] varchar(255) NULL," +"[FCBZ] varchar(255) NULL," +"[GHXH] varchar(10) NULL," +"[GHJTBW] varchar(255) NULL," +"[GHJSGZJFF] varchar(255) NULL," +"[GHBZ] varchar(255) NULL" 
fieldStructs=  TSBWJSSMBFields
CreateTable mdbName, "TSBWJSSMB",fieldStructs

JZMJBGFields = ""
JZMJBGFields =  "[FCXH] varchar(10) NULL," +"[FCBGBW] varchar(255) NULL," +"[FCJTQK] varchar(255) NULL," +"[FCBH] varchar(10) NULL," +"[FCSFBA] varchar(255) NULL," +"[FCFW] varchar(255) NULL," +"[FCMJ] double NULL," +"[FCBZ] varchar(255) NULL," +"[GHXH] varchar(10) NULL," +"[GHBGBW] varchar(255) NULL," +"[GHJTQK] varchar(255) NULL," +"[GHBH] varchar(10) NULL," +"[GHSFBA] varchar(255) NULL," +"[GHFW] varchar(255) NULL," +"[GHMJ] double NULL," +"[GHBZ] varchar(255) NULL" 
fieldStructs=  JZMJBGFields
CreateTable mdbName, "JZMJBG",fieldStructs
end function
Function CreateTable( ByVal mdbName, ByVal tableName , ByVal fieldStructs)
SSProcess.OpenAccessMdb mdbName
createTabSql = "CREATE TABLE [" & tableName & "] (" & fieldStructs & ");"
SSProcess.CreateAccessTable mdbName, createTabSql 
SSProcess.CloseAccessMdb mdbName 
end function

Function AddOne( ByRef startIndex )
	startIndex = startIndex + 1
	AddOne = startIndex
End Function