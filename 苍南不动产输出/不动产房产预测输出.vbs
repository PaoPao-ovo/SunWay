Dim GZLTJcount
Dim arGZLTJ
Dim zdFields(10000),TableName(10000)
Dim arvalues(10000)


'--------------------------ģ��h��Ϣ��jzqzb���׼����Ӧ��jzbsb�Ƿ�ɾ��


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
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"�ڵػ�����Ϣ��"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"�����õ�ʹ��Ȩ��ַ���"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"����"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"����ʸ��ע��"
  'SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"ʸ����"

  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"�ں���ַ��"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"��Ȼ��"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"����ʸ��ע��"
  SSProcess.SetDataXParameter "ExportLayer"&CStr(AddOne(startIndex)),"��"


  startIndex = 0
  SSProcess.SetDataXParameter "LayerRelationCount","192"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"�ڵػ�����Ϣ��:::ZDJBXX::"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"�����õ�ʹ��Ȩ��ַ���:JZD::::"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"����:::ZRZ::"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"����ʸ��ע��::::FCSL_ZJ:FCSL_ZJ"
  'SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"ʸ����:::FCSL_H::"

  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�:::ZDJBXX::"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��:JZD::::"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"�ں���ַ��:JZD::::"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"��Ȼ��:::ZRZ::"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"����ʸ��ע��::::FCSL_ZJ:FCSL_ZJ"
  SSProcess.SetDataXParameter "LayerRelation"&CStr(AddOne(startIndex)),"��:::FCSL_H::"

  startIndex = 0
  SSProcess.SetDataXParameter "TableFieldDefCount","100000"
'�ڵػ�����Ϣ��
'����,����(0��,1��,2��,3ע��,10�����湲��),EPS�ֶ���,�ͻ��ֶ���,[�ͻ��ֶα���,]ϵͳ�ֶ���,ȱʡֵ,�ֶ�����,�ֶγ���,С��λ"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,BSM,BSM,��ʶ��,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,YSDM,YSDM,Ҫ�ش���,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZDDM,ZDDM,�ڵش���,,,dbText:1,19,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,BDCDYH,BDCDYH,��������Ԫ��,,,dbText:1,28,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZDTZM,ZDTZM,�ڵ�������,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZL,ZL,����,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZDMJ,ZDMJ,�ڵ����,,,dbSingle:1,15,4"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,MJDW,MJDW,�����λ,,,dbText:1,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,YT,YT,��;,,,dbText:1,6,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,YTNAME,GHYTMC,��;����,,,dbMemo:1,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,DJ,DJ,�ȼ�,,,dbText,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,JG,JG,�۸�,,,dbSingle,15,4"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,JGDW,JGDW,�۸�λ,,,dbText,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,QLLX,QLLX,Ȩ������,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,QLXZ,QLXZ,Ȩ������,,,dbText,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,QLSDFS,QLSDFS,Ȩ���趨��ʽ,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,RJL,RJL,�ݻ���,,,dbSingle,4,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,JZMD,JZMD,�����ܶ�,,,dbSingle,3,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,JZXG,RJJZXGL,�����޸�,,,dbSingle,5,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZDSZD,ZDSZD,�ڵ�����-��,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZDSZN,ZDSZN,�ڵ�����-��,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZDSZX,ZDSZX,�ڵ�����-��,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZDSZB,ZDSZB,�ڵ�����-��:1,,,dbText,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,BZ,BZ,��ע,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,JZZMJ,JZMJ,�������,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,JZZDZMJ,JZWZDMJ,������ռ�����,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,SYQQSSJ,SYQQSSJ,ʹ��Ȩ��ʼʱ��,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,SYQJSSJ,SYQJSSJ,ʹ��Ȩ����ʱ��,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,TDSYQX,TDSYQX,����ʹ������,,,dbInteger,5,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,DCDWMC,DCDWMC,�����浥λ����,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,KJCS,KJCS,�ռ����,,,dbInteger,5,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,GDHTH,GDHTH,���غ�ͬ��,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,LHL,LHL,�̻���,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,QJKBSM,QJKBSM,ϵͳ��ʶ��,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZT,BGZT,״ֵ̬,,,dbInteger:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,SUOYQR,SUOYQR,����Ȩ��,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,SJYT,SJYT,ʵ����;,,,dbMemo,400,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,PZYT,PZYT,��׼��;,,,dbText,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,DBR,DBR,������,,,dbText,255,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,DaiLR,DLR,������,,,dbText,255,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,GYQLR,GYQLR,����Ȩ������Ϣ,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,PZMJ,PZMJ,��׼���,,,dbSingle,10,4"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,GMJJHYFLDM,GMJJHYDM,���񾭼���ҵ����,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,SM,SM,˵��,,,dbMemo,2000,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,QSDCJS,QSDCJS,Ȩ���������,,,dbMemo,2000,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,DCR,DCY,����Ա,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,DCRQ,DCRQ,��������,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,DJCLJS,DJCLJS,�ؼ���������,,,dbText,2000,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,CeLY,CLY,����Ա,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,CeLRQ,CLRQ,��������,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,YBZDDM,YBZDDM,Ԥ���ڵش���,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZDBLC,BLC,������,,,dbText,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,JZDWSM,JZDWSM,��ַ��λ˵��,,,dbMemo,500,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,ZYQSJXZXSM,JZXZXSM,��ַ������˵��,,,dbMemo,500,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,DJDCJGSHYJ,SHYJ,�ؼ�������������,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,SHR,SHR,�����,,,dbText,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"ʹ��Ȩ�ڵ�,2,SHRQ,SHRQ,�������,,,dbDate,15,2"

'�ڵؽ�ַ��
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,BSM,BSM,��ʶ��,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,ZDZHDM,ZDZHDM,�ڵ�/�ں�����,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,YSDM,YSDM,Ҫ�ش���,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,JZDH,JZDH,��ַ���,,,dbText,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,SXH,SXH,˳���,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,JBLX,JBLX,�������,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,JZDLX,JZDLX,��ַ������,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,XZBZ,XZBZ,X����ֵ,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,YZBZ,YZBZ,Y����ֵ,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,ZZBZ,ZZBZ,Z����ֵ,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,BZ,BZ,��ע,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,QLSDFS,QLSDFS,Ȩ���趨��ʽ,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,QJKBSM,QJKBSM,Ȩ�����ʶ��,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ڵؽ�ַ��,0,STATE,BGZT,״ֵ̬,,,dbInteger:1,2,0"

'�ں���ַ��
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,BSM,BSM,��ʶ��,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,ZDZHDM,ZDZHDM,�ڵ�/�ں�����,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,YSDM,YSDM,Ҫ�ش���,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,JZDH,JZDH,��ַ���,,,dbText,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,SXH,SXH,˳���,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,JBLX,JBLX,�������,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,JZDLX,JZDLX,��ַ������,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,XZBZ,XZBZ,X����ֵ,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,YZBZ,YZBZ,Y����ֵ,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,ZZBZ,ZZBZ,Z����ֵ,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,BZ,BZ,��ע,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,QLSDFS,QLSDFS,Ȩ���趨��ʽ,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,QJKBSM,QJKBSM,Ȩ�����ʶ��,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"�ں���ַ��,0,STATE,BGZT,״ֵ̬,,,dbInteger:1,2,0"
'����
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,BSM,BSM,��ʶ��,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,YSDM,YSDM,Ҫ�ش���,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,BDCDYH,BDCDYH,��������Ԫ��,,,dbText,28,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZDDM,ZDDM,�ڵش���,,,dbText:1,19,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZRZH,ZRZH,����,,,dbText:1,24,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZL,LDZL,¥������,,,dbMemo:1,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,XMMC,XMMC,��Ŀ����,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,JZWMC,JZWMC,����������,,,dbText,100,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,JGRQ,JGRQ,��������,,,dbDate,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,JZWGD,JZWGD,������߶�,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZZDMJ,ZZDMJ,��ռ�����,,,dbSingle,15:1,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZYDMJ,ZYDMJ,���õ����,,,dbSingle:1,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,YCJZMJ,YCJZMJ,Ԥ�⽨�����,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,SCJZMJ,SCJZMJ,ʵ�⽨�����,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZCS,ZCS,�ܲ���,,,dbInteger:1,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,DSCS,DSCS,���ϲ���,,,dbMemo:1,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,DXCS,DXCS,���²���,,,dbMemo:1,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,GHYT,GHYT,�滮��;,,,dbText:1,3,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,GHYTMC,YTMC,�滮��;����,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,FWJG,FWJG,�����ṹ,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZTS,ZTS,������,,,dbLong:1,6,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,BZ,BZ,��ע,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,GHJZMJ,GHJZMJ,�滮�������,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,GHXKZH,GHXKZH,�滮���֤��,,,dbText,30,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,YSXKZH,YSXKZH,Ԥ�����֤��,,,dbText,30,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,YSHZZH,YSHZZH,���պ�׼֤��,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,YCDXJZMJ,YCDXMJ,Ԥ��������,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,YCQTJZMJ,YCQTMJ,Ԥ���������,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,SCDXJZMJ,SCDXMJ,ʵ��������,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,SCQTJZMJ,SCQTMJ,ʵ���������,,,dbSingle,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,XQMC,XQMC,С������,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,DYS,DYS,��Ԫ��,,,dbInteger,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,QLSDFS,QLSDFS,Ȩ���趨��ʽ,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,JZWZT,JZWZT,������״̬,,,dbText:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZDBSM,ZDBSM,�ڵر�ʶ��,,,dbLong,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,QJKBSM,QJKBSM,Ȩ�����ʶ��,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,ZT,BGZT,״ֵ̬,,,dbInteger:1,2,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,DMZ,DMZ,ʩ�������,,,dbText,255,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,JZXS,JZXS,������ʽ,,,dbText:1,255,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,DYMJ,DYMJ,�������,,,dbDouble,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,DZWDM,DZWDM,������������,,,dbText,30,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,SYQMJ,SYQMJ,ʹ��Ȩ���,,,dbDouble,15,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��Ȼ��,2,FTTDMJ,FTTDMJ,��̯�������,,,dbDouble,15,3"

'����ʸ��ע��
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,BSM,BSM,��ʶ��,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,ZJNR,ZJNR,ע������,,,dbText:1,200,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,ZT,ZT,����,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,YS,YS,��ɫ,,,dbText,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,BS,BS,����,,,dbInteger,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,XZ,XZ,��״,,,dbText,1,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,XHX,XHX,�»���,,,dbText,1,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,KD,KD,���,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,GD,GD,�߶�,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,ZJFX,ZJFX,ע�Ƿ���,,,dbSingle,10,6"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,ZRZBSM,ZRZBSM,��Ȼ����ʶ��,,,dbLong:1,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,SJC,SJC,ʵ�ʲ�,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,BDCDYH,BDCDYH,����������Ԫ��,,,dbText,28,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,DYZJ,DYZJ,��ӡע��,,,dbInteger:1,1,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,LXMC,LXMC,ע����������,,,dbText:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,BZ,BZ,��ע,,,dbText,50,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,QJKBSM,QJKBSM,Ȩ�����ʶ��,,,dbLong,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"����ʸ��ע��,3,BGZT,BGZT,״ֵ̬,,,dbInteger:1,2,0"

'����ʸ����
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,BSM,BSM,��ʶ��,,,dbLong:1,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,ZRZBSM,ZBDCDYH,��Ȼ����ʶ��,,,dbLong:1,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,BDCDYH,BDCDYH,����������Ԫ��,,,dbText,28,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,SJC,SJC,ʵ�ʲ�,,,dbInteger:1,4,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,HH,HH,����,,,dbText,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,TNJZMJ,TNMJ,�������,,,dbSingle:1,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,FTJZMJ,FTMJ,��̯���,,,dbSingle,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,JZMJ,MJ,���,,,dbSingle,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,MJJSGS,MJJSGS,������㹫ʽ,,,dbMemo,,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,GC1,GC1,(�ײ�)�߳�,,,dbSingle,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,GC2,GC2,(����)�߳�,,,dbSingle,18,3"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,CG,CG,���,,,dbSingle,15,2"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,FTLX,FTLX,��̯����,,,dbInteger:1,5,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,MLXMC,MLXMC,����������,,,dbText:1,20,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,QJKBSM,QJKBSM,Ȩ�����ʶ��,,,dbInteger,10,0"
  SSProcess.SetDataXParameter "TableFieldDef"&CStr(AddOne(startIndex)),"��,2,ZT,BGZT,״ֵ̬,,,dbInteger:1,2,0"

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
zdFields(9)="���,��������,Ʒ���ͺ�,�������,�ȼ�����,�����춨��Ч��"
zdFields(10)="���,�������,�����;"
zdFields(11)="���,��������,������,��������λ,��ע"
zdFields(12)="���,����,ְ�ƻ�ְҵ�ʸ�,��Ҫ����ְ��,��ע"
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

'-----------------------------------------��дmdb�ֶ�ֵ-------------------------------------------
function SetMDBTableInfo(strFileName,GZLTJvalues,GZLTJcount,Fieldsvalue,TableNamevalue)
mdbName=replace(strFileName,".edb",".mdb")
  SSProcess.OpenAccessMdb mdbName 
	sql= "SELECT " & Fieldsvalue & " FROM "& TableNamevalue 

	'�򿪼�¼��
	SSProcess.OpenAccessRecordset mdbName, sql
   SSFunc.ScanString GZLTJvalues, ";", arvalues, valueCount
	'��Ӽ�¼
	for i =0 to GZLTJcount -1
		'strValues =replace(GZLTJvalues(i),"*","")
		strValues =arvalues(i+1)
'msgbox strValues
		SSProcess.AddAccessRecord mdbName, sql, Fieldsvalue,strValues
	next
	SSProcess.CloseAccessRecordset mdbname, sql
end function


'------------------------------------��ȡ���ֶ�ֵ---------------------------------------------------
Function GetTableInfo(ByRef MdbFileName,ByVal strSQL,ByRef GZLTJvalues,ByRef GZLTJcount)

SSProcess.OpenAccessMdb MdbFileName
	sql =strSQL
	EdbName =MdbFileName
	'�򿪼�¼��
	SSProcess.OpenAccessRecordset EdbName, sql
	'��ȡ��¼����
	RecordCount =SSProcess.GetAccessRecordCount (EdbName, sql)
	'ReDim arRecord(RecordCount)				'���¶����ַ������
	if RecordCount >0 then
		'����¼�α��Ƶ���һ��
		SSProcess.AccessMoveFirst EdbName, sql
      i=0
GZLTJvalues=" "
		'�����¼
		While (SSProcess.AccessIsEOF (EdbName, sql ) = False)
			'��ȡ��ǰ��¼����
			SSProcess.GetAccessRecord EdbName, sql, fields, values
'msgbox sql
         'GZLTJvalues(i)=values
GZLTJvalues=GZLTJvalues+";"+values
'msgbox values
			'�ƶ���¼�α�
			SSProcess.AccessMoveNext EdbName, sql
         i=i+1
		Wend
	end if
'msgbox GZLTJvalues
	'�رռ�¼��
GZLTJcount=i
	SSProcess.CloseAccessRecordset EdbName, sql
   SSProcess.CloseAccessMdb EdbName 
End Function

'-----------------------------------------mdb������ά��-----------------------------------------------------------------
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
GZLTJFields = "[���] varchar(10) NULL," +"[��������] varchar(255) NULL," +"[������] varchar(200) NULL," +"[��������λ] varchar(100) NULL," +"[��ע] varchar(255) NULL" 
fieldStructs=  GZLTJFields
CreateTable mdbName, "INFO_GZLTJ",fieldStructs

RYXXFields = ""
RYXXFields =  "[���] varchar(10) NULL," +"[����] varchar(100) NULL," +"[ְ�ƻ�ְҵ�ʸ�] varchar(200) NULL," +"[��Ҫ����ְ��] varchar(200) NULL," +"[��ע] varchar(255) NULL" 
fieldStructs=  RYXXFields
CreateTable mdbName, "INFO_RYXX",fieldStructs

YQSBFields = ""
YQSBFields =  YQSBFields +"[���] varchar(10) NULL," +"[��������] varchar(200) NULL," +"[Ʒ���ͺ�] varchar(200) NULL," +"[�������] varchar(100) NULL," +"[�ȼ�����] varchar(100) NULL," +"[�����춨��Ч��] varchar(100) NULL"
fieldStructs=  YQSBFields
CreateTable mdbName, "INFO_YQSB",fieldStructs

RJPZFields = ""
RJPZFields =  "[���] varchar(10) NULL," +"[�������] varchar(200) NULL," +"[�����;] varchar(255) NULL" 
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