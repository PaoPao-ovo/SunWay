
'========================================================Excel操作对象和文件路径操作对象======================================================

'路径操作对象
Dim FileSysObj
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Excel操作对象
Dim ExcelObj
Set ExcelObj = CreateObject("Excel.Application")

'============================================================检查记录配置====================================================================

'检查集项目名称
Dim strGroupName
strGroupName = "管线精度检查"

'检查集组名称
Dim strCheckName
strCheckName = "坐标精度检查"

'检查模型名称
Dim CheckmodelName
CheckmodelName = "自定义脚本检查类->坐标精度检查"

'检查描述
Dim strDescription
strDescription = "坐标精度超标"

'=========================================================管线点编码配置====================================================================

'管线点Code
Dim GxPoiCodes
GxPoiCodes = "32100004,32100005,32100006,32100007,32100009,32100010,32110111,32110121,32110131,32111401,32111501,32114401,32114501,33060211,33060221,33060231,33060241,33060251,34050411,34050421,34050431,34050441,34050451,34050461,34050471,34050481,34050491,34050511,34050521,34050531,34050541,34050551,34050561,34050571,34050581,34050591,34050711,34050721,34050731,34050741,34050751,34050761,34050771,34050781,34050791,34050911,34050921,34050931,34050941,34050951,34050961,34050971,34050981,34050991,38050111,38050121,38050131,38050141,38050151,38050212,38050213,38050214,38050215,38050216,38050222,38050224,38050225,38050226,38050227,38050232,38050233,38050234,38050235,38050236,38050411,38050421,38050431,38050441,38050451,38050511,38050541,38050711,38050721,38050731,38050741,38050751,38050761,38050771,38050781,38050791,38051111,38051121,38051131,38051141,38051151,38051211,38051221,38051231,38051241,38051251,38051311,38051321,38051331,38051341,38051351,45100612,45100613,45100614,45100615,45100616,45100622,45100623,45100624,45100625,45100626,45101111,45101121,45101131,45101141,45101151,45101161,45101171,45101181,45101191,46040812,46040814,46040816,46040817,46040818,46040819,46040820,46040821,46040822,46040823,46040824,46040825,46040826,46040827,46040828,46040829,46040830,46040831,46040832,46040833,46040834,46040835,46040838,46040839,46040840,46040842,46040843,46040844,46040845,46040846,46040847,46040848,46040850,46040851,46040852,46040853,46040854,51011312,51019901,51019902,51019903,51019904,51019905,51019906,51019907,51019908,51019909,51019910,51031101,51031302,51031402,51031601,51041101,52011302,52011601,52019901,52019902,52019903,52019904,52019905,52019906,52019907,52019908,52019909,52019910,53011302,53019901,53019902,53019903,53019904,53019905,53019906,53019907,53019908,53019909,53019910,53022302,53029901,53029902,53029903,53029904,53029905,53029906,53029907,53029908,53029909,53029910,53033302,53039901,53039902,53039903,53039904,53039905,53039906,53039907,53039908,53039909,53039910,53044302,53049901,53049902,53049903,53049904,53049905,53049906,53049907,53049908,53049909,53049910,53051102,53052102,53053102,53054102,54039901,54039902,54039903,54039904,54039905,54039906,54039907,54039908,54039909,54039910,54040101,54049901,54049902,54049903,54049904,54049905,54049906,54049907,54049908,54049909,54049910,54050101,54110101,54110201,54110202,54110301,54110401,54110501,54119901,54119902,54119903,54119904,54119905,54119906,54119907,54119908,54119909,54119910,54120101,54120201,54120202,54120301,54120401,54120501,54129901,54129902,54129903,54129904,54129905,54129906,54129907,54129908,54129909,54129910,54130101,54130201,54130202,54130301,54130401,54130501,54139901,54139902,54139903,54139904,54139905,54139906,54139907,54139908,54139909,54139910,54140101,54140201,54140202,54140301,54140401,54140501,54149901,54149902,54149903,54149904,54149905,54149906,54149907,54149908,54149909,54149910,54150101,54150201,54150202,54150301,54150401,54150501,54159901,54159902,54159903,54159904,54159905,54159906,54159907,54159908,54159909,54159910,54210101,54210201,54210701,54219901,54219902,54219903,54219904,54219905,54219906,54219907,54219908,54219909,54219910,54220101,54220201,54220701,54229901,54229902,54229903,54229904,54229905,54229906,54229907,54229908,54229909,54229910,54230101,54230201,54230701,54239901,54239902,54239903,54239904,54239905,54239906,54239907,54239908,54239909,54239910,54240101,54240201,54240701,54249901,54249902,54249903,54249904,54249905,54249906,54249907,54249908,54249909,54249910,54250101,54250201,54250701,54259901,54259902,54259903,54259904,54259905,54259906,54259907,54259908,54259909,54259910,54260101,54260201,54260701,54269901,54269902,54269903,54269904,54269905,54269906,54269907,54269908,54269909,54269910,54270101,54270201,54270701,54279901,54279902,54279903,54279904,54279905,54279906,54279907,54279908,54279909,54279910,54280101,54280201,54280701,54289901,54289902,54289903,54289904,54289905,54289906,54289907,54289908,54289909,54289910,54290101,54290201,54290701,54299901,54299902,54299903,54299904,54299905,54299906,54299907,54299908,54299909,54299910,54310501,54310701,54310901,54311001,54311101,54311201,54311301,54311401,54311501,54319901,54319902,54319903,54319904,54319905,54319906,54319907,54319908,54319909,54319910,54319913,54340501,54340701,54340901,54341001,54341101,54341201,54341301,54341401,54341501,54349901,54349902,54349903,54349904,54349905,54349906,54349907,54349908,54349909,54349910,54349913,54410301,54410401,54410501,54411101,54411201,54412101,54412201,54412212,54419901,54419902,54419903,54419904,54419905,54419906,54419907,54419908,54419909,54419910,54419913,54420301,54420401,54420501,54422101,54422201,54423101,54423201,54423212,54429901,54429902,54429903,54429904,54429905,54429906,54429907,54429908,54429909,54429910,54429913,54430201,54430301,54430401,54430501,54433101,54439901,54439902,54439903,54439904,54439905,54439906,54439907,54439908,54439909,54439910,54439913,54450201,54450212,54450301,54450401,54450501,54452101,54452111,54452201,54459901,54459902,54459903,54459904,54459905,54459906,54459907,54459908,54459909,54459910,54459913,54510101,54510201,54510301,54510401,54510601,54510701,54510801,54511501,54511511,54519901,54519902,54519903,54519904,54519905,54519906,54519907,54519908,54519909,54519910,54520101,54520201,54520301,54520401,54520601,54520701,54520801,54521501,54522501,54529901,54529902,54529903,54529904,54529905,54529906,54529907,54529908,54529909,54529910,54530101,54530201,54530301,54530401,54530601,54530701,54530801,54531501,54533501,54539901,54539902,54539903,54539904,54539905,54539906,54539907,54539908,54539909,54539910,54540101,54540201,54540301,54540401,54540501,54540601,54540701,54540801,54549901,54549902,54549903,54549904,54549905,54549906,54549907,54549908,54549909,54549910,54610501,54610701,54610801,54610901,54611001,54611101,54611201,54611301,54611401,54619901,54619902,54619903,54619904,54619905,54619906,54619907,54619908,54619909,54619910,54620501,54620701,54620801,54620901,54621001,54621101,54621201,54621301,54621401,54629901,54629902,54629903,54629904,54629905,54629906,54629907,54629908,54629909,54629910,54630501,54630701,54630801,54630901,54631001,54631101,54631201,54631301,54631401,54639901,54639902,54639903,54639904,54639905,54639906,54639907,54639908,54639909,54639910,54720501,54720701,54729901,54729902,54729903,54729904,54729905,54729906,54729907,54729908,54729909,54729910,54730501,54730701,54739901,54739902,54739903,54739904,54739905,54739906,54739907,54739908,54739909,54739910"

'=============================================================功能入口=======================================================================

Sub OnClick()
    
    AllVisible
    
    FileSysObj.CopyFile  SSProcess.GetSysPathName(7) & "输出模板\" & "测量精度调查表模板.xlsx",SSProcess.GetSysPathName(5) & "测量精度\" & "测量精度调查表.xlsx"
    
    OpenExcel SSProcess.GetSysPathName(5) & "测量精度\" & "测量精度调查表.xlsx",ExcleFile
    
    SetTableHeader
    
    GetPoiNameZero ZeorCount,ZeroArr,ZeroStr
    
    StartRow = 3
    
    Dim InsertCount
    InsertCount = 0
    
    For i = 0 To UBound(ZeroArr)
        CheckX = SSProcess.GetObjectAttr(ZeroArr(i),"SSObj_X")
        CheckY = SSProcess.GetObjectAttr(ZeroArr(i),"SSObj_Y")
        OriginPoiIds = SSProcess.SearchNearObjIDs(CheckX,CheckY,0.1,0,GxPoiCodes,0)
        OriginPoiArr = Split(OriginPoiIds,",", - 1,1)
        OriginCount = UBound(OriginPoiArr) + 1
        InsertCount = InsertCount + OriginCount
    Next 'i
    
    If InsertCount - 3 > 0 Then
        InsertRows 3,InsertCount - 3
    End If
    
    For i = 0 To UBound(ZeroArr)
        CheckX = SSProcess.GetObjectAttr(ZeroArr(i),"SSObj_X")
        CheckY = SSProcess.GetObjectAttr(ZeroArr(i),"SSObj_Y")
        OriginPoiIds = SSProcess.SearchNearObjIDs(CheckX,CheckY,0.1,0,GxPoiCodes,0)
        OriginPoiArr = Split(OriginPoiIds,",", - 1,1)
        OriginCount = UBound(OriginPoiArr) + 1
        InsertExcel ZeroArr(i),OriginPoiArr,OriginCount,StartRow,EndRow
    Next 'i

    GetResultInfo ZeorCount,3,EndRow,OverPoiCount,OverHeightCount,AveragePoi,AverageHeight,MiddlePoi,MiddleHei,OverPercent,OverPoiName,OverHeiName,MaxLen,MaxHei

    InsertResult EndRow + 1,ZeorCount,OverPoiCount,OverHeightCount,AveragePoi,AverageHeight,MiddlePoi,MiddleHei,OverPercent
    
    InsertBz 3,EndRow

    DelSelCol 6
    
    CloseExcel ExcleFile
    
    FreshTable MaxLen,MaxHei

    ClearCheckRecord
    
    AddRecord OverPoiName,OverHeiName

End Sub' OnClick

'==========================================================Excel填值===================================================================


'设置表名
Function SetTableHeader()
    SqlStr = "Select XMMC From 管线项目信息表 Where 管线项目信息表.ID = 1"
    GetSQLRecordAll SqlStr,Xmmc,Count
    ExcelObj.Cells(1,1) = Xmmc(0) & "精度检查表"
End Function' SetTableHeader

'填写表格
Function InsertExcel(ByVal ZeroId,ByVal OriginPoiArr,ByVal OriginCount,ByRef StartRow,ByRef EndRow)
    CheckX = SSProcess.GetObjectAttr(ZeroId,"SSObj_X")
    CheckY = SSProcess.GetObjectAttr(ZeroId,"SSObj_Y")
    CheckZ = SSProcess.GetObjectAttr(ZeroId,"SSObj_Z")
    
    EndRow = StartRow + OriginCount - 1
    
    j = 0
    For i = StartRow To EndRow
        ExcelObj.Cells(i,1) = i - 2
        ExcelObj.Cells(i,2) = SSProcess.GetObjectAttr(OriginPoiArr(j),"[WTDH]")
        InsertXYZ i,ZeroId,OriginPoiArr(j)
        j = j + 1
    Next 'i
    StartRow = EndRow + 1
End Function' InsertExcel

'获取结果信息
Function GetResultInfo(ByVal ZeroCount,ByVal StartRow,ByVal EndRow,ByRef OverPoiCount,ByRef OverHeightCount,ByRef AveragePoi,ByRef AverageHeight,ByRef MiddlePoi,ByRef MiddleHei,ByRef OverPercent,ByRef OverPoiName,ByRef OverHeiName,ByRef MaxLen,ByRef MaxHei)
    OverPoiCount = 0
    OverHeightCount = 0
    OverNum = 0
    AveragePoi = 0.00
    AverageHeight = 0.00
    SquarePoi = 0
    SquareHei = 0
    MiddlePoi = 0.00
    MiddleHei = 0.00
    OverPoiName = ""
    OverHeiName = ""
    MaxLen = 0.00
    MaxHei = 0.00
    For i = StartRow To EndRow
        If Transform(ExcelObj.Cells(i,10)) > 0.05 Then
            OverPoiCount = OverPoiCount + 1
            If OverPoiName = "" Then
                OverPoiName = "'" & ExcelObj.Cells(i,2) & "'"
            Else
                OverPoiName = OverPoiName & "," & "'" & ExcelObj.Cells(i,2) & "'"
            End If
        End If
        If Transform(ExcelObj.Cells(i,11)) > 0.03 Then
            OverHeightCount = OverHeightCount + 1
            If OverHeiName = "" Then
                OverHeiName = "'" & ExcelObj.Cells(i,2) & "'"
            Else
                OverHeiName = OverHeiName & "," & "'" & ExcelObj.Cells(i,2) & "'"
            End If
        End If
        If Transform(ExcelObj.Cells(i,10)) > MaxLen Then
            MaxLen = Transform(ExcelObj.Cells(i,10))
        End If
        If Transform(ExcelObj.Cells(i,11)) > MaxHei Then
            MaxHei = Transform(ExcelObj.Cells(i,11))
        End If
        TotalPoi = TotalPoi + Transform(ExcelObj.Cells(i,10))
        TotalHei = TotalHei + Transform(ExcelObj.Cells(i,11))
        SquarePoi = SquarePoi + Transform(ExcelObj.Cells(i,10)) ^ 2
        SquareHei = SquareHei + Transform(ExcelObj.Cells(i,11)) ^ 2
        If Transform(ExcelObj.Cells(i,10)) > 0.05 Or Transform(ExcelObj.Cells(i,11)) > 0.03 Then
            OverNum = OverNum + 1
        End If
    Next 'i
    OverPercent = Round((OverNum / ZeroCount),2) * 100
    AveragePoi = Round(TotalPoi / ZeroCount,3)
    AverageHeight = Round(TotalHei / ZeroCount,3)
    MiddlePoi = Round(Sqr(SquarePoi / ZeroCount),3)
    MiddleHei = Round(Sqr(SquareHei / ZeroCount),3)
End Function' GetResultInfo

'填写统计结果
Function InsertResult(ByVal ResultRow,ByVal ZeroCount,ByVal OverPoiCount,ByVal OverHeightCount,ByVal AveragePoi,ByVal AverageHeight,ByVal MiddlePoi,ByVal MiddleHei,ByVal OverPercent)
    If ZeroCount < 15 Then
        If OverPoiCount > 0 And OverHeightCount > 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，点位超差的点数" & OverPoiCount & "个，高程超差的个数" & OverHeightCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差平均值：" & AveragePoi & "（M），高程误差平均值：" & AverageHeight & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount > 0 And OverHeightCount = 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，点位超差的点数" & OverPoiCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差平均值：" & AveragePoi & "（M），高程误差平均值：" & AverageHeight & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount = 0 And OverHeightCount > 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个" & "高程超差的个数" & OverHeightCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差平均值：" & AveragePoi & "（M），高程误差平均值：" & AverageHeight & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        Else
            ResultString = "统计结果：检查点数：" & ZeroCount & "个" & "点位误差平均值：" & AveragePoi & "（M），高程误差平均值：" & AverageHeight & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        End If
    ElseIf ZeroCount >= 15 Then
        If OverPoiCount > 0 And OverHeightCount > 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，点位超差的点数" & OverPoiCount & "个，高程超差的个数" & OverHeightCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差中误差：" & MiddlePoi & "（M），高程误差中误差：" & MiddleHei & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount > 0 And OverHeightCount = 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，点位超差的点数" & OverPoiCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差中误差：" & MiddlePoi & "（M），高程误差中误差：" & MiddleHei & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        ElseIf OverPoiCount = 0 And OverHeightCount > 0 Then
            ResultString = "统计结果：检查点数：" & ZeroCount & "个，高程超差的个数" & OverHeightCount & "个，超差点位占总点数百分比" & OverPercent & "%，点位误差中误差：" & MiddlePoi & "（M），高程误差中误差：" & MiddleHei & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        Else
            ResultString = "统计结果：检查点数：" & ZeroCount & "个" & "点位误差中误差：" & MiddlePoi & "（M），高程误差中误差：" & MiddleHei & "（M）"
            ExcelObj.Cells(ResultRow,1) = ResultString
        End If
    End If
End Function' InsertResult

'填写XYZ值并计算差值
Function InsertXYZ(ByVal InsertRow,ByVal ZeroId,ByVal OriginId)
    ExcelObj.Cells(InsertRow,3) = Round(Transform(SSProcess.GetObjectAttr(OriginId,"SSObj_Y")),3)
    ExcelObj.Cells(InsertRow,4) = Round(Transform(SSProcess.GetObjectAttr(OriginId,"SSObj_X")),3)
    ExcelObj.Cells(InsertRow,5) = Round(Transform(SSProcess.GetObjectAttr(OriginId,"SSObj_Z")),3)
    ExcelObj.Cells(InsertRow,7) = Round(Transform(SSProcess.GetObjectAttr(ZeroId,"SSObj_Y")),3)
    ExcelObj.Cells(InsertRow,8) = Round(Transform(SSProcess.GetObjectAttr(ZeroId,"SSObj_X")),3)
    ExcelObj.Cells(InsertRow,9) = Round(Transform(SSProcess.GetObjectAttr(ZeroId,"SSObj_Z")),3)
    ExcelObj.Cells(InsertRow,10) = Round(CalDiff(SSProcess.GetObjectAttr(OriginId,"SSObj_Y"),SSProcess.GetObjectAttr(OriginId,"SSObj_X"),SSProcess.GetObjectAttr(ZeroId,"SSObj_Y"),SSProcess.GetObjectAttr(ZeroId,"SSObj_X")),3)
    ExcelObj.Cells(InsertRow,11) = Round(Abs(Transform(SSProcess.GetObjectAttr(ZeroId,"SSObj_Z")) - Transform(SSProcess.GetObjectAttr(OriginId,"SSObj_Z"))),3)
End Function' InsertXYZ

'填写备注
Function InsertBz(ByVal StartRow,ByVal EndRow)
    For i = StartRow To EndRow
        If Transform(ExcelObj.Cells(i,10)) > 0.05 And Transform(ExcelObj.Cells(i,11)) > 0.03 Then
            ExcelObj.Cells(i,12) = "平面超限、高程超限"
        ElseIf Transform(ExcelObj.Cells(i,10)) > 0.05 And (ExcelObj.Cells(i,11)) <= 0.03 Then
            ExcelObj.Cells(i,12) = "平面超限"
        ElseIf Transform(ExcelObj.Cells(i,10)) <= 0.05 And (ExcelObj.Cells(i,11)) > 0.03 Then
            ExcelObj.Cells(i,12) = "高程超限"
        End If
    Next 'i
End Function' InsertBz

'获取所有的0点的点名和ID
Function GetPoiNameZero(ByRef ZeorCount,ByRef ZeroArr(),ByRef ZeroStr)
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "0"
    SSProcess.SelectFilter
    ZeorCount = SSProcess.GetSelGeoCount
    ReDim  ZeroArr(ZeorCount - 1)
    For i = 0 To ZeorCount - 1
        ZeroArr(i) = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        If ZeroStr = "" Then
            ZeroStr = SSProcess.GetSelGeoValue(i,"SSObj_ID")
        Else
            ZeroStr = ZeroStr & "," & SSProcess.GetSelGeoValue(i,"SSObj_ID")
        End If
    Next 'i
End Function' GetPoiNameZero

'添加检查记录
Function AddRecord(ByVal OverPoiName,ByVal OverHeiName)
    If OverPoiName <> "" Then
        SqlString = "Select 地下管线点属性表.ID From 地下管线点属性表 Inner Join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And 地下管线点属性表.WTDH In " & "(" & OverPoiName & ")"
        GetSQLRecordAll SqlString,OverPoiArr,OverPoiCount
        For i = 0 To OverPoiCount - 1
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(OverPoiArr(i),"SSObj_X"),SSProcess.GetObjectAttr(OverPoiArr(i),"SSObj_Y"),0,0,OverPoiArr(i),""
        Next 'i
    End If
    strCheckName = "高程精度检查"
    CheckmodelName = "自定义脚本检查类->高程精度检查"
    strDescription = "高程精度超标"
    If OverHeiName <> "" Then
        SqlString = "Select 地下管线点属性表.ID From 地下管线点属性表 Inner Join GeoPointTB on 地下管线点属性表.ID = GeoPointTB.ID WHERE (GeoPointTB.Mark Mod 2)<>0 And 地下管线点属性表.WTDH In " & "(" & OverHeiName & ")"
        GetSQLRecordAll SqlString,OverHeiArr,OverHeiCount
        For i = 0 To OverHeiCount - 1
            SSProcess.AddCheckRecord strGroupName,strCheckName,CheckmodelName,strDescription,SSProcess.GetObjectAttr(OverHeiArr(i),"SSObj_X"),SSProcess.GetObjectAttr(OverHeiArr(i),"SSObj_Y"),0,0,OverHeiArr(i),""
        Next 'i
    End If
    SSProcess.ShowCheckOutput
    SSProcess.SaveCheckRecord
End Function' AddRecord

'===========================================================工具类函数===============================================================

'打开所有图层
Function AllVisible()
    Count = SSProcess.GetLayerCount
    For i = 0 To Count - 1
        Layername = SSProcess.GetLayerName (i)
        SSProcess.SetLayerStatus Layername, 1, 1
    Next
    SSProcess.RefreshView
End Function

'打开Excel表
Function OpenExcel(ByVal FilePath,ByRef ExcleFile)
    ExcelObj.Application.Visible = False
    Set ExcleFile = ExcelObj.WorkBooks.Open(FilePath)
    Set ExcelSheet = ExcleFile.WorkSheets("精度调查表")
    ExcelSheet.Activate
End Function

'保存关闭Excel表格
Function CloseExcel(ByVal ExcleFile)
    ExcleFile.Save
    ExcelObj.Quit
End Function' CloseExcel

'插入指定行数
Function InsertRows(ByVal StartRow,ByVal InsertCount)
    For i = 0 To InsertCount - 1
        ExcelObj.ActiveSheet.Rows(StartRow).Insert
    Next 'i
End Function' InsertRows

'删除指定的列
Function DelSelCol(ByVal ColNum)
    ExcelObj.ActiveSheet.Columns(ColNum).Delete
End Function' DelSelCol

'计算距离差值
Function CalDiff(ByVal X1,ByVal Y1,ByVal X2,ByVal Y2)
    CalDiff = Sqr((Transform(X1) - Transform(X2)) ^ 2 + (Transform(Y1) - Transform(Y2)) ^ 2)
End Function' CalDiff

'获取所有记录
Function GetSQLRecordAll(ByVal StrSqlStatement, ByRef SQLRecord(), ByRef iRecordCount)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    If StrSqlStatement = "" Then
        MsgBox "查询语句为空，操作停止！",48
    End If
    iRecordCount =  - 1
    SSProcess.OpenAccessRecordset ProJectName, StrSqlStatement
    RecordCount = SSProcess.GetAccessRecordCount (ProJectName, StrSqlStatement)
    If RecordCount > 0 Then
        iRecordCount = 0
        ReDim SQLRecord(RecordCount)
        SSProcess.AccessMoveFirst ProJectName, StrSqlStatement
        iRecordCount = 0
        While SSProcess.AccessIsEOF (ProJectName, StrSqlStatement) = 0
            fields = ""
            values = ""
            SSProcess.GetAccessRecord ProJectName, StrSqlStatement, fields, values
            SQLRecord(iRecordCount) = values
            iRecordCount = iRecordCount + 1
            SSProcess.AccessMoveNext ProJectName, StrSqlStatement
        WEnd
    End If
    SSProcess.CloseAccessRecordset ProJectName, StrSqlStatement
    SSProcess.CloseAccessMdb ProJectName
End Function

'数据类型转换
Function Transform(ByVal Values)
    If Values <> "" Then
        Values = CDbl(Values)
    Else
        Values = 0
        Exit Function
    End If
    Transform = Values
End Function'Transform

'清空检查集
Function ClearCheckRecord()
    SSProcess.RemoveCheckRecord strGroupName, strCheckName
End Function' ClearCheckRecord

'刷新二维表
Function FreshTable(ByVal MaxLen,ByVal MaxHei)
    ProJectName = SSProcess.GetProjectFileName
    SSProcess.OpenAccessMdb ProJectName
    SqlStr = "Update  管线项目信息表 Set DWZDJC = " & MaxLen & " Where 管线项目信息表.ID= 1"
    SSProcess.ExecuteAccessSql  ProJectName,SqlStr
    SqlStr = "Update  管线项目信息表 Set GCZDJC = " & MaxHei & " Where 管线项目信息表.ID= 1"
    SSProcess.ExecuteAccessSql  ProJectName,SqlStr
    SSProcess.CloseAccessMdb ProJectName
End Function