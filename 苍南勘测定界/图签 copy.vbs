Dim  fileName
Dim xmmc
Dim arID(1000),arID1(1000),arID2(1000)
Dim vArray1(2000), vArray2(2000), vArray3(2000)
Dim cvArray1(2000), cvArray2(2000), cvArray3(2000),vArray(3000)
Dim dileimc(5000)
Dim dileibm(5000)
Dim dlmchbm(5000)
Dim projectName
Dim X0
Dim Y0
Dim sxcd
Dim hxcd
Dim ztmc
Dim ztdx

Dim TKX0
Dim TKY0

Dim RoteAngle

Sub OnClick()
    
    '多个图廓，删除图廓，只保留一个
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 8888
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    If  geocount >= 2 Then
        For I = 0 To GEOCOUNT - 2
            id = SSProcess.GetSelGeoValue(0, "SSObj_ID")
            SSProcess.DeleteObject (ID)
        Next
    End If
    tukuoshuxing
    xmzj
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    
    id = SSProcess.GetSelGeoValue(0, "SSObj_ID")
    xmmc = SSProcess.GetObjectAttr(iD, "[xmmc]")
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 8888
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    
    If geocount > 0 Then
        SSProcess.SetMapStatus 1, 2
        
        For i = 0 To geocount - 1
            SSProcess.SetSelGeoValue 0, "[xmmc]", xmmc
            
            SSProcess.AddSelGeoToSaveGeoList i
        Next
        SSProcess.SetMapStatus 0, 2
        SSProcess.SaveBufferObjToDatabase
        
    End If
    
    Dim MapScale
    ztmc = "黑体"
    ztdx = 187
    MapScale = SSProcess.GetMapScale
    xs = 1000 / MapScale
    projectName = SSProcess.GetProjectFileName
    
    sql1 = "Select DISTINCT 地类图斑属性表.dlmc From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 "
    GetSQLRecordAll projectName,sql1,arSQLRecord1,iRecordCount1
    
    tbqcsl = iRecordCount1
    
    For i = 0 To iRecordCount1 - 1
        If  arSQLRecord1(i) <> "" Then
            Select Case arSQLRecord1(i)
                Case "水田"  ,"水浇地","旱地"
                gdsl = gdsl + 1
                nydsl = nydsl + 1
                GDMC = GDMC & "," & arSQLRecord1(i)
                
                Case "果园" ,"茶园","橡胶园","其他园地"
                ydsl = ydsl + 1
                nydsl = nydsl + 1
                GYMC = GYMC & "," & arSQLRecord1(i)
                Case "乔木林地","灌木林地","竹林地","红树林地","森林沼泽","灌丛沼泽","其他林地"
                ldsl = ldsl + 1
                nydsl = nydsl + 1
                LDMC = LDMC & "," & arSQLRecord1(i)
                Case "天然牧草地", "人工牧草地","沼泽草地","其他草地"
                cdsl = cdsl + 1
                nydsl = nydsl + 1
                CDMC = CDMC & "," & arSQLRecord1(i)
                Case "农村道路"
                ncdlsl = ncdlsl + 1
                nydsl = nydsl + 1
                NCDLMC = NCDLMC & "," & arSQLRecord1(i)
                Case "设施农用地", "田坎"
                nydqtsl = nydqtsl + 1
                nydsl = nydsl + 1
                NYDQTMC = NYDQTMC & "," & arSQLRecord1(i)
                Case "水库水面", "坑塘水面", "沟渠"
                nydsxsl = nydsxsl + 1
                nydsl = nydsl + 1
                NYDSXMC = NYDSXMC & "," & arSQLRecord1(i)
                Case "商业服务业设施用地", "物流仓储用地"
                sfsl = sfsl + 1
                jsydsl = jsydsl + 1
                SFMC = SFMC & "," & arSQLRecord1(i)
                Case "工业用地", "采矿用地", "盐田"
                gkydsl = gkydsl + 1
                jsydsl = jsydsl + 1
                GKMC = GKMC & "," & arSQLRecord1(i)
                Case "城镇住宅用地", "农村宅基地"
                zzydsl = zzydsl + 1
                jsydsl = jsydsl + 1
                ZZMC = ZZMC & "," & arSQLRecord1(i)
                Case "机关团体新闻出版社用地", "科教文卫用地", "公用设施用地", "公园与绿地"
                ggglsl = ggglsl + 1
                jsydsl = jsydsl + 1
                GYSSMC = GYSSMC & "," & arSQLRecord1(i)
                Case "特殊用地"
                tsydsl = tsydsl + 1
                jsydsl = jsydsl + 1
                TSMC = TSMC & "," & arSQLRecord1(i)
                Case "铁路用地", "轨道交通用地", "公路用地", "城镇村道路用地", "交通服务场站用地", "机场用地", "港口码头用地", "管道运输用地"
                jtydsl = jtydsl + 1
                jsydsl = jsydsl + 1
                JTMC = JTMC & "," & arSQLRecord1(i)
                Case "水工建筑用地"
                jsslsl = jsslsl + 1
                jsydsl = jsydsl + 1
                SGJZMC = SGJZMC & "," & arSQLRecord1(i)
                Case "空闲地"
                jsqtsl = jsqtsl + 1
                jsydsl = jsydsl + 1
                KXDMC = KXDMC & "," & arSQLRecord1(i)
                Case "河流水面", "湖泊水面", "沿海滩涂", "内陆滩涂", "沼泽地", "冰川及永久积雪"
                sysl = sysl + 1
                wlydsl = wlydsl + 1
                SYMC = SYMC & "," & arSQLRecord1(i)
                Case "盐碱地", "沙地", "裸土地", "裸岩石砾地"
                qttdsl = qttdsl + 1
                wlydsl = wlydsl + 1
                QTMC = QTMC & "," & arSQLRecord1(i)
                Case Else
            End Select
        End If
        
    Next
    
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    dkCount = SSProcess.GetSelGeoCount()
    
    Getdlmc dileimc,dileibm,shuliang
    
    count5 = shuliang + 1
    
    If nydsl = 0 Then
        count5 = count5 + 2
        tbqcsl = tbqcsl + 2
        gdsl = 2
        nydsl = 2
        GDMC = ",水田,旱地"
    End If
    If wlydsl = 0 Then
        count5 = count5 + 1
        tbqcsl = tbqcsl + 1
        sysl = sysl + 1
        wlydsl = 1
        SYMC = ",河流水面"
    End If
    If jsydsl = 0 Then
        count5 = count5 + 1
        tbqcsl = tbqcsl + 1
        zzydsl = 1
        jsydsl = 1
        ZZMC = ",农村宅基地"
    End If
    
    
    sxcd = 32.5 + (dkcount + 1) * 4.5
    hxcd = 20 + count5 * 10
    wid1 = 80 * xs / 2
    heig1 = 80 * xs
    wid2 = 80 * xs / 2
    heig2 = 80 * xs
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 8888
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    Dim arID(1000), idCount
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            TKID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            
            SSProcess.GetObjectPoint TKID, 0, x0, y0, z, pointtype, name
            SSProcess.GetObjectPoint TKID, 1, x1, y1, z, pointtype, name
            SSProcess.GetObjectPoint TKID, 2, x2, y2, z, pointtype, name
            SSProcess.GetObjectPoint TKID, 3, x3, y3, z, pointtype, name
            
            TKX0 = X0
            TKY0 = Y0
            
            SSProcess.XYSA X0,Y0,X1,Y1,Length1,Angle1,0
            
            RoteAngle = Angle1
            
            DrawArea TKX0,TKY0,x0,y0,x0+hxcd,y0,x0+hxcd,y0+sxcd,x0,y0+sxcd,9210056,"RGB(255,255,255)",3,RoteAngle

            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_Code", "==", "9210056"
            SSProcess.SelectFilter
            Count = SSProcess.GetSelGeoCount()
            
            For ll = 0 To Count - 1
                polygonID = SSProcess.GetSelGeoValue( ll, "SSObj_ID" )
                ids1 = SSProcess.SearchInPolyObjIDs(polygonID, 10, "", 0,1,1)
                If ids1 <> "" Then
                    SSFunc.ScanString ids1, ",", arID, idCount
                    For kk = 0 To idCount - 1
                        codee = SSProcess.GetObjectAttr( arID(kk), "SSObj_code" )
                        
                        If  codee <> "8888"  And codee <> "504"  And codee <> "7320"  And codee <> "1234"  And codee <> "7170" And codee <> "9120058"Then
                            SSProcess.DeleteObject arID(kk)
                            SSProcess.RefreshView
                        End If
                    Next
                End If
            Next
            
            
            '外围竖框
            
            DrawLine x0,y0,x0,y0,x0,y0 + sxcd,RoteAngle,1,"RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + hxcd,y0,x0 + hxcd,y0 + sxcd,RoteAngle,1,"RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + 3,y0 + 2,x0 + 3,y0 + sxcd - 8,RoteAngle,1, "RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + hxcd - 2,y0 + 2,x0 + hxcd - 2,y0 + sxcd - 8,RoteAngle,1, "RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + hxcd - 12,y0 + 2,x0 + hxcd - 12,y0 + sxcd - 8,RoteAngle,1, "RGB(255,255,255)", 3
            
            '内竖框
            DrawLine x0,y0,x0 + 18,y0 + 2,x0 + 18,y0 + sxcd - 8,RoteAngle,1,"RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0,y0 + sxcd,x0 + hxcd,y0 + sxcd,RoteAngle,1,"RGB(255,255,255)", 3
            
            DrawNote TKX0,TKY0,x0 + 10.5,y0 + sxcd - 19.25,RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "地块名称",3,ztmc
            
            DrawNote TKX0,TKY0,x0 + hxcd - 2 - 5,y0 + sxcd - 19.25,RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "合计",3,ztmc
            
            DrawLine x0,y0,x0 + 3,y0 + sxcd - 8,x0 + hxcd - 2,y0 + sxcd - 8,RoteAngle,1, "RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + 3,y0 + sxcd - 8,x0 + hxcd - 2,y0 + sxcd - 8,RoteAngle,1, "RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + 18,y0 + sxcd - 15.5,x0 + hxcd - 12,y0 + sxcd - 15.5,RoteAngle,1, "RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + 18,y0 + sxcd - 23,x0 + hxcd - 12,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + 3,y0 + sxcd - 30.5,x0 + hxcd - 2,y0 + sxcd - 30.5,RoteAngle,1, "RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + 3,y0 + sxcd - 8,x0 + hxcd - 2,y0 + sxcd - 8,RoteAngle,1, "RGB(255,255,255)", 3
            
            DrawLine x0,y0,x0 + 3,y0 + 2,x0 + hxcd - 2,y0 + 2,RoteAngle,1, "RGB(255,255,255)", 3
            
            '农用地左竖线
            DrawLine x0,y0,x0 + 18 + nydsl * 10,y0 + 2,x0 + 18 + nydsl * 10,y0 + sxcd - 8,RoteAngle,1, "RGB(255,255,255)", 3
            
            If nydsl <> ""Then
                DrawNote  x0,y0,x0 + 18 + nydsl * 5,y0 + sxcd - 11.5, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "农用地",3,ztmc
            End If
            
            If nydsl <> 0 Or wlydsl <> 0 Then DrawLine x0,y0,x0 + hxcd - 12 - wlydsl * 10,y0 + 2,x0 + hxcd - 12 - wlydsl * 10,y0 + sxcd - 8,RoteAngle,1, "RGB(255,255,255)", 3
            
            If wlydsl <> "" Then
                DrawNote  x0,y0,x0 + hxcd - 12 - wlydsl * 5,y0 + sxcd - 10.5, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "未利",3,ztmc
                DrawNote  x0,y0,x0 + hxcd - 12 - wlydsl * 5,y0 + sxcd - 13, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "用地",3,ztmc
            End If
            
            If jsydsl <> "" Then DrawNote  x0,y0,x0 + hxcd - 12 - wlydsl * 10 - jsydsl * 5,y0 + sxcd - 11.5, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "建设用地",3,ztmc
            
            For dk = 1 To dkcount + 1
                
                DrawLine x0,y0, x0 + 3,y0 + sxcd - 30.5 - 4.5 * dk,x0 + hxcd - 2,y0 + sxcd - 30.5 - 4.5 * dk,RoteAngle,1, "RGB(255,255,255)", 3
                If dk <> dkcount + 1 Then
                    
                    NumberChange dk,hzdk
                    
                    dkmc = "地块" & hzdk
                    
                    DrawNote x0,y0,x0 + 10.5,y0 + sxcd - 30.5 - 4.5 * dk + 2.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs,ztdx * xs, dkmc,3,ztmc
                Else
                    DrawNote x0,y0,x0 + 10.5,y0 + sxcd - 30.5 - 4.5 * dk + 2.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "合计",3,ztmc
                End If
            Next
        Next
        
        '图斑竖线
        For l = 1 To tbqcsl
            DrawLine x0,y0,x0 + 18 + l * 10,y0 + 2,x0 + 18 + l * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
        Next
        '耕地行竖线
        LJS = 0
        js = 1
        If gdsl > 0 Then
            DrawLine x0,y0,x0 + 18 + gdsl * 10,y0 + sxcd - 15.5,x0 + 18 + gdsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + gdsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "耕地",3,ztmc
            SUMTBMJ  GDMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = gdsl
        If ydsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + ydsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + ydsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + ydsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "园地",3,ztmc
            SUMTBMJ  GYMC,LJS,M2
            LJS = LJS + M2
        End If
        m = m + ydsl
        If Ldsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + ldsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + ldsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + ldsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "林地",3,ztmc
            SUMTBMJ  LDMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = m + ldsl
        If cdsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + cdsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + cdsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + cdsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "草地",3,ztmc
            SUMTBMJ  CDMC,LJS,M2
            LJS = LJS + M2
        End If
        m = m + cdsl
        
        If ncdlsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + ncdlsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + ncdlsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + ncdlsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "农村" & "\" & "道路",3,ztmc
            SUMTBMJ  NCDLMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = m + ncdlsl
        
        If nydqtsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + nydqtsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + nydqtsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + nydqtsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "其他" & "\" & "土地",3,ztmc
            SUMTBMJ  NYDQTMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = m + nydqtsl
        
        If nydsxsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + nydsxsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + nydsxsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + nydsxsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "水域或" & "\" & "水利设施" & "\" & "用地",3,ztmc
            SUMTBMJ  NYDSXMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = m + nydsxsl
        
        If sfsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + sfsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + sfsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + sfsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "商业服务" & "\" & "业用地",3,ztmc
            SUMTBMJ  SFMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = m + sfsl
        
        If gkydsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + gkydsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + gkydsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + gkydsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs,  "工矿" & "\" & "用地",3,ztmc
            SUMTBMJ  GKMC,LJS,M2
            LJS = LJS + M2
        End If
        m = m + gkydsl
        
        If zzydsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + zzydsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + zzydsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + zzydsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "住宅" & "\" & "用地",3,ztmc
            SUMTBMJ  ZZMC,LJS,M2
            LJS = LJS + M2
        End If
        m = m + zzydsl
        
        If ggglsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + ggglsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + ggglsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + ggglsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "公共管理" & "\" & "与公共" & "\" & "服务用地",3,ztmc
            SUMTBMJ  GYSSMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = m + ggglsl
        
        If tsydsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + tsydsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + tsydsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + tsydsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "特殊" & "\" & "用地",3,ztmc
            SUMTBMJ  TSMC,LJS,M2
            LJS = LJS + M2
        End If
        m = m + tsydsl
        
        If jtydsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + jtydsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + jtydsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + jtydsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "交通" & "\" & "用地",3,ztmc
            SUMTBMJ  JTMC,LJS,M2
            LJS = LJS + M2
        End If
        m = m + jtydsl
        
        If jsslsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + jsslsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + jsslsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + jsslsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "水域或" & "\" & "水利设施" & "\" & "用地",3,ztmc
            SUMTBMJ  SGJZMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = m + jsslsl
        
        If jsqtsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + jsqtsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + jsqtsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + jsqtsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "其他" & "\" & "土地",3,ztmc
            SUMTBMJ  KXDMC,LJS,M2
            LJS = LJS + M2
            
        End If
        m = m + jsqtsl
        
        If sysl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + sysl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + sysl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + sysl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "水域或" & "\" & "水利设施" & "\" & "用地",3,ztmc
            SUMTBMJ  SYMC,LJS,M2
            LJS = LJS + M2
        End If
        m = m + sysl
        
        
        If qttdsl > 0 Then
            DrawLine x0,y0,x0 + 18 + m * 10 + qttdsl * 10,y0 + sxcd - 15.5,x0 + 18 + m * 10 + qttdsl * 10,y0 + sxcd - 23,RoteAngle,1, "RGB(255,255,255)", 3
            DrawNote  x0,y0,x0 + 18 + m * 10 + qttdsl * 5,y0 + sxcd - 19.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, "其他" & "\" & "土地",3,ztmc
            SUMTBMJ  QTMC,LJS,M2
            LJS = LJS + M2
        End If
        m = m + qttdsl
        
        '填值
        DrawNote x0,y0,x0 + hxcd / 2,y0 + sxcd - 4 , RoteAngle,0, "RGB(255,255,255)", (ztdx + 22) * xs, (ztdx + 22) * xs, "土地分类汇总表",3,ztmc
        
    End If
    
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    dkCount = SSProcess.GetSelGeoCount()

    For c = 1 To dkCount
        sql = "Select SUM (地类图斑属性表.tbmj) From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE  ([GeoAreaTB].[Mark] Mod 2)<>0 and  地类图斑属性表.dkh= " & c
        GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
        If    iRecordCount > 0 Then
            DKZMJ = arSQLRecord(0)
            If DKZMJ = "" Then
                DKZMJ = 0
            End If
        Else
            
            DKZMJ = 0
        End If
        MJB4W DKZMJ
        DrawNote  x0,y0,x0 + hxcd - 7,y0 + sxcd - 30.5 - (c - 1) * 4.5 - 2.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, DKZMJ,3,ztmc
    Next
    sql = "Select SUM (地类图斑属性表.tbmj) From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE  ([GeoAreaTB].[Mark] Mod 2)<>0 "
    GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
    dikuaimj = 0
    dikuaimj = arSQLRecord(0)
    If dikuaimj = "" Then
        dikuaimj = 0
    End If

    MJB4W  dikuaimj

    DrawNote  x0,y0,x0 + hxcd - 7,y0 + 4.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs,dikuaimj,3,ztmc

    huatuli
    
    HZBZ
    
End Sub

Function makePoint35(x,y,code,color,polygonID,xmmc,zmj)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "[xmmc]", xmmc
    SSProcess.SetNewObjValue "[zdmj]", zmj
    SSProcess.SetNewObjValue "SSObj_LayerName", "乡镇属性点"
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function

Function makePoint45(x,y,code,color,polygonID,qsdw, qydh,zmj)
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", code
    SSProcess.SetNewObjValue "SSObj_Color", color
    SSProcess.SetNewObjValue "[qydh]", qydh
    SSProcess.SetNewObjValue "[qsdw]", qsdw
    SSProcess.SetNewObjValue "[zdmj]", zmj
    SSProcess.SetNewObjValue "SSObj_DataMark", polygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "村属性点"
    SSProcess.AddNewObjPoint x, y, 0, 0, ""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
End Function



Function xmzj
    MapScale = SSProcess.GetMapScale
    
    SSProcess.PushUndoMark
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9120045"
    SSProcess.SelectFilter
    geoecount = SSProcess.GetSelgeoCount

    For i = 0 To geoecount - 1
        SSProcess.DelSelgeo i
    Next
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9120035"
    SSProcess.SelectFilter
    geoecount = SSProcess.GetSelgeoCount
    'msgbox  geoecount
    For i = 0 To geoecount - 1
        SSProcess.DelSelgeo i
    Next
    
    xs = MapScale / 1000
    
    ztdx = 200 * xs
    xpl = 50
    ypl = 50
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    
    xmmc = SSProcess.GetSelGeoValue( 0, "[xmmc]" )
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 8888
    
    SSProcess.SelectFilter
    
    TKID = SSProcess.GetSelGeoValue( 0, "SSObj_ID" )
    
    SSProcess.GetObjectPoint TKID, 3, x0, y0, z, pointtype, name

    y = y0 - 20
    x = x0 + 120
    projectName = SSProcess.GetProjectFileName
    sql = "Select SUM (地类图斑属性表.tbmj) From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE  ([GeoAreaTB].[Mark] Mod 2)<>0 "
    GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
    mj1 = arSQLRecord(0)
    MJB4W mj1
    makePoint35 x,y,"510",RGB(255,0,0),4,xmmc,mj1
    
    sql1 = "Select DISTINCT 地类图斑属性表.qydh From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 "
    GetSQLRecordAll projectName,sql1,arSQLRecord1,iRecordCount1
    
    For i = 0 To  iRecordCount1 - 1
        sql2 = "Select SUM (地类图斑属性表.tbmj) From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE  ([GeoAreaTB].[Mark] Mod 2)<>0 and  地类图斑属性表.qydh= '" & arSQLRecord1(i) & "'"
        GetSQLRecordAll projectName,sql2,arSQLRecord2,iRecordCount2
        mj = arSQLRecord2(0)
        MJB4W     mj
        
        sql3 = "Select DISTINCT 地类图斑属性表.qsdw From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 and  地类图斑属性表.qydh= '" & arSQLRecord1(i) & "'"
        GetSQLRecordAll projectName,sql3,arSQLRecord3,iRecordCount3
        qsdw = arSQLRecord3(0)
        makePoint45 x,y - 15 * i * xs - 15 * xs,"511",RGB(255,0,0),4,qsdw,arSQLRecord1(i),mj
        
    Next
    
    
End Function

'图例外框
Function HZBZ
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 8888
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            
            '指南针外轮廓
            ID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            SSProcess.GetObjectPoint ID, 2, X,Y, z1, pointtype1, name1
            X1 = X - 44
            Y1 = Y - 42
            DrawArea X,Y,X,Y,X,Y1,X1,Y1,X1,Y,9210058,"RGB(255,255,255)", polygonID,RoteAngle
            DrawPoiot X,Y,X - 22,Y - 21,"912003303","RGB(255,255,255)",polygonID,RoteAngle
        Next
    End If
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", "9210058"
    SSProcess.SelectFilter
    Count = SSProcess.GetSelGeoCount()
    Dim arID(1000), idCount
    For ll = 0 To Count - 1
        polygonID = SSProcess.GetSelGeoValue( ll, "SSObj_ID" )
        ids1 = SSProcess.SearchInPolyObjIDs(polygonID, 10, "", 0,1,1)
        If ids1 <> "" Then
            SSFunc.ScanString ids1, ",", arID, idCount
            For kk = 0 To idCount - 1
                codee = SSProcess.GetObjectAttr( arID(kk), "SSObj_code" )
                If  codee <> "8888"  And codee <> "504"  And codee <> "7320"  And codee <> "1234"  And codee <> "7170" And codee <> "912003303" Then
                    SSProcess.DeleteObject arID(kk)
                    SSProcess.RefreshView
                End If
            Next
        End If
    Next
End Function

Function huatuli()
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 8888
    SSProcess.SelectFilter
    geocount = SSProcess.GetSelGeoCount()
    Dim arID(1000), idCount
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            ID = SSProcess.GetSelGeoValue( i, "SSObj_ID" )
            
            SSProcess.GetObjectPoint ID, 1, x, y, z, pointtype, name
            SSProcess.GetObjectPoint ID, 2, tlx,tly, z1, pointtype1, name1
            
            DrawDelArea x,y,tlx,tly,9210055,"RGB(255,255,255)",3,RoteAngle
            
            '删除地形
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_Code", "==", "9210055"
            SSProcess.SelectFilter
            Count = SSProcess.GetSelGeoCount()
            
            For ll = 0 To Count - 1
                polygonID = SSProcess.GetSelGeoValue( ll, "SSObj_ID" )
                ids1 = SSProcess.SearchInPolyObjIDs(polygonID, 10, "", 0,1,1)
                If ids1 <> "" Then
                    SSFunc.ScanString ids1, ",", arID, idCount
                    For kk = 0 To idCount - 1
                        codee = SSProcess.GetObjectAttr( arID(kk), "SSObj_code" )
                        If  codee <> "8888"  And codee <> "504"  And codee <> "7320"  And codee <> "1234"  And codee <> "7170" And codee <> "912003303" Then
                            SSProcess.DeleteObject arID(kk)
                            SSProcess.RefreshView
                        End If
                    Next
                End If
            Next
        Next
        
        ids = SSProcess.SearchInnerObjIDs(ID , 10 ,"504,1234,10", 0)
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
            
            Getdlmc dileimc,dileibm,shuliang
            
            For k = 0 To shuliang - 1
                DrawCode = "9120043"
                DrawColor = SSProcess.GetObjectAttr(vArray(j), "SSObj_Color")
                DrawName = dileimc(k)
                ZDrawCode = ZDrawCode & "," & DrawCode
                ZDrawColor = ZDrawColor & "," & DrawColor
                ZDrawName = ZDrawName & "," & DrawName
            Next
            '面积注记
            ZDrawCode = ZDrawCode & "," & "9210053"
            ZDrawColor = ZDrawColor & "," & "RGB(255,0,0)"
            ZDrawName = ZDrawName & "," & "面积注记"
            
            SSProcess.ClearSelection
            SSProcess.ClearSelectCondition
            SSProcess.SetSelectCondition "SSObj_Code", "==", 7170
            SSProcess.SelectFilter
            Count = SSProcess.GetSelGeoCount()
            If Count > 0 Then
                cjID = SSProcess.GetSelGeoValue( 0, "SSObj_ID" )
                DrawColor = SSProcess.GetObjectAttr(cjID, "SSObj_Color")
            End If
            
            ZDrawName = ZDrawName & "," & "村界"
            ZDrawCode = ZDrawCode & "," & "7170"
            ZDrawColor = ZDrawColor & "," & DrawColor
            
            ZDrawCode = ZDrawCode & "," & "3103013"
            ZDrawColor = ZDrawColor & "," & "RGB(255,0,0)"
            ZDrawName = ZDrawName & "," & "房屋"
            
            ZDrawCode = ZDrawCode & "," & "3802022"
            ZDrawColor = ZDrawColor & "," & "RGB(255,0,0)"
            ZDrawName = ZDrawName & "," & "栅栏"
            '不够再增加
            ZDrawCode = ZDrawCode & "," & "4403002"
            ZDrawColor = ZDrawColor & "," & "RGB(255,0,0)"
            ZDrawName = ZDrawName & "," & "小路"
            
            ZDrawCode = ZDrawCode & "," & "10"
            ZDrawColor = ZDrawColor & "," & "RGB(255,0,0)"
            ZDrawName = ZDrawName & "," & "界址坐标"
        End If
        
        HuiZHItuli tlx + 11.5 + 9.5,tly,TKID,ZDrawCode,ZDrawColor,ZDrawName,y - 11.5
        
    End If
End Function


Function HuiZHItuli(x0,y0,polygonID,ZDrawCode,ZDrawColor,ZDrawName,y2)
    
    ztmc = "宋体"
    ztdx = 250
    MapScale = SSProcess.GetMapScale
    xs = 1000 / MapScale
    ztdx = 200 * xs
    arDrawCode = Split(ZDrawCode,",")
    arDrawColor = Split(ZDrawColor,",")
    arDrawName = Split(ZDrawName,",")
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 7320
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    count5 = UBound(arDrawCode) + 2
    
    DrawNote x0,y0,x0 - 9.5 + 30,y0 - 8 , RoteAngle,0, "RGB(255,255,255)", 500 * xs, 500 * xs, "图例",polygonID,ztmc
    
    For j = 0 To UBound(arDrawCode)
        
        Select Case arDrawCode(j)
            Case "9210053"
            
            DrawLine x0,y0,x0 - 9.5 + 12,y0 - j * 15 - 24,x0 - 9.5 + 22,y0 - j * 15 - 24,RoteAngle,9210057, "RGB(255,0,0)", polygonID
            
            DrawArea x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,9210054,"RGB(255,255,255)", polygonID,RoteAngle
            
            DrawTuLiNote x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24, 0, "RGB(255,255,255)", ztdx, ztdx, arDrawName(j),polygonID,ztmc,RoteAngle
            
            DrawTuLiNote  x0,y0,x0 - 9.5 + 17,y0 - j * 15 - 26,9120016, "RGB(255,0,0)", 220 * xs, 220 * xs, "0.0044",polygonID,"黑体",RoteAngle
            
            DrawTuLiNote  x0,y0,x0 - 9.5 + 17,y0 - j * 15 - 22, 9120016, "RGB(255,0,0)", 220 * xs, 220 * xs, "水田（1）",polygonID,"黑体",RoteAngle
            
            DrawTuLiNote x0,y0,x0 - 9.5 + 10,y0 - j * 15 - 24,9120016, "RGB(255,0,0)", 220 * xs, 220 * xs, "2",polygonID,"黑体",RoteAngle
            
            Case "7170"
            
            DrawArea x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,9210054,"RGB(255,255,255)", polygonID,RoteAngle
            
            DrawLine_Cj  x0,y0,x0 - 9.5 + 10,y0 - j * 15 - 24,x0 - 9.5 + 24,y0 - j * 15 - 24,9107150, "RGB(0,0,255)", polygonID,RoteAngle
            
            DrawNote x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24,RoteAngle,0, "RGB(255,255,255)", ztdx, ztdx, "村界",polygonID,ztmc
            
            Case "1234"
            
            DrawArea x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,9210054,"RGB(255,255,255)", polygonID,RoteAngle
            
            DrawTbPoint x0,y0,x0 - 9.5 + 16,y0 - j * 15 - 24,"9120231", arDrawColor(j), polygonID,3,RoteAngle
            
            DrawNote x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24,RoteAngle,0,"RGB(255,255,255)", ztdx, ztdx, arDrawName(j),polygonID,ztmc

            DrawNote  x0,y0,x0 - 9.5 + 18,y0 - j * 15 - 24,RoteAngle,9135035,"RGB(255,0,0)", ztdx, ztdx, "J3",polygonID,ztmc

            Case "9120043"
            
            DrawArea_Tb x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,arDrawCode(j), arDrawColor(j), polygonID,dileibm(n),dileimc(n),RoteAngle

            DrawNote x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24,RoteAngle,0,"RGB(255,255,255)", ztdx, ztdx, dileimc(n),polygonID,ztmc

            n = n + 1

            Case "3103013"

            DrawArea  x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,9210054,"RGB(255,255,255)", polygonID,RoteAngle

            DrawArea  x0,y0,x0 - 9.5 + 10,y0 - j * 15 - 22,x0 - 9.5 + 24,y0 - j * 15 - 22,x0 - 9.5 + 24,y0 - j * 15 - 26,x0 - 9.5 + 10,y0 - j * 15 - 26,arDrawCode(j), arDrawColor(j), polygonID,RoteAngle

            DrawNote  x0,y0,x0 - 9.5 + 17,y0 - j * 15 - 24,RoteAngle,0,"RGB(255,255,255)", ztdx, ztdx, "砖2",polygonID,ztmc

            DrawNote  x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24,RoteAngle,0,"RGB(255,255,255)", ztdx, ztdx, arDrawName(j),polygonID,ztmc

            Case "3802022"  ,"4403002"

            DrawArea  x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,9210054,"RGB(255,255,255)", polygonID,RoteAngle

            DrawLine  x0,y0,x0 - 9.5 + 9,y0 - j * 15 - 24,x0 - 9.5 + 24,y0 - j * 15 - 24,RoteAngle,arDrawCode(j), "RGB(255,255,255)", polygonID

            DrawNote  x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24,RoteAngle,0,"RGB(255,255,255)", ztdx, ztdx, arDrawName(j),polygonID,ztmc

            Case "10"
            
            DrawArea  x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,9210054,"RGB(255,255,255)", polygonID,RoteAngle

            DrawLine x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - (28 - 1.2),RoteAngle,"1",  "RGB(255,255,255)", polygonID

            DrawLine x0,y0,x0 - 9.5 + 8,y0 - j * 15 - (28 - 1.2),x0 - 9.5 + 9.5,y0 - j * 15 - (28 - 4.1),RoteAngle,"1",  "RGB(255,255,255)", polygonID

            DrawLine x0,y0,x0 - 9.5 + 9.5,y0 - j * 15 - (28 - 4.1),x0 - 9.5 + 25,y0 - j * 15 - (28 - 4.1),RoteAngle,"1",  "RGB(255,255,255)", polygonID

            DrawTuLiNote  x0,y0,x0 - 9.5 + 17,y0 - j * 15 - 26,0, "RGB(255,255,255)", 220 * xs, 220 * xs, "Y= 542241.12",polygonID,"黑体",RoteAngle

            DrawTuLiNote  x0,y0,x0 - 9.5 + 17,y0 - j * 15 - 22, 0, "RGB(255,255,255)", 220 * xs, 220 * xs, "X=3046669.81",polygonID,"黑体",RoteAngle

            DrawNote  x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24,RoteAngle,0,"RGB(255,255,255)", ztdx, ztdx, arDrawName(j),polygonID,ztmc

            Case"504"

            DrawArea  x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,"9120013", arDrawColor(j), polygonID,RoteAngle

            DrawNote  x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24,RoteAngle,0,"RGB(255,255,255)", ztdx, ztdx, arDrawName(j),polygonID,ztmc

            Case Else
            
            DrawArea  x0,y0,x0 - 9.5 + 8,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 20,x0 - 9.5 + 26,y0 - j * 15 - 28,x0 - 9.5 + 8,y0 - j * 15 - 28,arDrawCode(j), arDrawColor(j), polygonID,RoteAngle

            DrawNote  x0,y0,x0 - 9.5 + 44,y0 - j * 15 - 24,RoteAngle,0,"RGB(255,255,255)", ztdx, ztdx, arDrawName(j),polygonID,ztmc
            
        End Select
    Next
End Function

Function Getdlmc(dileimc(),dileibm(),tbqcsl)
    projectName = SSProcess.GetProjectFileName
    sql = "Select DISTINCT 地类图斑属性表.dlmc,dlbm From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 "
    GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
    tbqcsl = iRecordCount
    
    For i = 0 To iRecordCount - 1
        SSFunc.Scanstring arSQLRecord(i),",",dlmchbm,tbcount
        If  dlmchbm(0) <> "" Then
            dileimc(i) = dlmchbm(0)
            dileibm(i) = dlmchbm(1)
        End If
    Next
    
End Function

'小数点变4位
Function MJB4W(MIANJI)
    WZ = InStr(MIANJI,".")
    CHANGDU = Len(MIANJI)
    If WZ = 0 Then
        If MIANJI <> 0then
            MIANJI = MIANJI & ".0000"
        Else
            MIANJI = MIANJI
        End If
    Else
        XSDWS = CHANGDU - WZ
        addWS = 4 - XSDWS
        For I = 1 To addWS
            MIANJI = MIANJI & "0"
        Next
    End If
End Function

Function SUMTBMJ(GDMC,HLJS,SL)
    ztdx = 187
    MapScale = SSProcess.GetMapScale
    xs = 1000 / MapScale
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    dkCount = SSProcess.GetSelGeoCount()
    
    
    CFMC = Split(GDMC,",")
    
    For Z = 1 To UBound(CFMC)
        '按长度分解字符串
        ZFCCD = Len(CFMC(Z))
        Select Case ZFCCD
            Case 2
            
            DrawNote TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 26.95, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs, CFMC(Z),3,ztmc
            Case 3
            DrawNote TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 26.95, RoteAngle,0, "RGB(255,255,255)",ztdx * xs, ztdx * xs, CFMC(Z),3,ztmc
            Case 4
            LEFTZ = Left(CFMC(Z),2)
            RightR = Right(CFMC(Z),2)
            dd = LEFTZ & "\" & RIGHTR
            
            
            DrawNote TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 26.95, RoteAngle,0, "RGB(255,255,255)", ztdx * xs,ztdx * xs, dd,3,ztmc
            Case Else
            If ZFCCD = 5 Then
                LEFTZ = Left(CFMC(Z),3)
                RightR = Right(CFMC(Z),2)
                dd = LEFTZ & "\" & RIGHTR
                DrawNote TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 26.95, RoteAngle,0, "RGB(255,255,255)", ztdx * xs,ztdx * xs, dd,3,ztmc
            End If
            If ZFCCD = 6 Then
                LEFTZ = Left(CFMC(Z),3)
                RightR = Right(CFMC(Z),3)
                dd = LEFTZ & "\" & RIGHTR
                DrawNote TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 26.95, RoteAngle,0, "RGB(255,255,255)", ztdx * xs,ztdx * xs, dd,3,ztmc
            End If
            If ZFCCD = 7 Then
                LEFTZ = Left(CFMC(Z),2)
                RightR = Left(LEFTZ,3)
                three = Left(RightR,2)
                dd = LEFTZ & "\" & RIGHTR & "\" & three
                DrawNote TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 26.95, RoteAngle,0, "RGB(255,255,255)", ztdx * xs,ztdx * xs, dd,3,ztmc
            End If
            If ZFCCD = 8 Then
                LEFTZ = Left(CFMC(Z),3)
                RightR = Left(LEFTZ,3)
                three = Left(RightR,2)
                dd = LEFTZ & "\" & RIGHTR & "\" & three
                DrawNote TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 26.95, RoteAngle,0, "RGB(255,255,255)", ztdx * xs,ztdx * xs, dd,3,ztmc
            End If
            If ZFCCD > 8 Then
                LEFTZ = Left(CFMC(Z),4)
                guodu = Right(CFMC(Z),ZFCCD - 4)
                guoduz = Left(guodu,4)
                RightR = Right(CFMC(Z),ZFCCD - 8)
                dd = LEFTZ & "\" & guoduz & "\" & RIGHTR
                DrawNote TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 26.95, RoteAngle,0, "RGB(255,255,255)", ztdx * xs,ztdx * xs, dd,3,ztmc
            End If
            
        End Select
        zongmj = 0
        For B = 1 To dkCount
            sql8 = "Select SUM (地类图斑属性表.tbmj) From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE 地类图斑属性表.DLMC= '" & CFMC(Z) & "' and ([GeoAreaTB].[Mark] Mod 2)<>0 and  地类图斑属性表.dkh= " & B
            GetSQLRecordAll projectName,sql8,arSQLRecord8,iRecordCount8
            If    iRecordCount8 > 0 Then
                MIANJI = arSQLRecord8(0)
                If MIANJI = "" Then
                    MIANJI = 0
                End If
            Else
                
                MIANJI = 0
                
            End If
            If MIANJI = 0 Then
                DrawNote  TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 30.5 - (B - 1) * 4.5 - 2.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs,MIANJI,3,ztmc
            Else
                'gai
                MJB4W MIANJI
                DrawNote  TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 30.5 - (B - 1) * 4.5 - 2.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs,MIANJI,3,ztmc
            End If
            ZONGMJ = ZONGMJ + MIANJI
            
            If  B = DKCOUNT Then
                MJB4W  ZONGMJ
                DrawNote  TKX0,TKY0,x0 + 18 + (Z + HLJS) * 10 - 5,y0 + sxcd - 30.5 - (B) * 4.5 - 2.25, RoteAngle,0, "RGB(255,255,255)", ztdx * xs, ztdx * xs,ZONGMJ,3,ztmc
            End If
        Next
        SL = UBound(CFMC)
    Next
    
End Function

'获取地类名称唯一地类图斑数量及名称编码
Function Getdlmc(dileimc(),dileibm(),tbqcsl)
    projectName = SSProcess.GetProjectFileName
    sql = "Select DISTINCT 地类图斑属性表.dlmc,dlbm From 地类图斑属性表 INNER JOIN GeoAreaTB ON 地类图斑属性表.ID = GeoAreaTB.ID WHERE ([GeoAreaTB].[Mark] Mod 2)<>0 "
    GetSQLRecordAll projectName,sql,arSQLRecord,iRecordCount
    tbqcsl = iRecordCount
    'msgbox iRecordCount
    For i = 0 To iRecordCount - 1
        SSFunc.Scanstring arSQLRecord(i),",",dlmchbm,tbcount
        If  dlmchbm(0) <> "" Then
            dileimc(i) = dlmchbm(0)
            dileibm(i) = dlmchbm(1)
        End If
    Next
    
End Function

Function GetSQLRecordAll(ByRef mdbName, ByRef StrSqlStatement, ByRef arSQLRecord(), ByRef iRecordCount)
    
    SSProcess.OpenAccessMdb mdbName
    iRecordCount =  - 1
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
    SSProcess.CloseAccessMdb mdbName
End Function

Function NumberChange(Number,BigNumber)
    number = CStr(number)
    strNumer = "1,2,3,4,5,6,7,8,9,0"
    strBigNumber = "一,二,三,四,五,六,七,八,九,十"
    CD = Len (Number)
    If CD = 1 Then
        artempNumber = Split(strNumer,",")
        artempBigNumber = Split(strBigNumber,",")
        For i = 0 To 9
            If  artempNumber(i) = Number  Then
                BigNumber = artempBigNumber(i)
            End If
        Next
    Else
        LEFTZ = Left(Number,1)
        artempNumber = Split(strNumer,",")
        artempBigNumber = Split(strBigNumber,",")
        For i = 0 To 9
            If  artempNumber(i) = LEFTZ  Then
                ONE = artempBigNumber(i)
            End If
        Next
        LEFTR = Right(Number,1)
        For i = 0 To 9
            If  artempNumber(i) = LEFTR  Then
                TWO = artempBigNumber(i)
            End If
        Next
        Select Case leftz
            Case "1"
            If LEFTR = 0 Then
                BigNumber = "十"
            Else
                BigNumber = "十" & TWO
            End If
            Case Else
            If LEFTR = 0 Then
                BigNumber = ONE & TWO
            Else
                BigNumber = ONE & "十" & TWO
            End If
        End Select
    End If
    
End Function

Function tukuoshuxing
    
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 504
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            xmsj = SSProcess.GetSelGeoValue( i, "[xmsj]" )
            xmmc = SSProcess.GetSelGeoValue( i, "[xmmc]" )
        Next
    End If
    MapScale = SSProcess.GetMapScale
    SSProcess.ClearSelection
    SSProcess.ClearSelectCondition
    SSProcess.SetSelectCondition "SSObj_Code", "==", 8888
    SSProcess.SelectFilter
    geoCount = SSProcess.GetSelGeoCount()
    If geoCount > 0 Then
        For i = 0 To geoCount - 1
            id = SSProcess.GetSelGeoValue (i,"SSObj_ID")
            SSProcess.SetObjectAttr id,"[xmmc]",xmmc
            SSProcess.SetObjectAttr id,"[ctff]",xmsj & "修编成图。"
            SSProcess.SetObjectAttr id,"[blc]",MapScale
        Next
    End If
End Function

'获取第二点坐标
Function GetXYOffset(ByVal X0,ByVal Y0,ByVal Angle,ByVal Length, ByRef X3 , ByRef Y3)
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    X3 = X0 + (Length * COSVal)
    Y3 = Y0 + (Length * SinVal)
    
End Function' GetXYOffset

'线旋转
Function DrawLine(ByVal X0,ByVal Y0,ByVal X1,ByVal Y1,ByVal X2,ByVal Y2,ByVal Angle,ByVal Code,ByVal Color,ByVal PolygonID)
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    X3 = X0 + (X1 - X0) * COSVal - (Y1 - Y0) * SinVal
    Y3 = Y0 + (X1 - X0) * SinVal + (Y1 - Y0) * COSVal
    
    X4 = X0 + (X2 - X0) * COSVal - (Y2 - Y0) * SinVal
    Y4 = Y0 + (X2 - X0) * SinVal + (Y2 - Y0) * COSVal
    
    SSProcess.CreateNewObj 1

    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_DataMark", PolygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "勘测定界图廓"
    SSProcess.AddNewObjPoint X3, Y3, 0, 0, ""
    SSProcess.AddNewObjPoint X4, Y4, 0, 0, ""

    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function' DrawLine

'注记旋转(对应MakeNote、MakeNote1、MakeNote2)
Function DrawNote(ByVal X0,ByVal Y0,ByVal X1,ByVal Y1,ByVal Angle,ByVal Code,ByVal Color,ByVal Width,ByVal Height,ByVal FontString,ByVal PolygonID,ByVal ZtMc)
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    X2 = X0 + (X1 - X0) * COSVal - (Y1 - Y0) * SinVal
    Y2 = Y0 + (X1 - X0) * SinVal + (Y1 - Y0) * COSVal
    
    SSProcess.CreateNewObj 3

    Angle = SSProcess.RadianToDeg(Angle)

    SSProcess.SetNewObjValue "SSObj_FontClass", "0"
    SSProcess.SetNewObjValue "SSObj_FontInterval", "80"
    SSProcess.SetNewObjValue "SSObj_FontString", FontString
    SSProcess.SetNewObjValue "SSObj_FontWordAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontStringAngle", Angle
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_DataMark", PolygonID
    SSProcess.SetNewObjValue "SSObj_FontName", ZtMc
    SSProcess.SetNewObjValue "SSObj_LayerName", "勘测定界图廓"
    SSProcess.SetNewObjValue "SSObj_FontInterval", "20"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth",Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint X2,Y2,0,0,""

    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function' DrawNote

'面旋转
Function DrawArea(ByVal X0,ByVal Y0,ByVal X1,ByVal Y1,ByVal X2,ByVal Y2,ByVal X3,ByVal Y3,ByVal X4,ByVal Y4,ByVal Code,ByVal Color,ByVal PolygonID,ByVal Angle)
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    ResultX1 = X0 + (X1 - X0) * COSVal - (Y1 - Y0) * SinVal
    ResultY1 = Y0 + (X1 - X0) * SinVal + (Y1 - Y0) * COSVal
    
    ResultX2 = X0 + (X2 - X0) * COSVal - (Y2 - Y0) * SinVal
    ResultY2 = Y0 + (X2 - X0) * SinVal + (Y2 - Y0) * COSVal
    
    ResultX3 = X0 + (X3 - X0) * COSVal - (Y3 - Y0) * SinVal
    ResultY3 = Y0 + (X3 - X0) * SinVal + (Y3 - Y0) * COSVal
    
    ResultX4 = X0 + (X4 - X0) * COSVal - (Y4 - Y0) * SinVal
    ResultY4 = Y0 + (X4 - X0) * SinVal + (Y4 - Y0) * COSVal
    
    SSProcess.CreateNewObj 2

    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_LayerName", "勘测定界图廓"
    SSProcess.AddNewObjPoint ResultX1,ResultY1, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX2,ResultY2, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX3,ResultY3, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX4,ResultY4, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX1,ResultY1, 0, 0, ""

    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function' DrawArea

'点旋转
Function DrawPoiot(ByVal X0,ByVal Y0,ByVal X1,ByVal Y1,ByVal Code,ByVal Color,ByVal PolygonID,ByVal Angle)
    
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    ResultX1 = X0 + (X1 - X0) * COSVal - (Y1 - Y0) * SinVal
    ResultY1 = Y0 + (X1 - X0) * SinVal + (Y1 - Y0) * COSVal
    
    Angle = 90 - SSProcess.RadianToDeg(Angle)
    
    ' Angle = SSProcess.DegToDms(Angle)
    
    SSProcess.CreateNewObj 0
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_DataMark", PolygonID
    SSProcess.SetNewObjValue "SSObj_Angle", Angle
    SSProcess.SetNewObjValue "SSObj_LayerName", "勘测定界图廓"
    SSProcess.AddNewObjPoint ResultX1,ResultY1,0,0,""
    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function' DrawPoiot

'图例范围框
Function DrawDelArea(ByVal RotationX1,ByVal RotationY1,ByVal RotationX2,ByVal RotationY2,ByVal Code,ByVal Color,ByVal PolygonID,ByVal Angle)
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    ResultX1 = RotationX1 + 2 * COSVal
    ResultY1 = RotationY1 + 2 * SinVal
    
    ResultX2 = RotationX1 + 71.5 * COSVal
    ResultY2 = RotationY1 + 71.5 * SinVal
    
    ResultX3 = RotationX2 + 71.5 * COSVal
    ResultY3 = RotationY2 + 71.5 * SinVal
    
    ResultX4 = RotationX2 + 2 * COSVal
    ResultY4 = RotationY2 + 2 * SinVal
    
    SSProcess.CreateNewObj 2

    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_LayerName", "勘测定界图廓"
    SSProcess.AddNewObjPoint ResultX1,ResultY1, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX2,ResultY2, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX3,ResultY3, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX4,ResultY4, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX1,ResultY1, 0, 0, ""

    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function' DrawDelArea

'图例注记Note(对应MakeNote3)
Function DrawTuLiNote(ByVal X0,ByVal Y0,ByVal X1,ByVal Y1,ByVal Code,ByVal Color,ByVal Width,ByVal Height,ByVal FontString,ByVal PolygonID,ByVal ZtMc,ByVal Angle)
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    ResultX1 = X0 + (X1 - X0) * COSVal - (Y1 - Y0) * SinVal
    ResultY1 = Y0 + (X1 - X0) * SinVal + (Y1 - Y0) * COSVal
    
    SSProcess.CreateNewObj 3

    Angle = SSProcess.RadianToDeg(Angle)

    SSProcess.SetNewObjValue "SSObj_FontWordAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontStringAngle", Angle
    SSProcess.SetNewObjValue "SSObj_FontClass", Code
    SSProcess.SetNewObjValue "SSObj_FontString", FontString
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_DataMark", PolygonID
    SSProcess.SetNewObjValue "SSObj_FontName", ZtMc
    SSProcess.SetNewObjValue "SSObj_LayerName", "面积注记"
    SSProcess.SetNewObjValue "SSObj_FontAlignment", "0"
    SSProcess.SetNewObjValue "SSObj_FontWidth",Width
    SSProcess.SetNewObjValue "SSObj_FontHeight", Height
    SSProcess.AddNewObjPoint ResultX1,ResultY1,0,0,""

    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase
    
End Function' DrawTuLiNote


'图斑点旋转(对应MakePointtb)
Function DrawTbPoint(ByVal X0,ByVal Y0,ByVal X1,ByVal Y1,ByVal Code,ByVal Color,ByVal PolygonID,ByVal JzDh,ByVal Angle)
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    ResultX1 = X0 + (X1 - X0) * COSVal - (Y1 - Y0) * SinVal
    ResultY1 = Y0 + (X1 - X0) * SinVal + (Y1 - Y0) * COSVal

    SSProcess.CreateNewObj 0

    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_DataMark", PolygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "勘测定界图廓"
    SSProcess.SetNewObjValue "[jzdh]",jzdh
    SSProcess.AddNewObjPoint ResultX1, ResultY1, 0, 0, ""

    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase

End Function ' DrawTbPoint

'村界线旋转(对应MakeLinecj)
Function DrawLine_Cj(ByVal X0,ByVal Y0,ByVal X1,ByVal Y1,ByVal X2,ByVal Y2,ByVal Code,ByVal Color,ByVal PolygonID,ByVal Angle)

    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    X3 = X0 + (X1 - X0) * COSVal - (Y1 - Y0) * SinVal
    Y3 = Y0 + (X1 - X0) * SinVal + (Y1 - Y0) * COSVal
    
    X4 = X0 + (X2 - X0) * COSVal - (Y2 - Y0) * SinVal
    Y4 = Y0 + (X2 - X0) * SinVal + (Y2 - Y0) * COSVal

    SSProcess.CreateNewObj 1
    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_DataMark", PolygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "勘测村界图例"
    SSProcess.AddNewObjPoint X3, Y3, 0, 0, ""
    SSProcess.AddNewObjPoint X4, Y4, 0, 0, ""

    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase

End Function ' DrawLine_Cj

'地类图斑面旋转(对应MakeAreatb)
Function DrawArea_Tb(ByVal X0,ByVal Y0,ByVal X1,ByVal Y1,ByVal X2,ByVal Y2,ByVal X3,ByVal Y3,ByVal X4,ByVal Y4,ByVal Code,ByVal Color,ByVal PolygonID,ByVal tbdlbm,ByVal tbdlmc,ByVal Angle)
    
    SinVal = Sin(Angle)
    COSVal = Cos(Angle)
    
    ResultX1 = X0 + (X1 - X0) * COSVal - (Y1 - Y0) * SinVal
    ResultY1 = Y0 + (X1 - X0) * SinVal + (Y1 - Y0) * COSVal
    
    ResultX2 = X0 + (X2 - X0) * COSVal - (Y2 - Y0) * SinVal
    ResultY2 = Y0 + (X2 - X0) * SinVal + (Y2 - Y0) * COSVal
    
    ResultX3 = X0 + (X3 - X0) * COSVal - (Y3 - Y0) * SinVal
    ResultY3 = Y0 + (X3 - X0) * SinVal + (Y3 - Y0) * COSVal
    
    ResultX4 = X0 + (X4 - X0) * COSVal - (Y4 - Y0) * SinVal
    ResultY4 = Y0 + (X4 - X0) * SinVal + (Y4 - Y0) * COSVal
    
    SSProcess.CreateNewObj 2

    SSProcess.SetNewObjValue "SSObj_Code", Code
    SSProcess.SetNewObjValue "SSObj_Color", Color
    SSProcess.SetNewObjValue "SSObj_DataMark", PolygonID
    SSProcess.SetNewObjValue "SSObj_LayerName", "地类图斑图例"
    SSProcess.SetNewObjValue "[dlbm]",tbdlbm
    SSProcess.SetNewObjValue "[dlmc]",tbdlmc

    SSProcess.AddNewObjPoint ResultX1, ResultY1, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX2, ResultY2, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX3, ResultY3, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX4, ResultY4, 0, 0, ""
    SSProcess.AddNewObjPoint ResultX1, ResultY1, 0, 0, ""

    SSProcess.AddNewObjToSaveObjList
    SSProcess.SaveBufferObjToDatabase

End Function ' DrawArea_Tb

