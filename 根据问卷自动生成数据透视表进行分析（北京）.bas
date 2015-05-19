Attribute VB_Name = "模块2"
Sub 根据问卷生成数据透视表()

Dim destws As Worksheet '报告结果表
Dim sourcews As Worksheet '数据来源表

Dim i, j, k, flag, count, wholecol As Integer
Dim txt1, txt2 As String
Dim tbdest As String '数据透视表位置
Dim tbname As String '数据透视表名称
Dim inipos As Integer '每个问题的初始位置
Dim datasource As String '数据来源

Dim ptfieldname As String '字段名称
Dim ptcaption As String '字段标题
Dim ptcount, ptfiled As Integer ' 每个问题的选项数目

Dim pt As Integer ' 问题数目
Dim pvc As PivotCache '问卷数据cache
Dim pvt As PivotTable '问卷数据表


Dim rowcount, colcount As Integer '绘制数据透视表位置的行列
Dim frowcount, fcolcount As Integer ' 计算比例的公式的位置
Dim chtrowcount, chtcolcount As Integer '绘图数据源地址
Dim strformula As String ' 计算公式
Dim straddress As String '数据透视表地址

Dim chtpos, chtpos2 As Range '绘图数据地址
Dim chtdata As Range ' 绘图的数据源
Dim chtxvalue As String '坐标轴标注
Dim chtchart As Chart '生成的图
Dim drawchart As Boolean '是否绘图


datasource = "原始数据!$A$1:$VW$1715"

wholecol = Range(datasource).Columns.count

Set destws = ThisWorkbook.Sheets.Add 'Sheet4
destws.Name = "问卷数据分析结果"

Application.ScreenUpdating = False

drawchart = True



Set ptc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=datasource, Version:=xlPivotTableVersion15)


i = 9
j = 9
flag = 0
While i < wholecol + 1
count = 1
txt1 = Sheet2.Cells(1, i)
txt2 = Sheet2.Cells(1, i + 1)

inipos = i '记录每个问题开始的列数

While txt1 = txt2
txt1 = Sheet2.Cells(1, i)
txt2 = Sheet2.Cells(1, i + 1)
i = i + 1
count = count + 1
flag = 1
Wend

If flag = 0 Then

    ptfiled = count '第pt号问题有ptfiled个选项
            i = i + 1
                j = j + 1
Else

        ptfiled = count - 1 '第pt号问题有ptfiled个选项
                j = j + 1

End If

flag = 0




''''''''''''''''''''''''''''''''''''''''''''''
'算出每个问题的选项数目
'
''''''''''''''''''''''''''''''''''''''''''''''

pt = j '第pt个问题

tbname = txt1 + CStr(pt) '数据透视表名称

'数据透视表初始位置

rowcount = 10 + (pt - 9) * 50
colcount = 1
destws.Cells(rowcount - 1, 5) = tbname
straddress = destws.Cells(rowcount, colcount).Address

'tbdest = "Sheet3!R" + CStr(rowcount) + "C" + CStr(colcount) '构造透视表左上角单元位置
tbdest = destws.Name + "!R" + CStr(rowcount) + "C" + CStr(colcount)

straddress = destws.Cells(rowcount, colcount).Address


'生成数据透视表

Set pvt = ptc.CreatePivotTable(TableDestination:=tbdest, TableName:=tbname, DefaultVersion:=xlPivotTableVersion15)

        With pvt
            '加入筛选项

            '.PivotFields("行业类型_1").Orientation = xlPageField
            '.PivotFields("行业类型_1").Position = 1
     
            .PivotFields("重点产业").Orientation = xlPageField
            .PivotFields("重点产业").Position = 1

            
          ' .PivotFields("行业类型_2").Orientation = xlPageField
          '  .PivotFields("行业类型_2").Position = 2
            
            .PivotFields("区县").Orientation = xlPageField
            .PivotFields("区县").Position = 2
            
'            .PivotFields("所属行业").Orientation = xlPageField
'           .PivotFields("所属行业").Position = 4
'
'            .PivotFields("企业类型").Orientation = xlPageField
'            .PivotFields("企业类型").Position = 5
            '统计企业总数
             .AddDataField pvt.PivotFields("单位名称"), "企业数量", xlCount
        
                '自动加入统计字段
                For k = 1 To ptfiled  ' 加入问题的每个选项
                    ptfiledname = Sheet2.Cells(1, inipos + k - 1)
                    ptcaption = Sheet2.Cells(2, inipos + k - 1)
                    '
                    If k <> 1 Then
                    ptfiledname = ptfiledname + CStr(k)
                    End If
                    frowcount = rowcount + 2
                    fcolcount = k + 1
                      If ptfiled > 2 Then ' 如果是多个选项，计算出相对于总数的比例
                         .AddDataField pvt.PivotFields(ptfiledname), ptcaption, xlSum
                         
                         strformula = "=GETPIVOTDATA(" + Chr(34) + ptcaption + Chr(34) + "," + straddress + ")" + "/GETPIVOTDATA(" + Chr(34) + "企业数量" + Chr(34) + "," + straddress + ")"  '计算比例
                         With Range(destws.Cells(frowcount, fcolcount).Address)
                         .Formula = strformula ' 自动计算比例
                   
                         .Style = "Percent" '设置为百分比
                         .NumberFormatLocal = "0.0%"
                  
                         End With
                         Else
                        .AddDataField pvt.PivotFields(ptfiledname), ptcaption, xlSum
                         strformula = "=" + Cells(frowcount - 1, fcolcount).Address '"=GETPIVOTDATA(" + Chr(34) + ptcaption + Chr(34) + "," + straddress + ")"
                          With Range(destws.Cells(frowcount, fcolcount).Address)
                            .Formula = strformula ' 自动计算比例
                          End With
                       End If
                Next
                
        '绘制统计图
        If drawchart = True Then
        
              If ptfiled > 2 Then '选项超过2个的，画雷达图
                   
                       
                Set chtdata = destws.Range(destws.Cells(frowcount, 2), destws.Cells(frowcount, 1 + ptfiled))
                Set chtpos = destws.Range(Cells(frowcount + 2, 1).Address) '确定绘图位置


                destws.Shapes.AddChart2(317, xlRadarMarkers, chtpos.Left, chtpos.Top).Select '生成雷达图
                chtxvalue = "=" + destws.Name + "!" + destws.Cells(rowcount, 2).Address + ":" + destws.Cells(rowcount, 1 + ptfiled).Address '生成坐标轴位置单元格区域地址
                
                ActiveChart.SetSourceData Source:=chtdata
                ActiveChart.HasLegend = False
                ActiveChart.HasTitle = False

               ActiveChart.Axes(xlValue).Select
                Selection.Delete

                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.SetElement (msoElementDataLabelCallout)
                ActiveChart.SetElement (msoElementDataLabelNone)
                ActiveChart.SetElement (msoElementDataLabelShow)
                ActiveChart.FullSeriesCollection(1).XValues = chtxvalue '坐标轴显示内容
                
                
                Set chtdata = destws.Range(destws.Cells(frowcount, 2), destws.Cells(frowcount, 1 + ptfiled))
                Set chtpos = destws.Range(Cells(frowcount + 2, 10).Address) '确定绘图位置
   
                destws.Shapes.AddChart2(216, xlBarClustered, chtpos.Left, chtpos.Top).Select '生成雷达图
                chtxvalue = "=" + destws.Name + "!" + destws.Cells(rowcount, 2).Address + ":" + destws.Cells(rowcount, 1 + ptfiled).Address '生成坐标轴位置单元格区域地址
               
                
                ActiveChart.SetSourceData Source:=chtdata
                ActiveChart.HasLegend = False
                ActiveChart.HasTitle = False

               ActiveChart.Axes(xlValue).Select
                Selection.Delete
               ActiveChart.Axes(xlValue).MajorGridlines.Select
                Selection.Delete
                
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.SetElement (msoElementDataLabelCallout)
                ActiveChart.SetElement (msoElementDataLabelNone)
                ActiveChart.SetElement (msoElementDataLabelShow)
                ActiveChart.FullSeriesCollection(1).XValues = chtxvalue '坐标轴显示内容
   

                End If
                
                If ptfiled = 2 Then '只有2个选项的问题，画横道图
                   
                Set chtdata = destws.Range(destws.Cells(frowcount, 2), destws.Cells(frowcount, 1 + ptfiled))
                Set chtpos = destws.Range(Cells(frowcount + 2, 1).Address) '确定绘图位置
 
                destws.Shapes.AddChart2(216, xlBarClustered, chtpos.Left, chtpos.Top).Select '生成雷达图
                chtxvalue = "=" + destws.Name + "!" + destws.Cells(rowcount, 2).Address + ":" + destws.Cells(rowcount, 1 + ptfiled).Address '生成坐标轴位置单元格区域地址
               
                
                ActiveChart.SetSourceData Source:=chtdata
                ActiveChart.HasLegend = False
                ActiveChart.HasTitle = False

               ActiveChart.Axes(xlValue).Select
                Selection.Delete
               ActiveChart.Axes(xlValue).MajorGridlines.Select
                Selection.Delete
                
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.SetElement (msoElementDataLabelCallout)
                ActiveChart.SetElement (msoElementDataLabelNone)
                ActiveChart.SetElement (msoElementDataLabelShow)
                ActiveChart.FullSeriesCollection(1).XValues = chtxvalue '坐标轴显示内容
                End If
            End If
        
        End With

Wend
Application.ScreenUpdating = True

End Sub
