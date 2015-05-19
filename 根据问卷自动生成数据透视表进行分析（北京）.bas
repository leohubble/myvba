Attribute VB_Name = "ģ��2"
Sub �����ʾ���������͸�ӱ�()

Dim destws As Worksheet '��������
Dim sourcews As Worksheet '������Դ��

Dim i, j, k, flag, count, wholecol As Integer
Dim txt1, txt2 As String
Dim tbdest As String '����͸�ӱ�λ��
Dim tbname As String '����͸�ӱ�����
Dim inipos As Integer 'ÿ������ĳ�ʼλ��
Dim datasource As String '������Դ

Dim ptfieldname As String '�ֶ�����
Dim ptcaption As String '�ֶα���
Dim ptcount, ptfiled As Integer ' ÿ�������ѡ����Ŀ

Dim pt As Integer ' ������Ŀ
Dim pvc As PivotCache '�ʾ�����cache
Dim pvt As PivotTable '�ʾ����ݱ�


Dim rowcount, colcount As Integer '��������͸�ӱ�λ�õ�����
Dim frowcount, fcolcount As Integer ' ��������Ĺ�ʽ��λ��
Dim chtrowcount, chtcolcount As Integer '��ͼ����Դ��ַ
Dim strformula As String ' ���㹫ʽ
Dim straddress As String '����͸�ӱ��ַ

Dim chtpos, chtpos2 As Range '��ͼ���ݵ�ַ
Dim chtdata As Range ' ��ͼ������Դ
Dim chtxvalue As String '�������ע
Dim chtchart As Chart '���ɵ�ͼ
Dim drawchart As Boolean '�Ƿ��ͼ


datasource = "ԭʼ����!$A$1:$VW$1715"

wholecol = Range(datasource).Columns.count

Set destws = ThisWorkbook.Sheets.Add 'Sheet4
destws.Name = "�ʾ����ݷ������"

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

inipos = i '��¼ÿ�����⿪ʼ������

While txt1 = txt2
txt1 = Sheet2.Cells(1, i)
txt2 = Sheet2.Cells(1, i + 1)
i = i + 1
count = count + 1
flag = 1
Wend

If flag = 0 Then

    ptfiled = count '��pt��������ptfiled��ѡ��
            i = i + 1
                j = j + 1
Else

        ptfiled = count - 1 '��pt��������ptfiled��ѡ��
                j = j + 1

End If

flag = 0




''''''''''''''''''''''''''''''''''''''''''''''
'���ÿ�������ѡ����Ŀ
'
''''''''''''''''''''''''''''''''''''''''''''''

pt = j '��pt������

tbname = txt1 + CStr(pt) '����͸�ӱ�����

'����͸�ӱ��ʼλ��

rowcount = 10 + (pt - 9) * 50
colcount = 1
destws.Cells(rowcount - 1, 5) = tbname
straddress = destws.Cells(rowcount, colcount).Address

'tbdest = "Sheet3!R" + CStr(rowcount) + "C" + CStr(colcount) '����͸�ӱ����Ͻǵ�Ԫλ��
tbdest = destws.Name + "!R" + CStr(rowcount) + "C" + CStr(colcount)

straddress = destws.Cells(rowcount, colcount).Address


'��������͸�ӱ�

Set pvt = ptc.CreatePivotTable(TableDestination:=tbdest, TableName:=tbname, DefaultVersion:=xlPivotTableVersion15)

        With pvt
            '����ɸѡ��

            '.PivotFields("��ҵ����_1").Orientation = xlPageField
            '.PivotFields("��ҵ����_1").Position = 1
     
            .PivotFields("�ص��ҵ").Orientation = xlPageField
            .PivotFields("�ص��ҵ").Position = 1

            
          ' .PivotFields("��ҵ����_2").Orientation = xlPageField
          '  .PivotFields("��ҵ����_2").Position = 2
            
            .PivotFields("����").Orientation = xlPageField
            .PivotFields("����").Position = 2
            
'            .PivotFields("������ҵ").Orientation = xlPageField
'           .PivotFields("������ҵ").Position = 4
'
'            .PivotFields("��ҵ����").Orientation = xlPageField
'            .PivotFields("��ҵ����").Position = 5
            'ͳ����ҵ����
             .AddDataField pvt.PivotFields("��λ����"), "��ҵ����", xlCount
        
                '�Զ�����ͳ���ֶ�
                For k = 1 To ptfiled  ' ���������ÿ��ѡ��
                    ptfiledname = Sheet2.Cells(1, inipos + k - 1)
                    ptcaption = Sheet2.Cells(2, inipos + k - 1)
                    '
                    If k <> 1 Then
                    ptfiledname = ptfiledname + CStr(k)
                    End If
                    frowcount = rowcount + 2
                    fcolcount = k + 1
                      If ptfiled > 2 Then ' ����Ƕ��ѡ����������������ı���
                         .AddDataField pvt.PivotFields(ptfiledname), ptcaption, xlSum
                         
                         strformula = "=GETPIVOTDATA(" + Chr(34) + ptcaption + Chr(34) + "," + straddress + ")" + "/GETPIVOTDATA(" + Chr(34) + "��ҵ����" + Chr(34) + "," + straddress + ")"  '�������
                         With Range(destws.Cells(frowcount, fcolcount).Address)
                         .Formula = strformula ' �Զ��������
                   
                         .Style = "Percent" '����Ϊ�ٷֱ�
                         .NumberFormatLocal = "0.0%"
                  
                         End With
                         Else
                        .AddDataField pvt.PivotFields(ptfiledname), ptcaption, xlSum
                         strformula = "=" + Cells(frowcount - 1, fcolcount).Address '"=GETPIVOTDATA(" + Chr(34) + ptcaption + Chr(34) + "," + straddress + ")"
                          With Range(destws.Cells(frowcount, fcolcount).Address)
                            .Formula = strformula ' �Զ��������
                          End With
                       End If
                Next
                
        '����ͳ��ͼ
        If drawchart = True Then
        
              If ptfiled > 2 Then 'ѡ���2���ģ����״�ͼ
                   
                       
                Set chtdata = destws.Range(destws.Cells(frowcount, 2), destws.Cells(frowcount, 1 + ptfiled))
                Set chtpos = destws.Range(Cells(frowcount + 2, 1).Address) 'ȷ����ͼλ��


                destws.Shapes.AddChart2(317, xlRadarMarkers, chtpos.Left, chtpos.Top).Select '�����״�ͼ
                chtxvalue = "=" + destws.Name + "!" + destws.Cells(rowcount, 2).Address + ":" + destws.Cells(rowcount, 1 + ptfiled).Address '����������λ�õ�Ԫ�������ַ
                
                ActiveChart.SetSourceData Source:=chtdata
                ActiveChart.HasLegend = False
                ActiveChart.HasTitle = False

               ActiveChart.Axes(xlValue).Select
                Selection.Delete

                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.SetElement (msoElementDataLabelCallout)
                ActiveChart.SetElement (msoElementDataLabelNone)
                ActiveChart.SetElement (msoElementDataLabelShow)
                ActiveChart.FullSeriesCollection(1).XValues = chtxvalue '��������ʾ����
                
                
                Set chtdata = destws.Range(destws.Cells(frowcount, 2), destws.Cells(frowcount, 1 + ptfiled))
                Set chtpos = destws.Range(Cells(frowcount + 2, 10).Address) 'ȷ����ͼλ��
   
                destws.Shapes.AddChart2(216, xlBarClustered, chtpos.Left, chtpos.Top).Select '�����״�ͼ
                chtxvalue = "=" + destws.Name + "!" + destws.Cells(rowcount, 2).Address + ":" + destws.Cells(rowcount, 1 + ptfiled).Address '����������λ�õ�Ԫ�������ַ
               
                
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
                ActiveChart.FullSeriesCollection(1).XValues = chtxvalue '��������ʾ����
   

                End If
                
                If ptfiled = 2 Then 'ֻ��2��ѡ������⣬�����ͼ
                   
                Set chtdata = destws.Range(destws.Cells(frowcount, 2), destws.Cells(frowcount, 1 + ptfiled))
                Set chtpos = destws.Range(Cells(frowcount + 2, 1).Address) 'ȷ����ͼλ��
 
                destws.Shapes.AddChart2(216, xlBarClustered, chtpos.Left, chtpos.Top).Select '�����״�ͼ
                chtxvalue = "=" + destws.Name + "!" + destws.Cells(rowcount, 2).Address + ":" + destws.Cells(rowcount, 1 + ptfiled).Address '����������λ�õ�Ԫ�������ַ
               
                
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
                ActiveChart.FullSeriesCollection(1).XValues = chtxvalue '��������ʾ����
                End If
            End If
        
        End With

Wend
Application.ScreenUpdating = True

End Sub
