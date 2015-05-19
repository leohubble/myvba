Attribute VB_Name = "模块1"

Sub 宏1()
Attribute 宏1.VB_ProcData.VB_Invoke_Func = " \n14"


'Dim i As Integer
'
'  For i = 2 To 45
'
'  Workbooks("联系方式").Worksheets("address").Cells(i, 2) = Workbooks("联系方式").Worksheets("联系方式").Cells(1 + 4 * (i - 2), 2) 'B1 公司名称
'  Workbooks("联系方式").Worksheets("address").Cells(i, 3) = Workbooks("联系方式").Worksheets("联系方式").Cells(1 + 4 * (i - 2), 6) 'F1 所属省市
'  Workbooks("联系方式").Worksheets("address").Cells(i, 4) = Workbooks("联系方式").Worksheets("联系方式").Cells(2 + 4 * (i - 2), 2) 'B2 通信地址
'  Workbooks("联系方式").Worksheets("address").Cells(i, 5) = Workbooks("联系方式").Worksheets("联系方式").Cells(2 + 4 * (i - 2), 7) 'G2 邮编
'  Workbooks("联系方式").Worksheets("address").Cells(i, 6) = Workbooks("联系方式").Worksheets("联系方式").Cells(3 + 4 * (i - 2), 2) 'B3 联系人
'  Workbooks("联系方式").Worksheets("address").Cells(i, 7) = Workbooks("联系方式").Worksheets("联系方式").Cells(3 + 4 * (i - 2), 7) 'G3 手机
'  Workbooks("联系方式").Worksheets("address").Cells(i, 8) = Workbooks("联系方式").Worksheets("联系方式").Cells(3 + 4 * (i - 2), 4) 'D3 联系电话
'  Workbooks("联系方式").Worksheets("address").Cells(i, 9) = Workbooks("联系方式").Worksheets("联系方式").Cells(4 + 4 * (i - 2), 2) 'B4 传真
'  Workbooks("联系方式").Worksheets("address").Cells(i, 10) = Workbooks("联系方式").Worksheets("联系方式").Cells(4 + 4 * (i - 2), 4) 'D4 email
'
'  Next
 Dim objWdApp As Object
 Dim path As String
 Dim file As String
 Dim txt As String
 Dim count As Integer
 
 For count = 1 To 1

 'file = "dir\filename.doc"
 file = Workbooks("联系方式").Worksheets("sheet1").Cells(count, 4)
 path = "E:\4\_success\" & file
    Set objWdApp = CreateObject("word.application")
    objWdApp.Documents.Open (path)

   txt = objWdApp.Documents(1).Tables(1).Cell(1, 2).Range.Text   'B1 公司名称
   Workbooks("联系方式").Worksheets("address").Cells(count, 2) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(1, 4).Range.Text
   Workbooks("联系方式").Worksheets("address").Cells(count, 3) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(2, 2).Range.Text
   Workbooks("联系方式").Worksheets("address").Cells(count, 4) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(2, 4).Range.Text
   Workbooks("联系方式").Worksheets("address").Cells(count, 5) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(3, 2).Range.Text
   Workbooks("联系方式").Worksheets("address").Cells(count, 6) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(3, 4).Range.Text
   Workbooks("联系方式").Worksheets("address").Cells(count, 7) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(3, 6).Range.Text
   Workbooks("联系方式").Worksheets("address").Cells(count, 8) = txt

  txt = objWdApp.Documents(1).Tables(1).Cell(4, 2).Range.Text
   Workbooks("联系方式").Worksheets("address").Cells(count, 9) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(4, 4).Range.Text
   Workbooks("联系方式").Worksheets("address").Cells(count, 10) = txt

   objWdApp.Documents.Close
   objWdApp.Quit
  Next
End Sub

