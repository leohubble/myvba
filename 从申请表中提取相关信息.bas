Attribute VB_Name = "ģ��1"

Sub ��1()
Attribute ��1.VB_ProcData.VB_Invoke_Func = " \n14"


'Dim i As Integer
'
'  For i = 2 To 45
'
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 2) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(1 + 4 * (i - 2), 2) 'B1 ��˾����
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 3) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(1 + 4 * (i - 2), 6) 'F1 ����ʡ��
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 4) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(2 + 4 * (i - 2), 2) 'B2 ͨ�ŵ�ַ
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 5) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(2 + 4 * (i - 2), 7) 'G2 �ʱ�
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 6) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(3 + 4 * (i - 2), 2) 'B3 ��ϵ��
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 7) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(3 + 4 * (i - 2), 7) 'G3 �ֻ�
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 8) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(3 + 4 * (i - 2), 4) 'D3 ��ϵ�绰
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 9) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(4 + 4 * (i - 2), 2) 'B4 ����
'  Workbooks("��ϵ��ʽ").Worksheets("address").Cells(i, 10) = Workbooks("��ϵ��ʽ").Worksheets("��ϵ��ʽ").Cells(4 + 4 * (i - 2), 4) 'D4 email
'
'  Next
 Dim objWdApp As Object
 Dim path As String
 Dim file As String
 Dim txt As String
 Dim count As Integer
 
 For count = 1 To 1

 'file = "dir\filename.doc"
 file = Workbooks("��ϵ��ʽ").Worksheets("sheet1").Cells(count, 4)
 path = "E:\4\_success\" & file
    Set objWdApp = CreateObject("word.application")
    objWdApp.Documents.Open (path)

   txt = objWdApp.Documents(1).Tables(1).Cell(1, 2).Range.Text   'B1 ��˾����
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 2) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(1, 4).Range.Text
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 3) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(2, 2).Range.Text
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 4) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(2, 4).Range.Text
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 5) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(3, 2).Range.Text
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 6) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(3, 4).Range.Text
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 7) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(3, 6).Range.Text
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 8) = txt

  txt = objWdApp.Documents(1).Tables(1).Cell(4, 2).Range.Text
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 9) = txt

   txt = objWdApp.Documents(1).Tables(1).Cell(4, 4).Range.Text
   Workbooks("��ϵ��ʽ").Worksheets("address").Cells(count, 10) = txt

   objWdApp.Documents.Close
   objWdApp.Quit
  Next
End Sub

