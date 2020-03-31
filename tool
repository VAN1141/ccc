Sub test()
Dim a As String
Dim i As Long
Dim b As String
Dim c As String
Dim d As String

i = Sheets(1).Range("R6").Value
a = CStr(Sheets(2).Cells(i, 1).Value)
b = CStr(Sheets(2).Cells(i, 2).Value)

    Sheets(1).TextBox1 = a

If Sheets(1).TextBox2 = b Then
    Sheets(1).Range("R6").Value = Sheets(1).Range("R6").Value + 1
    i = Sheets(1).Range("R6").Value
a = CStr(Sheets(2).Cells(i, 1).Value)
b = CStr(Sheets(2).Cells(i, 2).Value)
c = CStr(Sheets(2).Cells(i - 1, 3).Value)
d = CStr(Sheets(2).Cells(i - 1, 4).Value)

Sheets(1).Select
Sheets("Sheet1").TextBox1 = a

Else: MsgBox ("sai, nhap lai")
'b = b + 1
'b = Range("A2").Value
' a = Range("A1").Value
'ThisWorkbook.Sheets("Sheet1").TextBox1 = a
'Else
'MsgBox "sai"
End If
Sheets(1).TextBox3 = c
Sheets(1).TextBox4 = d
Sheets(1).TextBox2 = ""
End Sub
