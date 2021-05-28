Attribute VB_Name = "NewSheet"

'''This macro is to add new sheet next to the active sheet with a specific name

'''1.
Sub ToAddNewSheet1()

Sheets.Add.Name = Range("C2")
End Sub

'''2.
Sub ToAddNewSheet2()

Sheets.Add After:=Sheets(Sheets.count)
End Sub

'''3.
Sub ToAddNewSheet3()
Dim Sheet_Count As Integer
Sheet_Count = Range("A2:A5").Rows.count

Dim Sheet_Name As String
Dim i As Integer
For i = 1 To Sheet_Count
    Sheet_Name = Worksheets("Option").Range("A2:A5").Cells(i, 1).Value
    Sheets.Add.Name = Sheet_Name
    
Next i

End Sub


''' Add new sheet buton

Sub Add_New_Sheet_Button()
    'Dim a As String
    'a = ActiveSheet.Cells(1, 1).Value
    'Sheets.Add After:=Sheets(Sheets.Count)
    'ActiveSheet.Name = a
    'Sheets(1).Activate
    'ActiveSheet.Cells(1, 1).Select
    'ActiveSheet.Cells(1, 1).Value = "NASDAQ"
    
Dim sheet_to_create As String
Dim i As Long
sheet_to_create = ActiveSheet.Cells(1, 1).Value
For i = 1 To (Worksheets.count)
    If LCase(Sheets(i).Name) = LCase(sheet_to_create) Then
    MsgBox "This Sheet already exist!"
    Exit Sub
    End If
Next
    
    Sheets.Add After:=Sheets(Sheets.count)
    Sheets(ActiveSheet.Name).Name = sheet_to_create
End Sub

'''

Sub Add_Sheet_With_Name()
    Sheets.Add After:=Sheets("Option")
    'Sheets.Add before:=Sheets("")
    Sheets.Add.Name = ("List_ref_categories")
End Sub
