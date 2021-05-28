Attribute VB_Name = "LoopDelete"

'''vba macros to loop files over a specific folder and loop over worksheet to delete specific words

Sub LoopAllExcelFileFolders()
    Dim ws As Worksheet
    Dim FinalRow, i As Long
    FinalRow = Cells(Rows.count, 1).End(xlUp).Row
    Dim wb As Workbook
    Dim myPath As String
    Dim myFile As String
    Dim myExtension As String
    Dim FldrPicker As FileDialog
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationManual
    
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
        With FldrPicker
          .Title = "D:\vba-course\test"
          .AllowMultiSelect = False
            If .Show <> -1 Then GoTo NextCode
            myPath = .SelectedItems(1) & "\"
        End With
NextCode:
    myPath = myPath
    If myPath = "" Then GoTo ResetSettings
    myExtension = "*.xls*"

    myFile = Dir(myPath & myExtension)

    Do While myFile <> ""
        Set wb = Workbooks.Open(FileName:=myPath & myFile)
        DoEvents
        For Each ws In Worksheets
            For i = FinalRow To 1 Step -1
                If ws.Cells(i, 1).Value = "sold" Or _
                ws.Cells(i, 1).Value = "crm" Then
                    ws.Cells(i, 1).ClearContents
                End If
            Next i
        Next ws
        wb.Close SaveChanges:=True
        DoEvents
        myFile = Dir
    Loop
    MsgBox "Task Complete!"
    
ResetSettings:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub