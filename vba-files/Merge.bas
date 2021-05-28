Attribute VB_Name = "Merge"

Sub MergeMultipleDataIntoMasterWb()
    Dim Filepath As String
    Dim Folderpath As String
    Dim FileName As String
    Dim erow As Long
    
    Folderpath = ActiveWorkbook.Path & "\SPX_Data\"
    Filepath = Folderpath & "*.xls*"
    FileName = Dir(Filepath)
    
    Dim LastRow As Long, LastColumn As Long
    Do While FileName <> ""
    Workbooks.Open (Folderpath & FileName)
    
    LastRow = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
    LastColumn = ActiveSheet.Cells(1, Columns.count).End(xlToLeft).Column
    Range(Cells(2, 1), Cells(LastRow, LastColumn)).Copy
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Close True
    
    
    erow = Sheet1.Cells(Rows.count, 1).End(xlUp).Offset(1, 0).Row
    ActiveSheet.Paste Destination:=Worksheets("Sheet1").Range(Cells(erow, 1), Cells(erow, 14))
    
    FileName = Dir
    
    Loop
    Application.DisplayAlerts = True
    
End Sub