Attribute VB_Name = "Openwb"


Sub OpenWorkbook() 'To open a new workbook in a given location
    Dim Filepath As String
    Filepath = ActiveWorkbook.Path
    
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filepath & "\Final_Sample_MSc_Thesis.xlsx")
    
    'MsgBox "You've just opened " & wb.Name
    'MsgBox "The file  location of what you have just opened is " & wb.Path
    
End Sub