
Attribute VB_Name = "DeleteSheet"

Sub Delete_Sheets()
    Application.DisplayAlerts = False
    
    Sheets(18).Select
    Sheets(18).Delete
    
    Application.DisplayAlerts = True
End Sub