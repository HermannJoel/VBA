Attribute VB_Name = "Consolidation"

Sub MergeMultipleSheetsIntoOneSheet2()
    'Dim wsSrc As Worksheet   'src = Source, Dest = Destination
    'Dim wsDest As Worksheet
    'Dim rngSrc As Range
    'Dim rngDest As Range
    'Dim lngLastCol As Long, lngSrcLastRow As Long, lngDestLastRow As Long
    
    'Set wsDest = ThisWorkbook.Worksheets("SPX")
    'lngDestLastRow = LastOccupiedRowNum(wsDest)
    'lngLastCol = LastOccupiedColNum(wsDest)
    
    'Set destination range
    'Set rngDest = wsDest.Cells(lngDestLastRow + 1, 1)
    
    'Loop through all sheets
    'For Each wsSrc In ThisWorbook.Worksheets
    
        'If wsSrc.Name <> "SPX" Then
        'lngSrcLastRow = LastOccupiedRowNum(wsSrc)
        
        'With wsSrc
            'Set rngSrc = .Range(.Cells(2, 1), .Cells(lngSrcLastRow, lngLastCol))
            'rngSrc.Copy Destination:=rngDest
        'End With
        
        'lngDestLastRow = LastOccupiedRowNum(wsDest)
        'Set rngDest = wsDest.Cells(lngDestLastRow + 1, 1)
        
        'End If
        
    'Next wsSrc
       
'End Sub
Sub MergeWorkbooksIntoMasterWorksheet()
    Dim strDirFiles As String, strFile As String, Filepath As String
    Dim wbDst As Workbook, wbSrc As Workbook
    Dim wsDst As Worksheet, wsSrc As Worksheet
    Dim Idx As Long, SrclastRow As Long, SrcLastCol As Long, DstLastRow As Long, DstLastCol As Long, _
        dstFistFileRow As Long
    Dim rngSrc As Range, rngDst As Range, rngFile As Range
    Dim ColFileNames As Collection
    Set ColFileNames = New Collection
    
    strDirFiles = "D:\vba-course\MID_Data"
    Set wbDst = Workbooks.Add
    Set wsDst = wbDst.ActiveSheet
    
    'Store all FileNames in a collection
    strFile = Dir(strDirFiles & "\*.xlsx*")
    Do While Len(strFile) > 0
        ColFileNames.Add Item:=strFile
        strFile = Dir
    Loop
    'Checkpoint: make sure ColFileNames has the File Names
    'Dim varDebug As Variant
    'For Each varDebug In ColFileNames
        'Debug.Print varDebug
    'Next varDebug
    
    For Idx = 1 To ColFileNames.count
    
        'Assign File path
        Filepath = strDirFiles & "\" & ColFileNames(Idx)
        'Open the workbook and store a reference to the data sheet
        Set wbSrc = Workbooks.Open(Filepath)
        Set wsSrc = wbSrc.Worksheets(1)
        
        'Identify last row & column
        SrclastRow = LastOccupiedRowNum(wsSrc)
        SrcLastCol = LastOccupiedColNum(wsSrc)
        With wsSrc
            Set rngSrc = .Range(.Cells(1, 1), .Cells(SrclastRow, SrcLastCol))
        End With
        'Checkpoin: Check the source range
        'wsSrc.Range("A1").Select
        'rngSrc.Select
        
        'To cpy the header for the 1st Loop
        If Idx <> 1 Then
            Set rngSrc = rngSrc.Offset(1, 0).Resize(rngSrc.Rows.count - 1)
        End If
        
        'Checkpoint: Check that header row has been removed for the following loop
        'wsSrc.Range("A1").Select
        'rngSrc.Select
        
        'To copy source data to destination
        If Idx = 1 Then
            DstLastRow = 1
            Set rngDst = wsDst.Cells(1, 1)
        Else
            DstLastRow = LastOccupiedRowNum(wsDst)
            Set rngDst = wsDst.Cells(DstLastRow + 1, 1)
        End If
        rngSrc.Copy Destination:=rngDst  '<~To copy paste
        
        If Idx = 1 Then
            DstLastCol = LastOccupiedColNum(wsDst)
            wsDst.Cells(1, DstLastCol + 1) = "Source Filename"
        End If
        
        With wsDst
            dstFistFileRow = DstLastRow + 1
            
            DstLastRow = LastOccupiedRowNum(wsDst)
            DstLastCol = LastOccupiedColNum(wsDst)
            
            Set rngFile = .Range(.Cells(dstFistFileRow, DstLastCol), .Cells(DstLastRow, DstLastCol))
            
            'To check destination Range
            'wbDst.Range("A1").Select
            'rngFile.Select
            
            'To write file name in the identified range
            rngFile.Value = wbSrc.Name
        
        End With
        'To close source workbook and repeat
        wbSrc.Close SaveChanges:=False
        
    Next Idx
    
    MsgBox "Data Combined Successfully", vbInformation, "Merge Macros"
            
            
End Sub
Function LastOccupiedRowNum(Sheet As Worksheet)
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                        After:=.Range("A1"), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        End With
    Else
        lng = 1
    End If
    LastOccupiedRowNum = lng
    
End Function
Function LastOccupiedColNum(Sheet As Worksheet)
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                        After:=.Range("A1"), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        End With
    Else
        lng = 1
    End If
    LastOccupiedColNum = lng
End Function


