Dim strSheet As String
Dim strSpreadsheetPath As String
Dim BoolMatchCase As Boolean
Dim intChangesBefore As Integer
Dim intChangesAfter As Integer
Dim intChangesNet As Integer


Sub checkVariantBrands()
'specify worksheet name
strSheet = "Brands"
BoolMatchCase = True
'run the Find/Replace routine
CheckVariantSpelling
End Sub

Sub checkVariantUK()
'specify worksheet name
strSheet = "UKoed"
BoolMatchCase = False
'run the Find/Replace routine
CheckVariantSpelling
End Sub


Function CheckVariantSpelling()
    
    'set up the timer
    Dim TimeStart As Single, TimeEnd As Single
    TimeStart = Timer
        
    
    'how many tracked changes in the document before we start?
    intChangesBefore = ActiveDocument.Revisions.Count
    
    
    'Spreadsheet location on the disk
    'strSpreadsheetPath = "C:\spellingvariants.xlsx"
    strSpreadsheetPath = Environ$("USERPROFILE") & "\my documents\spellingvariants.xlsx"
    
    'get user confirmation for this procedure
    Dim askConfirm As Integer
    askConfirm = MsgBox("Using " & strSpreadsheetPath & "  ", vbYesNo + vbExclamation, "Are you sure?")
    Select Case askConfirm
        Case vbNo
            Exit Function
    End Select
    
    ''check the spread sheet exists
    'If File_Exists(strSpreadsheetPath) = True Then
    ''do no thing... carry on
    'Else
    'MsgBox strSpreadsheetPath & " does not exist"
    'Exit Function
    'End If
    
    
    'turn on Track Changes?
    'Dim askTC As Integer
    'askTC = MsgBox("Do you want to turn on Track Changes (recommended!) before running this procedure?", vbYesNo + vbExclamation, "Turn on Track Changes?")
    'Select Case askTC
    'Case vbYes
    With ActiveDocument
        .TrackRevisions = True
        .ShowRevisions = True
    End With
    'End Select
    
    
    'define string for manual checking after global replace
    ' strCheck = "CHECK=>"
    strCheck = ""
    
    'zero the counter
    intCounter = 0
    
    'specify column number for FIND string
    intFind = 1
    
    'specify column number for REPLACE string
    intRepl = 2
    
    Set objExcel = CreateObject("Excel.Application")
    objExcel.DisplayAlerts = False
    
    'Excel will(True) or won't(False) be visible while reading the data
    objExcel.Visible = False
    
    Set objWorkbook = objExcel.Workbooks.Open(strSpreadsheetPath)
    
    'specify worksheet name
    Set objSheet = objWorkbook.Sheets(strSheet)
    
    'Which column to use for searching
    Set colFind = objSheet.UsedRange.Columns(intFind)
    
    'Column to use for replace
    Set colReplace = objSheet.UsedRange.Columns(intRepl)
    
    'Find/Replace the words in the active document
    Set objContent = ActiveDocument.Content
    objContent.Find.MatchWholeWord = True
    objContent.Find.MatchCase = BoolMatchCase
    
    'starting in row 2
    For i = 2 To colFind.Rows.Count
    If colFind.Cells(i, 1) > "" And colReplace.Cells(i, 1) > "" Then
    objContent.Find.Execute FindText:=colFind.Cells(i, 1), _
    ReplaceWith:=strCheck & colReplace.Cells(i, 1), Replace:=wdReplaceAll
    intCounter = intCounter + 1
    End If
    Next
    
    'How many tracked changes where made ?  each subtitution is 2 changes
    intChangesAfter = ActiveDocument.Revisions.Count
    intChangesNet = intChangesAfter - intChangesBefore
    intChangesNet = intChangesNet / 2
    
    'turn off tracking for new changes, but leave the current changes visable.
    With ActiveDocument
    .TrackRevisions = False
    .ShowRevisions = True
    End With
    
    'close out timer
    TimeEnd = Timer
    TimeTester = Round(TimeEnd - TimeStart, 0)
    
    
    MsgBox intChangesNet & " words substituted. " & " Dictionary contains: " & intCounter & " words." & vbCrLf & TimeTester & " seconds"
    
    objWorkbook.Close
    objExcel.Quit


End Function
