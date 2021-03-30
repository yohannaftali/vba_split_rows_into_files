Sub SplitRowsIntoFiles()
'
' SplitRowsIntoFiles Macro
' Split Rows Into Files by Yohan Naftali
'
' Keyboard Shortcut: Ctrl+s
'
Dim dataRange As Range
Dim cuttedRange As Range
Dim rowPerFile As Integer
Dim headerRange As Range
Dim MainWS As Worksheet

On Error Resume Next

Set headerRange = Application.InputBox("Select Header Range", Header, "$1:$1", Type:=8)

Set dataRange = Application.Selection
Set dataRange = Application.InputBox("Select Data Range", “Data”, dataRange.Address, Type:=8)

rowPerFile = Application.InputBox("Row Per File", idString, 30, Type:=1)

Debug.Print "After select Data Range"
Set MainWS = dataRange.Parent
Set cuttedRange = dataRange.Rows(1)

Application.ScreenUpdating = False

For i = 1 To dataRange.Rows.Count Step rowPerFile
    Debug.Print "i: " & i
    resizeCount = rowPerFile
    If (dataRange.Rows.Count) < rowPerFile Then resizeCount = dataRange.Rows.Count
    cuttedRange.Resize(resizeCount).Copy
    Debug.Print "count: " & resizeCount
    Application.Worksheets.Add after:=Application.Worksheets(Application.Worksheets.Count)
    Application.ActiveSheet.Range("A1").PasteSpecial
    Set cuttedRange = cuttedRange.Offset(rowPerFile)
Next

Dim WS As Worksheet
For Each WS In Worksheets
    If WS.Name <> MainWS.Name Then
        MainWS.Rows(headerRange.Address).Copy
        WS.Rows("1:1").Insert Shift:=xlDown
    End If
Next WS
Application.CutCopyMode = False
Application.ScreenUpdating = True

mainWorkbookName = ActiveWorkbook.Name
mainWorkbookFullName = ActiveWorkbook.FullName
mainWorkbookPath = ActiveWorkbook.Path

For Each WS In Worksheets
    If WS.Name <> MainWS.Name Then
        Application.DisplayAlerts = False
        Name = Left(mainWorkbookName, (InStrRev(mainWorkbookName, ".", -1, vbTextCompare) - 1)) & "_" & WS.Name
        WS.Select
        WS.Move
        Set wb = ActiveWorkbook
        With wb
            .SaveAs FileName:=mainWorkbookPath & Application.PathSeparator & Name & ".xlsx", FileFormat:=xlOpenXMLWorkbook
            With .Sheets(1)
                .Paste
                .Name = SheetName
            End With
            .Close False
        End With
        Set wb = Nothing
        Application.DisplayAlerts = True
        Workbooks(mainWorkbookFullName).Activate
    End If
Next WS
    
Application.CutCopyMode = False
Application.ScreenUpdating = True

End Sub
