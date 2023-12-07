vba_code = """

Dim OriginalValues As New Collection ' Store the original values
Dim ShouldRunWorkbookOpen As Boolean ' Flag to control whether Workbook_Open should run
Dim WorkbookOpenHasRun As Boolean ' Flag to track if Workbook_Open has already run
Dim LogData As Collection ' Store the changes to be logged


Private Sub Workbook_Open()
    Set OriginalValues = InitializeOriginalValues
    WorkbookOpenHasRun = True
    Set LogData = New Collection ' Initialize the collection for logging data
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If WorkbookOpenHasRun Then
        CheckAndLogChanges ' Call the CheckAndLogChanges subroutine before saving
        If Not SaveAsUI Then
            CheckCopyPath
        End If
    End If
End Sub

Function InitializeOriginalValues()
    Dim OriginalValues As New Collection ' Store the original values
    

    ' Store the original values when called
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Sheets
        If SheetIsInRange(Sh.Name) Then
            Dim cell As Range
            For Each cell In Sh.UsedRange
                Dim CellKey As String
                CellKey = Sh.Name & "_" & cell.Address
                ' Check if the cell is not empty before adding it to the collection
                If Not IsEmpty(cell.value) Then
                    ' Store both key and value as a single string
                    OriginalValues.Add cell.value, CellKey
                End If
            Next cell
        End If
    Next Sh

    Set InitializeOriginalValues = OriginalValues
End Function

Sub CheckCopyPath()
    ' Show "Bitte warten" message in the status bar
    Application.StatusBar = "Bitte warten...
    ' Check if cell A1 in "Bereichs Informationen" is empty
    Dim wsMA As Worksheet
    Set wsMA = ThisWorkbook.Sheets("Bereichs Informationen")
    Dim selectedPath As String
    
    If IsEmpty(wsMA.Range("A1").value) Then
        wsMA.Unprotect "{0}"
        Dim response As VbMsgBoxResult
        response = MsgBox("Möchtest du ein Kopie, für die Mitarbeiter mit read_only Modus erstellen? Beim Speichern wird die Kopie Automatisch mit aktualisiert. (Der Pfad wird in der A-Celle der Seite Bereichs Informationen versteckt)", vbYesNo)
        
        If response = vbYes Then
            ' Ask the user to select a directory for saving copies
            Dim folderDialog As FileDialog
            Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
            folderDialog.Title = "Select a directory for saving copies"
            
            If folderDialog.Show = -1 Then ' If the user selects a folder
                selectedPath = folderDialog.SelectedItems(1)
                wsMA.Range("A1").value = selectedPath ' Save the selected path in cell A1
                CopySheetsToDirectory selectedPath  
            End If
        Else
            wsMA.Range("A1").value = "no" ' Save "no" in cell A1 to skip this action in the future
        End If
        wsMA.Protect "{0}"
    Else
        If wsMA.Range("A1").Value <> "no" Then
            CopySheetsToDirectory wsMA.Range("A1").Value
        End If
    End If
    ' Clear the "Bitte warten" message from the status bar
    Application.StatusBar = False
End Sub

Sub CopySheetsToDirectory(destinationPath As String)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False ' Disable alerts for sheet deletion

    Dim newWb As Workbook
    Dim filePath As String
    Dim activeSheetName As String

    ' Get the name of the current workbook without the file extension
    Dim currentWorkbookName As String
    currentWorkbookName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "_MA.xlsx"

    ' Define the file path for the consolidated workbook
    filePath = destinationPath & "\\" & currentWorkbookName

    ' Check if the destination directory exists
    If Len(Dir(destinationPath, vbDirectory)) = 0 Then
        Dim wsMA As Worksheet
        Set wsMA = ThisWorkbook.Sheets("Bereichs Informationen")
        wsMA.Unprotect "{0}"
        wsMA.Range("A1").value = ""
        wsMA.Protect "{0}"
        ' Close the consolidated workbook without saving changes
        Application.DisplayAlerts = True ' Enable alerts
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "Error: Daten konnten nicht in MA Datei gespeichert werden. Bitte nochmal speichern und neuen Pfad bestimmen."
        Exit Sub
    End If
    
    Dim i As Integer
    ' Check if the file already exists
    If Len(Dir(filePath)) > 0 Then
        ' If the file exists, open it and update the sheets
        Set newWb = Workbooks.Open(filePath, WriteResPassword:="DasIstEinGeheim")

        ' Delete sheet in newWb that is within the specified range
        For Each newSheet In newWb.Sheets
            Dim sheetName As String
            sheetName = newSheet.Name
            newSheet.Delete

            ' Check if the sheet with the same name exists in ThisWorkbook
            On Error Resume Next
            ThisWorkbook.Sheets(sheetName).Copy After:=newWb.Sheets(newWb.Sheets.Count)
            On Error GoTo 0
        Next newSheet

        activeSheetName = ThisWorkbook.ActiveSheet.Name

        ' Activate the sheet in newWb based on the active sheet name in ThisWorkbook
        On Error Resume Next ' Handle the case where the sheet may not exist in newWb
        newWb.Sheets(activeSheetName).Activate
        On Error GoTo 0 ' Reset error handling to default

        newWb.ActiveSheet.Range("A1").Select

        newWb.Save
    Else
        ' If the file doesn't exist, create a new workbook
        Set newWb = Workbooks.Add
        ThisWorkbook.Sheets.Copy Before:=newWb.Sheets(1)
        ' Delete the sheets you don't need in the new workbook
        For i = newWb.Sheets.Count To 1 Step -1
            If Not SheetIsInRange(newWb.Sheets(i).Name) Then
                newWb.Sheets(i).Delete
            End If
        Next i
        ' Save the consolidated workbook with a password

        activeSheetName = ThisWorkbook.ActiveSheet.Name

        ' Activate the sheet in newWb based on the active sheet name in ThisWorkbook
        On Error Resume Next ' Handle the case where the sheet may not exist in newWb
        newWb.Sheets(activeSheetName).Activate
        On Error GoTo 0 ' Reset error handling to default

        newWb.ActiveSheet.Range("A1").Select

        newWb.SaveAs filePath, WriteResPassword:="DasIstEinGeheim"
 
    End If

    ' Close the consolidated workbook without saving changes
    newWb.Close SaveChanges:=False


    Application.DisplayAlerts = True ' Enable alerts
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Function SheetIsInRange(sheetName As String) As Boolean
    Dim i As Integer
    For i = 1 To 52
        If sheetName = "KW" & i Then
            SheetIsInRange = True
            Exit Function
        End If
    Next i
    SheetIsInRange = False
End Function


Function KeyExists(col As Collection, key As Variant) As Boolean
    On Error Resume Next
    col.Add "", key
    KeyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    ' Reverse the result
    KeyExists = Not KeyExists
End Function

Function IsNotDouble(value As Variant) As Boolean
    ' Check if the value is not a Double data type
    IsNotDouble = Not IsNumeric(value) Or VarType(value) <> vbDouble
End Function


Sub CheckAndLogChanges()
    Dim ws2 As Worksheet
    Dim key As String
    Dim value As Variant
    Dim CellKey As String
    Dim Sh As Worksheet
    Dim cell As Range
    Dim targetRange As Range
    Dim cellValue As Variant
    Dim cellType As String
    
    Set ws2 = ThisWorkbook.Worksheets("Log")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ShouldRunWorkbookOpen = False
    
    If WorkbookOpenHasRun Then
        For Each Sh In ThisWorkbook.Sheets
            If SheetIsInRange(Sh.Name) Then
                Set targetRange = Sh.UsedRange
                
                Dim data As Variant
                data = targetRange.Value ' Read entire range into array
                
                For i = 1 To targetRange.Rows.Count
                    For j = 1 To targetRange.Columns.Count
                        CellKey = Sh.Name & "_" & targetRange.Cells(i, j).Address
                        cellValue = data(i, j)
                        
                        If IsNotDouble(cellValue) Then
                            If KeyExists(OriginalValues, CellKey) Then
                                value = OriginalValues(CellKey)
                                
                                If value <> cellValue Then
                                    LogData.Add Array(Now(), Environ("UserName"), Sh.Name, "Cell " & targetRange.Cells(i, j).Address, IIf(IsEmpty(cellValue), "", cellValue), value)
                                    ShouldRunWorkbookOpen = True
                                End If
                            ElseIf Not IsEmpty(cellValue) Then
                                LogData.Add Array(Now(), Environ("UserName"), Sh.Name, "Cell " & targetRange.Cells(i, j).Address, IIf(IsEmpty(cellValue), "", cellValue), "")
                                ShouldRunWorkbookOpen = True
                            End If
                        End If
                    Next j
                Next i
            End If
        Next Sh
    End If
    
    If LogData.Count > 0 Then
        UpdateLogTable
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    If ShouldRunWorkbookOpen Then
        ShouldRunWorkbookOpen = False
        Set OriginalValues = InitializeOriginalValues
        WorkbookOpenHasRun = True
    End If
End Sub

Sub UpdateLogTable()
    Dim ws2 As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim rowData As Variant

    Set ws2 = ThisWorkbook.Worksheets("Log")

    ws2.Unprotect "{0}"
    
    ' Check if the table exists, and if not, create it
    On Error Resume Next
    Set tbl = ws2.ListObjects("LogTable")
    On Error GoTo 0

    If tbl Is Nothing Then
        ' Create a new table with the specified header
        Set tbl = ws2.ListObjects.Add(xlSrcRange, ws2.Range("A1:F1"), , xlYes)
        tbl.Name = "LogTable"
        tbl.HeaderRowRange.Cells(1, 1).value = "Datum"
        tbl.HeaderRowRange.Cells(1, 2).value = "Benutzer"
        tbl.HeaderRowRange.Cells(1, 3).value = "Seite"
        tbl.HeaderRowRange.Cells(1, 4).value = "Zelle"
        tbl.HeaderRowRange.Cells(1, 5).value = "Neuer Inhalt"
        tbl.HeaderRowRange.Cells(1, 6).value = "Alter Inhalt"

        ' Set column widths for all columns
        Dim col As ListColumn
        For Each col In tbl.ListColumns
            col.Range.ColumnWidth = 20
        Next col

        ' Set cell alignment to center for the entire table
        tbl.HeaderRowRange.HorizontalAlignment = xlCenter
        tbl.HeaderRowRange.VerticalAlignment = xlCenter
    End If

    ' Show "Bitte warten" message in the status bar
    Application.StatusBar = "Bitte warten...
        
    ' Add multiple rows to the table using a loop
    For Each rowData In LogData
        Set newRow = tbl.ListRows.Add(Position:=1, AlwaysInsert:=True)
        newRow.Range.value = rowData
    Next rowData
    
    ' Clear the "Bitte warten" message from the status bar
    Application.StatusBar = False

    ws2.Protect "{0}"

    ' Clear the LogData collection after updating the table
    Set LogData = New Collection
End Sub
"""
