Sub FolderStructureToColumns_Improved()
    Dim ws As Worksheet
    Dim fldr As FileDialog
    Dim BaseFolder As String
    Dim wsName As String
    Dim OutputRow As Long
    Dim maxDepth As Integer
    Dim fileCount As Long
    wsName = "Folder Structure Columns"

    ' Safely delete old output sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Dim wsDelete As Worksheet
    For Each wsDelete In ThisWorkbook.Worksheets
        If wsDelete.Name = wsName Then
            wsDelete.Delete
            Exit For
        End If
    Next wsDelete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Ask user to select a folder
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select the root folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No folder selected. Exiting macro.", vbExclamation
            Exit Sub
        End If
        BaseFolder = .SelectedItems(1)
    End With

    ' First pass: get max depth and file count
    fileCount = 0
    maxDepth = GetMaxFolderDepthAndCount(BaseFolder, 1, fileCount)
    If fileCount = 0 Then
        MsgBox "No files found in the selected folder.", vbExclamation
        Exit Sub
    End If

    ' Create worksheet and headers
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = wsName

    Dim i As Integer
    For i = 1 To maxDepth
        ws.Cells(1, i).Value = "Level " & i
    Next i
    ws.Cells(1, maxDepth + 1).Value = "File Name"
    ws.Cells(1, maxDepth + 2).Value = "File Path"
    ws.Cells(1, maxDepth + 3).Value = "File Size (KB)"
    ws.Cells(1, maxDepth + 4).Value = "Date Modified"
    ws.Cells(1, maxDepth + 5).Value = "Hyperlink"
    OutputRow = 2

    Application.ScreenUpdating = False
    Call WriteFolderRowsImproved(BaseFolder, ws, OutputRow, Array(), maxDepth)
    Application.ScreenUpdating = True

    ws.Columns.AutoFit

    ' Summary
    ws.Cells(OutputRow + 1, 1).Value = "Total Files:"
    ws.Cells(OutputRow + 1, 2).Value = fileCount
    ws.Cells(OutputRow + 2, 1).Value = "Max Depth:"
    ws.Cells(OutputRow + 2, 2).Value = maxDepth

    MsgBox "Folder structure export complete!" & vbCrLf & vbCrLf & _
           "Total files: " & fileCount & vbCrLf & _
           "Max depth: " & maxDepth, vbInformation
End Sub

Function GetMaxFolderDepthAndCount(FolderPath As String, currentDepth As Integer, ByRef fileCount As Long) As Integer
    Dim FSO As Object, Folder As Object, SubFolder As Object
    Dim maxSubDepth As Integer, subDepth As Integer
    Dim hasFile As Boolean

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(FolderPath)
    If Folder.Files.Count > 0 Then
        hasFile = True
        fileCount = fileCount + Folder.Files.Count
    Else
        hasFile = False
    End If

    maxSubDepth = IIf(hasFile, currentDepth, 0)

    For Each SubFolder In Folder.SubFolders
        subDepth = GetMaxFolderDepthAndCount(SubFolder.Path, currentDepth + 1, fileCount)
        If subDepth > maxSubDepth Then maxSubDepth = subDepth
    Next SubFolder

    GetMaxFolderDepthAndCount = maxSubDepth
End Function

Sub WriteFolderRowsImproved(ByVal FolderPath As String, ws As Worksheet, ByRef OutputRow As Long, ByVal ParentLevels As Variant, ByVal maxDepth As Integer)
    Dim FSO As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim FileItem As Object
    Dim Levels() As String
    Dim i As Integer

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(FolderPath)

    ' Prepare Levels array for this row
    If IsEmpty(ParentLevels) Or (IsArray(ParentLevels) And UBound(ParentLevels) = -1) Then
        ReDim Levels(0)
        Levels(0) = Folder.Name
    ElseIf IsArray(ParentLevels) And ParentLevels(0) = "" Then
        ReDim Levels(0)
        Levels(0) = Folder.Name
    Else
        ReDim Levels(UBound(ParentLevels) + 1)
        For i = 0 To UBound(ParentLevels)
            Levels(i) = ParentLevels(i)
        Next i
        Levels(UBound(Levels)) = Folder.Name
    End If

    ' Only write file rows, never empty folder rows
    For Each FileItem In Folder.Files
        For i = 0 To UBound(Levels)
            ws.Cells(OutputRow, i + 1).Value = Levels(i)
        Next i
        ' Fill unused level columns with blanks
        For i = UBound(Levels) + 1 To maxDepth - 1
            ws.Cells(OutputRow, i + 1).Value = ""
        Next i
        ws.Cells(OutputRow, maxDepth + 1).Value = FileItem.Name
        ws.Cells(OutputRow, maxDepth + 2).Value = FileItem.Path
        ws.Cells(OutputRow, maxDepth + 3).Value = Format(FileItem.Size / 1024, "0.00")
        ws.Cells(OutputRow, maxDepth + 4).Value = FileItem.DateLastModified
        ws.Hyperlinks.Add Anchor:=ws.Cells(OutputRow, maxDepth + 5), Address:=FileItem.Path, TextToDisplay:="Open File"
        OutputRow = OutputRow + 1
    Next FileItem

    ' Recurse into subfolders
    For Each SubFolder In Folder.SubFolders
        Call WriteFolderRowsImproved(SubFolder.Path, ws, OutputRow, Levels, maxDepth)
    Next SubFolder
End Sub