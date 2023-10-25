VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_InterFace 
   Caption         =   "Module Type Creator"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   OleObjectBlob   =   "frm_InterFace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_InterFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ToggleSelection(listBox As Object, btnToggle As Object)
    Static isSelected As Boolean
    Dim i As Integer
    isSelected = Not isSelected
    For i = 0 To listBox.ListCount - 1
        listBox.Selected(i) = isSelected
    Next i
    btnToggle.Caption = IIf(isSelected, "none", "all")
End Sub

Private Sub btnToggleSelectionLevels_Click()
    ToggleSelection lstBox_Levels, btnToggleSelectionLevels
End Sub

Private Sub btnToggleSelectionModules_Click()
    ToggleSelection lstBox_Modules, btnToggleSelectionModules
End Sub

Private Sub btnToggleSelectionInfills_Click()
    ToggleSelection lstBox_Infills, btnToggleSelectionInfills
End Sub

Private Sub CmndBtn_OK_Click()
    
    Dim fso As Object, i As Integer
    Dim savePath As String, modulePath As String, infillPath As String
    Dim ws As Worksheet
    Dim csvFilename As String, txtFilename As String
    Dim desiredName As String
    Dim fileContent As String
    Dim UTFStream As Object
    Dim FileSize As Long  ' Declare FileSize here once for the entire scope
    
    
    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get the selected save path
    savePath = txtPath.Value
    If savePath = "" Then
        MsgBox "Please select a save path."
        Exit Sub
    End If
    
    ' Create root-level subfolders 'Families' and 'Level Loads' if they don't exist
Dim familiesPath As String, levelLoadsPath As String
familiesPath = savePath & "\Families\"
levelLoadsPath = savePath & "\Level Loads\"

If Not fso.FolderExists(familiesPath) Then MkDir familiesPath
If Not fso.FolderExists(levelLoadsPath) Then MkDir levelLoadsPath

' Create subfolders for Modules and Infills under 'Families' if they don't exist
modulePath = familiesPath & "Modules\"
infillPath = familiesPath & "Infills\"
If Not fso.FolderExists(modulePath) Then MkDir modulePath
If Not fso.FolderExists(infillPath) Then MkDir infillPath

' Create 'Excel' subfolders under Modules and Infills for storing .csv files
Dim moduleExcelPath As String, infillExcelPath As String
moduleExcelPath = modulePath & "Excel\"
infillExcelPath = infillPath & "Excel\"
If Not fso.FolderExists(moduleExcelPath) Then MkDir moduleExcelPath
If Not fso.FolderExists(infillExcelPath) Then MkDir infillExcelPath

' Update the paths for csvFilename to point to the new Excel subfolders
csvFilename = moduleExcelPath & desiredName & ".csv"  ' For Modules
csvFilename = infillExcelPath & desiredName & ".csv"  ' For Infills
 

    
Dim tabToFile As Object
Set tabToFile = CreateObject("Scripting.Dictionary")

' Add tab name to filename mappings for Modules and Infills
tabToFile.Add "EXP_1-infill", "Trinam-QT-WW Module-1 Infill"
tabToFile.Add "EXP_2-infill", "Trinam-QT-WW Module-2 Infill"
tabToFile.Add "EXP_3-infill", "Trinam-QT-WW Module-3 Infill"
tabToFile.Add "EXP_4-infill", "Trinam-QT-WW Module-4 Infill"
tabToFile.Add "EXP_5-infill", "Trinam-QT-WW Module-5 Infill"
tabToFile.Add "EXP_6-infill", "Trinam-QT-WW Module-6 Infill"

tabToFile.Add "EX_GL", "Trinam-QT_Vision Glass"
tabToFile.Add "EX_M82", "Trinam-QT_M82 Balcony Door"
tabToFile.Add "EX_SF", "Trinam-QT_Storefront Door"
tabToFile.Add "EX_LV", "Trinam-QT_Grille"
tabToFile.Add "EX_SP", "Trinam-QT_Mono Spandrel"
tabToFile.Add "EX_SV", "Trinam-QT_Mono Spandrel Vent"
tabToFile.Add "EX_SB", "Trinam-QT_Spandrel IGU"
tabToFile.Add "EX_SW", "Trinam-QT_Spandrel IGU Vent"
tabToFile.Add "EX_F", "Trinam-QT_R3 Flush Panel"
tabToFile.Add "EX_FV", "Trinam-QT_R3 Panel Vent"
tabToFile.Add "EX_RP", "Trinam-QT_Recessed Panel"
tabToFile.Add "EX_LS", "Trinam-QT_Grille Sandwich Panel"
tabToFile.Add "EX_SH", "Trinam-QT_Shift Panel 2.0"
tabToFile.Add "EX_P", "Trinam-QT_Projected Panel"
tabToFile.Add "EX_M90", "Trinam-QT_Operable M90"
tabToFile.Add "EX_F92", "Trinam-QT_Operable F92"
tabToFile.Add "EX_CVLA", "Trinam-QT_Clearview Door Left Active"
tabToFile.Add "EX_CVRA", "Trinam-QT_Clearview Door Right Active"
tabToFile.Add "EX_SVLA", "Trinam-QT_Sunview Door Left Active"
tabToFile.Add "EX_SVRA", "Trinam-QT_Sunview Door Right Active"

' Save selected Levels as .XLSX files
Application.DisplayAlerts = False  ' Turn off overwrite warnings
For i = 0 To lstBox_Levels.ListCount - 1
    If lstBox_Levels.Selected(i) Then
        Set ws = ThisWorkbook.Sheets(Mid(lstBox_Levels.List(i), 2))
        ws.Copy
        ActiveWorkbook.SaveAs Filename:=levelLoadsPath & lstBox_Levels.List(i) & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=False
    End If
Next i
Application.DisplayAlerts = True  ' Turn the warnings back on


' Save selected Modules as .CSV files
For i = 0 To lstBox_Modules.ListCount - 1
    If lstBox_Modules.Selected(i) Then
        Set ws = ThisWorkbook.Sheets(lstBox_Modules.List(i))
        If tabToFile.Exists(lstBox_Modules.List(i)) Then
            desiredName = tabToFile(lstBox_Modules.List(i))
        Else
            MsgBox "Key not found: " & lstBox_Modules.List(i)
            Exit For
        End If
        csvFilename = moduleExcelPath & desiredName & ".csv"  ' Updated to use the Excel subfolder
        SaveWorksheetAsUTF8CSV ws, csvFilename
        txtFilename = modulePath & desiredName & ".txt"
        If fso.FileExists(csvFilename) Then
            fso.CopyFile csvFilename, txtFilename
            ModifyTxtFile txtFilename  ' <-- Add this line here
        End If
    End If
Next i

' Save selected Infills as .CSV files
For i = 0 To lstBox_Infills.ListCount - 1
    If lstBox_Infills.Selected(i) Then
        Set ws = ThisWorkbook.Sheets(lstBox_Infills.List(i))
        If tabToFile.Exists(lstBox_Infills.List(i)) Then
            desiredName = tabToFile(lstBox_Infills.List(i))
        Else
            MsgBox "Key not found: " & lstBox_Infills.List(i)
            Exit For
        End If
        csvFilename = infillExcelPath & desiredName & ".csv"  ' Updated to use the Excel subfolder
        SaveWorksheetAsUTF8CSV ws, csvFilename
        txtFilename = infillPath & desiredName & ".txt"
        If fso.FileExists(csvFilename) Then
            fso.CopyFile csvFilename, txtFilename
            ModifyTxtFile txtFilename  ' <-- Add this line here
        End If
    End If
Next i



    
    MsgBox "Files have been saved successfully."
End Sub
Private Sub CmndBtn_Path_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        ' Set the initial folder to the workbook's directory
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show = -1 Then
            txtPath.Value = .SelectedItems(1)
        End If
    End With
End Sub


Private Sub lstBox_Levels_Click()
    ' Placeholder for any additional logic
End Sub

Private Sub lstBox_Modules_Click()
    ' Placeholder for any additional logic
End Sub

Private Sub lstBox_Infills_Click()
    ' Placeholder for any additional logic
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim formattedName As String
    
    ' Clear the list boxes before populating
    lstBox_Levels.Clear
    lstBox_Modules.Clear
    lstBox_Infills.Clear
    
    ' Loop through each worksheet and add its name to the appropriate list box
    For Each ws In ThisWorkbook.Sheets
        ' Check for numeric names for Levels
        If IsNumeric(ws.Name) Then
            formattedName = "L" & ws.Name
            lstBox_Levels.AddItem formattedName
        End If
        
        ' Check for infill pattern names
        If ws.Name Like "EXP_*infill" Then
            lstBox_Modules.AddItem ws.Name
        End If
        
        ' Check for names starting with EXP_
        If ws.Name Like "EX_*" Then
            lstBox_Infills.AddItem ws.Name
        End If
    Next ws
    ' Set the default save path to the workbook's directory
    txtPath.Value = ThisWorkbook.Path
End Sub

Sub SaveWorksheetAsUTF8CSV(ws As Worksheet, csvFilename As String)
    Dim i As Long, j As Long
    Dim cellValue As String
    Dim rowValues As String
    Dim allRows As String
    Dim UTFStream As Object
    
    ' Loop through each row and column to build the CSV content
    For i = 1 To ws.UsedRange.Rows.Count
        rowValues = ""
        For j = 1 To ws.UsedRange.Columns.Count
            cellValue = CStr(ws.Cells(i, j).Value)  ' Explicitly convert to string
            rowValues = rowValues & cellValue & ","
        Next j
        rowValues = Left(rowValues, Len(rowValues) - 1)  ' Remove trailing comma
        allRows = allRows & rowValues & vbCrLf
    Next i
    
    ' Use ADODB.Stream to save as UTF-8 encoded CSV
    Set UTFStream = CreateObject("ADODB.Stream")
    UTFStream.Type = 2  ' adTypeText
    UTFStream.Charset = "utf-8"
    UTFStream.Open
    UTFStream.WriteText allRows
    UTFStream.SaveToFile csvFilename, 2  ' adSaveCreateOverWrite
    UTFStream.Close
    Set UTFStream = Nothing
    
End Sub
Sub ModifyTxtFile(txtFilename As String)
    Dim fileContent As String
    Dim lines As Variant
    Dim i As Integer
    
    ' Read file content
    With CreateObject("Scripting.FileSystemObject").OpenTextFile(txtFilename, 1)
        fileContent = .ReadAll
        .Close
    End With
    
    ' Split content by lines
    lines = Split(fileContent, vbCrLf)
    
    ' Remove empty or comma-only lines
    Dim newContent As String
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" And Replace(Trim(lines(i)), ",", "") <> "" Then
            newContent = newContent & lines(i) & vbCrLf
        End If
    Next i
    
    ' Write the modified content back to the file
    With CreateObject("Scripting.FileSystemObject").OpenTextFile(txtFilename, 2, True)
        .Write newContent
        .Close
    End With
End Sub
