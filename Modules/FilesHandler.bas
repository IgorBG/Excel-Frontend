Attribute VB_Name = "FilesHandler"
Option Explicit

Public Function getOpenedFilesList(Optional WindowTitle As String, Optional FileExtFiltr As String, Optional MultiChoice As Boolean, Optional InitialFldr As String) As Variant
    Dim OpenFileDialog As FileDialog
    Dim TempStr As String
    Dim v As Variant
    Const DELIM As String = ";"
TempStr = vbNullString
Set OpenFileDialog = Application.FileDialog(msoFileDialogOpen)
With OpenFileDialog
    .AllowMultiSelect = MultiChoice
    If Not WindowTitle = vbNullString Then .Title = WindowTitle
    If Not FileExtFiltr = vbNullString Then
        .Filters.Clear
        .Filters.Add UCase(FileExtFiltr) & " File (*." & FileExtFiltr & ")", "*." & FileExtFiltr, 1
    End If
    If Not InitialFldr = vbNullString Then .InitialFileName = InitialFldr & "\*"
    If .Show = -1 Then
        For Each v In .SelectedItems
            TempStr = TempStr & v & DELIM
        Next v
        TempStr = Left(TempStr, Len(TempStr) - Len(DELIM)) 'èçòðèâàìå ïîñëåäíàòà çàïåòàÿ
    Else
        Exit Function
    End If
End With
getOpenedFilesList = Split(TempStr, DELIM)
TempStr = vbNullString
End Function

Public Sub CopyFile(ByVal SourceFullPath As String, ByVal TargetFullPath As String)
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(SourceFullPath, TargetFullPath)
End Sub
       
        
        
        'Not my code! TODO: license
Public Sub List_all_files()            'This is an event handler. It exectues when the user presses the run button determines if the user selects a directory from the folder dialog
Dim intResult As Integer
Dim strPath As String           'the path selected by the user from the folder dialog
Dim objFSO As Object            'Filesystem object
Dim intCountRows As Integer     'the current number of rows

Application.FileDialog(msoFileDialogFolderPicker).Title = "Èçáåðåòå ïàïêà"
intResult = Application.FileDialog(msoFileDialogFolderPicker).Show  'the dialog is displayed to the user

If intResult <> 0 Then      'checks if user has cancled the dialog
Call Optimization_ON
    strPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Set objFSO = CreateObject("Scripting.FileSystemObject")         'Create an instance of the FileSystemObject
    intCountRows = GetAllFiles(strPath, ROW_FIRST, objFSO)          'loops through each file in the directory and prints their names and path
    Call GetAllFolders(strPath, objFSO, intCountRows)               'loops through all the files and folder in the input path
Call Optimization_OFF
End If

End Sub
        'Not my code! TODO: license
    Private Function GetAllFiles(ByVal strPath As String, ByVal intRow As Integer, ByRef objFSO As Object) As Integer
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Long
    i = intRow - ROW_FIRST + 1
    Set objFolder = objFSO.GetFolder(strPath)
    For Each objFile In objFolder.Files
            Cells(i + ROW_FIRST - 1, 1).Value = objFile.Name 'print file name
            Cells(i + ROW_FIRST - 1, 2).Value = objFile.Path 'print file path
            i = i + 1
    Next objFile
    GetAllFiles = i + ROW_FIRST - 1
    End Function
          'Not my code! TODO: license  
    Private Sub GetAllFolders(ByVal strFolder As String, ByRef objFSO As Object, ByRef intRow As Integer)
    Dim objFolder As Object
    Dim objSubFolder As Object
    
    Set objFolder = objFSO.GetFolder(strFolder)         'Get the folder object
    For Each objSubFolder In objFolder.subfolders       'loops through each file in the directory and prints their names and path
        intRow = GetAllFiles(objSubFolder.Path, intRow, objFSO)
        Call GetAllFolders(objSubFolder.Path, objFSO, intRow)   'recursive call to to itsself
    Next objSubFolder
    End Sub
