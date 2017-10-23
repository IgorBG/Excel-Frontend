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
        TempStr = Left(TempStr, Len(TempStr) - Len(DELIM)) 'изтриваме последната запетая
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