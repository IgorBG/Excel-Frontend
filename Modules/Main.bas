Attribute VB_Name = "Main"
Option Explicit
Public IsInitialized As Boolean
Public NastrSheet As Worksheet
Public OSheet As Worksheet      'Orders Sheet
Public PakSheet As Worksheet    'Generated pakcs Sheet
Public PrintSheet As Worksheet  'Worksheet with label for print
Public ClntSheet As Worksheet     'Clients sheet
Public ListSheet As Worksheet   'Sheet for different sheets
Public Nastr As Collection      'All the setting in one collection
Public SysMsg As String
Public PrintData As Variant

Public RowClicked As Long
Public ColClicked As Long

Public PrintTemplateString As String
Public PrintPanelReady As Boolean
Public PrintPanelRow As Long
Public PrintPanelCol As Long




Public Sub Optimization_ON()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
End Sub

Public Sub Optimization_OFF()
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.MaxChange = 0.001

Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True
End Sub

Public Sub EmergencyExit(ErrMsg As String)
Call Optimization_OFF
IsInitialized = False
PrintPanelReady = False
MsgBox "Програмата е спряна. Причина: " & ErrMsg: End
End Sub
Public Sub UserExit()
Call Optimization_OFF
IsInitialized = False
PrintPanelReady = False
End
End Sub



Public Sub ResetInicialization()

Set OSheet = Nothing
Set PakSheet = Nothing
Set PrintSheet = Nothing
Set ClntSheet = Nothing
Set NastrSheet = Nothing

Set Nastr = Nothing
Set PrintData = Nothing

SysMsg = vbNullString
IsInitialized = False
PrintPanelReady = False
PrintTemplateString = vbNullString
Call Optimization_OFF
End Sub
