Attribute VB_Name = "WSWorker"

Public Sub ResetWorksheet(ByRef WS As Worksheet)
    'The script clears everything from the worksheet exept the columnwidth
    'It brings back the normal view of the page, positioning the screen on the first row
    WS.Activate
    ActiveWindow.View = xlNormalView
    WS.PageSetup.PrintArea = ""
    ActiveWindow.ScrollRow = 1 'positioning the screen on the first row
    WS.Cells.Clear
    WS.Rows.EntireRow.RowHeight = WS.StandardHeight
    WS.Cells.PageBreak = xlPageBreakNone
    WS.Range("A1").Select
End Sub

Private Sub CleanPrint(ByVal WS As Worksheet, Source As Variant, StartRow As Long, StartCol As Long, LastRow As Long, LastCol As Long, Optional Transpose As Boolean)
Dim EventsTriger As Boolean
'TODO: Добави лимита на резултатите от търсенето към настройките
EventsTriger = Application.EnableEvents
Application.EnableEvents = False
With WS.Range(WS.Cells(StartRow, StartCol), WS.Cells(LastRow, LastCol))
    .ClearContents
    .ClearContents  ' ако има филтри върху листа, не всичко се изтрива от първи път
    .ClearFormats
End With
    
    If IsArray(Source) Then
        If UBound(Source, 2) > 65000 Then Call EmergencyExit("Прекалено много резултати за показване (" & UBound(Source, 2) & " при лимит 65000). Свийте критериите за търсенето за да съкратите броя на резултатите.")
        Call PrintArray(Source, WS.Cells(StartRow, StartCol), Transpose)
    End If
Application.EnableEvents = EventsTriger
End Sub
Public Sub PrintArray(v As Variant, rng As Range, Optional Transpose As Boolean)
    If Not IsArray(v) Then Exit Sub
        Select Case Transpose
            Case True: rng.Resize(UBound(v, 2) - LBound(v, 2) + 1, UBound(v, 1) - LBound(v, 1) + 1) = TransposeArray(v)
            Case False: rng.Resize(UBound(v, 1) - LBound(v, 1) + 1, UBound(v, 2) - LBound(v, 2) + 1) = v
        End Select
End Sub

Public Sub ClearPages(ByVal WS As Worksheet, Optional StartRow As Long = 1, Optional clearShapes As Boolean = False, Optional ShapeTag As String)
    Dim i As Long, LastRow As Long
    Dim rng As Range
On Error GoTo ErrHandler
If Not IsInitialized Then Call Inicial_Main
Call Optimization_ON
    LastRow = 65536
    WS.Rows(StartRow & ":" & LastRow).ClearContents
    WS.Rows(StartRow & ":" & LastRow).ClearFormats
    If clearShapes Then
        Dim shp As Shape
        For Each shp In PrintSheet.Shapes
         If shp.name Like "*" & ShapeTag & "*" Then shp.Delete
        Next shp
    End If
Call Optimization_OFF

Exit Sub

ErrHandler:
Call EmergencyExit("Програмата не успя да изчисти съдържанието от лист " & WS.name)
End Sub

Public Sub DeleteUserMenu()
    Dim CntxMenu As CommandBar
    Dim UserMenu As CommandBarControl
Set CntxMenu = Application.CommandBars("Cell")
For Each UserMenu In CntxMenu.Controls
    If UserMenu.Tag = "AddedByUser" Then UserMenu.Delete    ' Delete the menu added by user.
Next UserMenu
End Sub

Public Sub SortClients()
Dim TempRange As Range, ClNastr As Collection
Dim LastRow As Long
If Not IsInitialized Then Call Inicial_Main
On Error GoTo ErrHandler
    Set ClNastr = Nastr("RazpredKlientiMark")
    LastRow = RAZPSheet.Cells(Rows.Count, ClNastr("ObektCol").Val).End(xlUp).Row
    If LastRow > ClNastr("StartRow").Val Then
        Set TempRange = RAZPSheet.Range(RAZPSheet.Cells(ClNastr("StartRow").Val, 1), RAZPSheet.Cells(LastRow, ClNastr("LastCol").Val))
        TempRange.Sort Key1:=RAZPSheet.Cells(ClNastr("StartRow").Val, 1), Order1:=xlAscending, _
                                    Key2:=RAZPSheet.Cells(ClNastr("StartRow").Val, 2), Order2:=xlAscending, _
                                    Key3:=RAZPSheet.Cells(ClNastr("StartRow").Val, 4), Order3:=xlAscending, _
                                    Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                                    DataOption1:=xlSortNormal, DataOption2:=xlSortNormal, DataOption3:=xlSortNormal
    End If
Exit Sub

ErrHandler:
Call EmergencyExit("Функция SortClients")
End Sub

Public Sub InsertBlok(ByRef TransRange As Range, ByRef DestinRange As Range, Optional ByRef NewLastRow As Long)
TransRange.Copy Destination:=DestinRange
NewLastRow = DestinRange.Row + TransRange.Rows.Count - 1
End Sub


Public Sub SetColumnWidth()
If Not IsInitialized Then Call Inicial_Main
    Dim v As Variant, SColl As Collection
    Set SColl = Nastr("TovarGACols")
        For Each v In SColl
            TGASheet.Columns(CInt(v.Prop)).ColumnWidth = getColWidth(v.Val, v.Prop, TGASheet)
        Next v
    TGASheet.Columns("A:P").Copy
    TMASheet.Columns("A:P").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub
        
        Public Sub SetRngFormat(ByRef rRng As Range, InString As String)
            Dim vSplit() As String, i As Long
            vSplit = Split(LCase(InString), ";")
            With rRng
                For i = LBound(vSplit) To UBound(vSplit)
                    Select Case vSplit(i)
                        Case "hc": .HorizontalAlignment = xlCenter: GoTo NextFormat
                        Case "hr": .HorizontalAlignment = xlRight: GoTo NextFormat
                        Case "hl": .HorizontalAlignment = xlLeft: GoTo NextFormat
                        Case "vc": .VerticalAlignment = xlCenter: GoTo NextFormat
                        Case "vb": .VerticalAlignment = xlBottom: GoTo NextFormat
                        Case "vt": .VerticalAlignment = xlTop: GoTo NextFormat
                        Case "ei": .Font.Italic = True: GoTo NextFormat
                        Case "eb": .Font.Bold = True: GoTo NextFormat
                        Case "eu": .Font.Underline = True: GoTo NextFormat
                        Case "m": .Merge: GoTo NextFormat
                        Case "ww": .WrapText = True: GoTo NextFormat
                        Case "ft": .NumberFormat = "@": GoTo NextFormat
                        Case "color:grey": .Interior.ColorIndex = 16: GoTo NextFormat
                        Case Else: EmergencyExit ("Неизвестен тип формат '" & vSplit(i) & "', посочен в настройките. Целият формат:" & InString)
NextFormat:
                    End Select
                Next i
            End With
        End Sub
        Public Sub SetBorder(ByRef rRng As Range, InString As String)
            Dim vSplit() As String, i As Long
            vSplit = Split(LCase(InString), ";")
            With rRng
                For i = LBound(vSplit) To UBound(vSplit)
                    Select Case vSplit(i)
                        Case "o"    'outside border only
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                            .Borders(xlEdgeTop).LineStyle = xlContinuous
                        Case "e": .Borders.LineStyle = xlContinuous ' each edge border
                        Case Else: EmergencyExit ("Неизвестна настройка за разчертаване на границите на клектите '" & vSplit(i) & "', посочен в описанието на шаблона за печат. Целият формат:" & InString)
                    End Select
                Next i
            End With
        End Sub

        
        Public Function getColWidth(ByVal Target_Width As Double, ByVal col As Byte, WS As Worksheet) As Double
            Dim TempBool As Boolean
            Dim ratio As Double
            Dim Init_Width As Double, Dot_Step As Double, StepsToTarget As Double
        TempBool = Application.ScreenUpdating
        Application.ScreenUpdating = False
        With WS
            If Target_Width = 0 Then .Columns(col).ColumnWidth = 0: GoTo Ready
            If .Columns(col).Width = Target_Width Then GoTo Ready

            Init_Width = .Columns(col).Width
            .Columns(col).ColumnWidth = .Columns(col).ColumnWidth + 0.1
            Dot_Step = .Columns(col).Width - Init_Width
            StepsToTarget = (Target_Width - .Columns(col).Width) / Dot_Step
            .Columns(col).ColumnWidth = .Columns(col).ColumnWidth + (0.1 * StepsToTarget)
            
            ratio = Target_Width / .Columns(col).Width
            .Columns(col).ColumnWidth = .Columns(col).ColumnWidth * ratio
            
            While .Columns(col).Width > Target_Width
                .Columns(col).ColumnWidth = .Columns(col).ColumnWidth - 0.1
            Wend
            While .Columns(col).Width < Target_Width
                .Columns(col).ColumnWidth = .Columns(col).ColumnWidth + 0.1
            Wend
Ready:
            getColWidth = .Columns(col).ColumnWidth
        End With
        Application.ScreenUpdating = TempBool
        End Function

Public Function RangepoID(IndicateString As String) As Range
If Not IsInitialized Then Call Inicial_Main
Const FRString As String = "#Start", LCString As String = "#Lcol", LRCString As String = "#Lrow"
Dim h As Long, i As Long, j As Long, k As Long
Dim TempStr As String
Dim NastrLastRow As Long
NastrLastRow = NastrSheet.Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To NastrLastRow
    If NastrSheet.Cells(i, 1).Value = IndicateString Then
        For h = i + 1 To NastrLastRow
            If NastrSheet.Cells(h, 1).Value = FRString Then
                For j = 1 To NastrSheet.Cells(h, Columns.Count).End(xlToLeft).Column
                    If NastrSheet.Cells(h, j).Value = LCString Then
                        For k = h + 1 To 65536
                            If NastrSheet.Cells(k, j).Value = LRCString Then
                                TempStr = NastrSheet.Cells(k - 1, j - 1).Address
                                GoTo SuccesExit
                            End If
                        Next k
                        GoTo ErrExit
                    End If
                Next j
                GoTo ErrExit
            End If
        Next h
        GoTo ErrExit
    End If
Next i

ErrExit:
    MsgBox "Не мога да намеря граници на блока с име " + IndicateString + _
            " от работния лист с име " + NastrSheet.name + "."
    Set RangepoID = NastrSheet.Range("$B$1:$B$1")
    Exit Function

SuccesExit:
Set RangepoID = NastrSheet.Range("$B$" & h & ":" & TempStr)

End Function



Public Function IsMarkedParticularRange(ByVal Marked As Range, Optional ByVal CheckCol As Long, _
                                Optional ByVal StartRow As Long, Optional ByVal LastRow As Long) As Boolean
    Dim RngCell As Range
IsMarkedParticularRange = False
If CheckCol > 0 Then
    For Each RngCell In Selection.Rows
        If Not RngCell.Column = CheckCol Then GoTo FalseExit
    Next RngCell
End If
If StartRow > 0 Then
    For Each RngCell In Selection.Rows
        If RngCell.Row < StartRow Then GoTo FalseExit
    Next RngCell
End If
If LastRow > 0 Then
    For Each RngCell In Selection.Rows
        If RngCell.Row > LastRow Then GoTo FalseExit
    Next RngCell
End If
    'Here code will executes if only if it is needed range
        IsMarkedParticularRange = True
        Exit Function
            
FalseExit:
Exit Function

End Function

