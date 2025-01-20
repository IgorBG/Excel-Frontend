Attribute VB_Name = "WSWorker"

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
Public Sub InsertBlok(ByRef TransRange As Range, ByRef DestinRange As Range, Optional ByRef NewLastRow As Long)
TransRange.Copy Destination:=DestinRange
NewLastRow = DestinRange.Row + TransRange.Rows.Count - 1
End Sub

Public Function getWSbyCode(ByVal lngType As Long) As Worksheet
On Error GoTo ErrHandler
    Select Case lngType
        Case 0
            Set getWSbyCode = KOMSkladSheet
        Case 1
            Set getWSbyCode = KOMPorchSheet
        Case 2
            Set getWSbyCode = ArtLoadSheet
        Case 3
            Set getWSbyCode = MatrLoadSheet
    End Select
Exit Function
ErrHandler:
Call EmergencyExit("Функция getWSbyCode")
End Function
Public Function getColWithCollectionByCode(ByVal lngType As Long) As Collection
    Dim HeadName As String
If Not IsInitialized Then Call Inicial_Main
On Error GoTo ErrHandler
    Select Case lngType
        Case 0
            HeadName = "KomisSkladColWidth"
        Case 1
            HeadName = "KomisPorychColWidth"
    End Select
If Not ObjectIsInCollection(Nastr, HeadName) Then Call Inicial_Add(Nastr, HeadName)
Set getColWithCollectionByCode = Nastr(HeadName)
Exit Function
ErrHandler:
Call EmergencyExit("Функция getColWithCollectionByCode")
End Function
Public Sub SetColumnWidth(ByRef WS As Worksheet, valColl As Collection)
If Not IsInitialized Then Call Inicial_Main
Call Optimization_ON
    With valColl
        For i = 1 To 21
            WS.Columns(i).ColumnWidth = getColWidth(.Item(i).Val, .Item(i).Prop, WS)
        Next i
    End With
Call Optimization_OFF
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
