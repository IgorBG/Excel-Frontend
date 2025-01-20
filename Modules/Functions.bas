Attribute VB_Name = "Functions"
Public Function ValueIsInCollection(col As Collection, Key As String) As Boolean
Dim obj As Variant
On Error GoTo err
    ValueIsInCollection = True
    obj = col.Item(Key)
    Exit Function
err:
    ValueIsInCollection = False
End Function

Public Function ObjectIsInCollection(col As Collection, Key As String) As Boolean
Dim obj As Object
On Error GoTo err
    ObjectIsInCollection = True
    Set obj = col.Item(Key)
    Exit Function
err:
    ObjectIsInCollection = False
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


Public Function TransposeArray(inArray As Variant, Optional ShiftX As Long = 0, Optional ShiftY As Long = 0) As Variant
'The funtion transposes a 2D array and can shift the bases of X and Y in the result Array
    Dim XLower As Long, XUpper As Long, YLower  As Long, YUpper As Long
    Dim x As Long, y As Long
    Dim outArray As Variant
On Error GoTo ErrHandler
If Not IsArray(inArray) Then GoTo ErrHandler
    XLower = LBound(inArray, 2)
    YLower = LBound(inArray, 1)
    XUpper = UBound(inArray, 2)
    YUpper = UBound(inArray, 1)
    ReDim outArray(XLower + ShiftX To XUpper + ShiftX, YLower + ShiftY To YUpper + ShiftY)
    For x = XLower To XUpper
        For y = YLower To YUpper
            outArray(x + ShiftX, y + ShiftY) = inArray(y, x)
        Next y
    Next x
    TransposeArray = outArray
Exit Function
ErrHandler:
Set inArray = Nothing
Debug.Print x, y
Debug.Print "Empty Array in TransposeArray Function"
End Function

Public Function GetArray(ByRef WS As Worksheet, ByVal StartRow As Long, ByVal StartCol As Long, ByVal LastRow As Long, ByVal LastCol As Long) As Variant
GetArray = WS.Range(WS.Cells(StartRow, StartCol), WS.Cells(LastRow, LastCol))
End Function

Public Function Sum1DArray(ByVal inArray As Variant) As Double
Dim i As Long
    For i = LBound(inArray) To UBound(inArray)
        Sum1DArray = Sum1DArray + inArray(i)
    Next i
End Function
Public Sub Reset1DArray(ByRef inArray As Variant)
Dim i As Long
    For i = LBound(inArray) To UBound(inArray)
        inArray(i) = 0
    Next i
End Sub

Public Function RangepoID(IndicateString As String) As Range
If Not IsInitialized Then Call Inicial_Main
Const FRString As String = "#Start", LCString As String = "#Lcol", LRCString As String = "#Lrow"
Dim h As Long, i As Long, j As Long, k As Long
Dim TempStr As String
Dim NastrLastRow As Long
NastrLastRow = NastrSheet.Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To NastrLastRow
    If NastrSheet.Cells(i, 1).value = IndicateString Then
        For h = i + 1 To NastrLastRow
            If NastrSheet.Cells(h, 1).value = FRString Then
                For j = 1 To NastrSheet.Cells(h, Columns.Count).End(xlToLeft).Column
                    If NastrSheet.Cells(h, j).value = LCString Then
                        For k = h + 1 To 65536
                            If NastrSheet.Cells(k, j).value = LRCString Then
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

Public Function getNewBean(ByVal Property As String, value As Variant) As CBean
On Error GoTo ErrHandler
Set getNewBean = New CBean
    getNewBean.Prop = Property
    getNewBean.Val = value
Exit Function
ErrHandler:
Call EmergencyExit("Функция getNewBean")
End Function


Public Function CheckedValue(ExpVarType As String, CanNull As Boolean, ByVal Source As Variant, Optional Limit As Long = 0, Optional ReturnNullAs As Variant) As Variant
'Processes simple validation for different ExpVarTypes
    Const ERR_MSG_BASE As String = "Възникна грешка при добавяне на '"
    Dim TempMsg As String
    Dim SourceIsNull As Boolean
On Error GoTo ErrHandler
'First, check for empty string input and checks wheather Null value is acceptable
SourceIsNull = False
    Select Case VarType(Source)
        Case VbVarType.vbNull: SourceIsNull = True
        Case VbVarType.vbString: If Source = vbNullString Then SourceIsNull = True
        Case Else: If Source = 0 Or Source = Empty Then SourceIsNull = True
    End Select
        
    If SourceIsNull Then
        If CanNull = True Then
            If IsMissing(ReturnNullAs) Then
                Select Case LCase(ExpVarType)
                    Case "string": ReturnNullAs = vbNullString
                    Case Else: ReturnNullAs = 0
                End Select
            End If
        CheckedValue = ReturnNullAs:  GoTo CorrectedValue
        Else: TempMsg = Source & "' празна стойност в задължителното поле.": GoTo ErrHandler
        End If
    End If

Select Case LCase(ExpVarType)
Case "long"
    If IsNumeric(Source) Then
        If CLng(Source) = 0 And CanNull = False Then TempMsg = Source & "' празна стойност в задължителното поле.": GoTo ErrHandler
    End If
    If IsNumeric(Source) = False And Len(Source) > 0 Then TempMsg = Source & "' като число.": GoTo ErrHandler
    GoTo ValueOK
Case "string"
    If Not Application.IsText(CStr(Source)) Then TempMsg = "' данните като текст.": GoTo ErrHandler
    If Limit > 0 Then If Len(CStr(Source)) > Limit Then Source = Left(CStr(Source), Limit)
    GoTo ValueOK
Case "date"
    If Not IsDate(Source) Then TempMsg = Source & "' като дата.": GoTo ErrHandler
    If CDbl(CDate(Source)) = 0 And CanNull = False Then TempMsg = Source & "' празна стойност в задължителното поле.": GoTo ErrHandler
    GoTo ValueOK
Case "double"
    If IsError(CDbl(Source)) Then TempMsg = Source & "' като число.": GoTo ErrHandler
    If CDbl(Source) = 0 And CanNull = False Then TempMsg = Source & "' празна стойност в задължителното поле.": GoTo ErrHandler
    GoTo ValueOK
End Select
ValueOK:
CheckedValue = Source
CorrectedValue:
Exit Function
   
ErrHandler:
Call EmergencyExit(ERR_MSG_BASE & TempMsg)
End Function
Public Function ReplaceSymbolsInText(ByVal txt As String, BadSymbols As String, GoodSymbol As String) As String
Dim iCount As Integer
    For iCount = 1 To Len(BadSymbols)
        txt = Replace(txt, Mid(BadSymbols, iCount, 1), GoodSymbol, , , vbTextCompare)
    Next
    ReplaceSymbolsInText = txt
End Function
