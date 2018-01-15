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




Public Function TransposeArray(myArray As Variant, Optional ShiftX As Long = 0, Optional ShiftY As Long = 0) As Variant
Dim x As Long
Dim Y As Long
Dim XUpper As Long
Dim YUpper As Long
Dim tempArray As Variant
If Not IsArray(myArray) Then GoTo ErrHandler
    XLower = LBound(myArray, 2) + ShiftX
    YLower = LBound(myArray, 1) + ShiftY
    XUpper = UBound(myArray, 2) + ShiftX
    YUpper = UBound(myArray, 1) + ShiftY
    ReDim tempArray(XUpper, YUpper)
    For x = XLower To XUpper
        For Y = YLower To YUpper
            tempArray(x, Y) = myArray(Y, x)
        Next Y
    Next x
    TransposeArray = tempArray
Exit Function
ErrHandler:
Set myArray = Nothing
Debug.Print "Empty Array in TransposeArray Function"
End Function

Public Function GetArray(ByRef WS As Worksheet, ByVal StartRow As Long, ByVal StartCol As Long, ByVal LastRow As Long, ByVal LastCol As Long) As Variant
GetArray = WS.Range(WS.Cells(StartRow, StartCol), WS.Cells(LastRow, LastCol))
End Function

Public Function Sum1DArray(ByVal myArray As Variant) As Double
Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        Sum1DArray = Sum1DArray + myArray(i)
    Next i
End Function
Public Sub Reset1DArray(ByRef myArray As Variant)
Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        myArray(i) = 0
    Next i
End Sub
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
    If CLng(CDate(Source)) = 0 And CanNull = False Then TempMsg = Source & "' празна стойност в задължителното поле.": GoTo ErrHandler
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


Public Function NullHandledString(ByRef InData As Variant) As String
    'Adopted. The function avoids Null values in strings. Converting Null as "", and keepeing the same any full string
    NullHandledString = InData & vbNullString
End Function
