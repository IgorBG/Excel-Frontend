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




Public Function TransposeArray(inArray As Variant, Optional ShiftX As Long = 0, Optional ShiftY As Long = 0) As Variant
'The funtion transposes a 2D array and can shift the bases of X and Y in the result Array
    Dim XLower As Long, XUpper As Long, YLower  As Long, YUpper As Long
    Dim x As Long, y As Long
    Dim outArray As Variant
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
