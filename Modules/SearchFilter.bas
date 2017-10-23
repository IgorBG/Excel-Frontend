Attribute VB_Name = "SearchFilter"
Option Explicit

Public Sub DBSearchToWS(ByVal Subject As String, Optional ByVal cFilters As Collection, Optional OrderBy As String)
'sqlCmd = "SELECT RecID, NPnum, Sklad, OrderID, DateEnd, Client, ZaqvkaNum , NomenklNom, NomenklIme, Qnty FROM Orders LEFT JOIN ManufactureOrders ON Orders.NPId = ManufactureOrders.NPID "
'sqlCmd = "SELECT NPID, NPnum, DateNP, Company,  Direction, Note, ClientPurch, MOProgress, OrdCnt, AprWait, CutListWait FROM MOMonitorQuery "
    Dim WS As Worksheet
    Dim sqlCmd As String
    Dim Transpose As Boolean
    Dim FiltrMark As String
    Dim cMark As Collection

    Dim DBRequest As CDBDataRequest
    Dim Tek_Filter As CNastrojka
Select Case Subject
    Case "Order"
        Set cMark = Nastr.Item("OrdMark")
        Set WS = OSheet
        FiltrMark = "OrdFltr"
        Transpose = True
    Case "NP"
        Set cMark = Nastr.Item("NPMark")
        Set WS = NPSheet
        FiltrMark = "NPFltr"
        Transpose = True
    Case Else
        EmergencyExit ("Необозначен режим на работа")
End Select

Set DBRequest = getDBRequest(Subject, DReq)
    DBRequest.FilterReset
    For Each Tek_Filter In cFilters
        With Nastr.Item(FiltrMark).Item(Tek_Filter.Ime)
            Call DBRequest.FilterAddByValues(.Item("AtributeName").Stojnost, .Item("CompareMode").Stojnost, Tek_Filter.Stojnost, .Item("OutputType").Stojnost)
        End With
    Next Tek_Filter
Call AddDatasource(Subject, DS, DBRequest.getDatasource, True)
Call CleanPrint(WS, DS(Subject).getContent, cMark("StartRow").Stojnost, cMark("StartCol").Stojnost, 65536, 255, Transpose)

End Sub

Public Function GetFiltersFromWS(FltrMark As String, WS As Worksheet) As Collection
    Dim TekFilter As Collection
    Dim TekNastr As CNastrojka
    Dim cFilters As New Collection
For Each TekFilter In Nastr.Item(FltrMark)
        If Not Len(WS.Cells(TekFilter("Row").Stojnost, TekFilter("Column").Stojnost).Value) = 0 Then
            Set TekNastr = New CNastrojka
                TekNastr.Ime = TekFilter("Name").Stojnost
                TekNastr.Stojnost = WS.Cells(TekFilter("Row").Stojnost, TekFilter("Column").Stojnost).Value
                cFilters.Add TekNastr
            Set TekNastr = Nothing
        End If
Next TekFilter
Set GetFiltersFromWS = cFilters
End Function


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



