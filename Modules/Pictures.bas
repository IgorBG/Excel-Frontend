Attribute VB_Name = "Pictures"

Sub InsertImage(ByRef PicRange As Range, ByVal PicPath As String, Optional ByVal AdjustWidth As Boolean, _
                        Optional ByVal AdjustHeight As Boolean, Optional ByVal WS As Worksheet, Optional GivenName As String)
    Dim ph As Picture
    Dim TempBool As Boolean
    Dim K_picture As Double, K_PicRange As Double
On Error Resume Next
TempBool = Application.ScreenUpdating
Application.ScreenUpdating = False
If WS Is Nothing Then Set WS = ActiveSheet

Call DeleteShapes("Picture", WS) ' delete old pictures
Set ph = PicRange.Parent.Pictures.Insert(PicPath) ' Insert new picture
If Not Len(GivenName) = 0 Then ph.Name = GivenName ' Rename the picture

K_picture = ph.Width / ph.Height    ' sides ratio of the picture
K_PicRange = PicRange.Width / PicRange.Height    ' sides ratio of the range

If AdjustWidth Then ph.Width = PicRange.Width: ph.Height = ph.Width / K_picture
If AdjustHeight Then ph.Height = PicRange.Height: ph.Width = ph.Height * K_picture

If K_picture > K_PicRange Then 'If the picture is too wide (wider then the range) just limit the picture width to the range width
    ph.Width = PicRange.Width: ph.Height = PicRange.Height
End If

    ' Put the picture to the center of the range
   ph.Top = PicRange.Top: ph.Left = PicRange.Left + (PicRange.Width / 2) - (ph.Width / 2)
Application.ScreenUpdating = TempBool
End Sub


Public Sub DeleteShapes(ByVal Tag As String, WS As Worksheet)
        Dim shp As Shape
    For Each shp In WS.Shapes
         If shp.Name Like "*" & Tag & "*" Then shp.Delete
    Next shp
End Sub
