Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class Commands
    <CommandMethod("GetStationLabelLocation")>
    Public Sub GetStationLabelLocation()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim editor As Editor = doc.Editor

        Dim labelLocations = AlignmentLabelHelper.GetStationLabelLocation()
        If labelLocations.Count > 0 Then
            For Each location As Point3d In labelLocations
                editor.WriteMessage($"Label Location: X={location.X}, Y={location.Y}, Z={location.Z}" & vbCrLf)
            Next
        Else
            editor.WriteMessage("No station labels found." & vbCrLf)
        End If
    End Sub
End Class
