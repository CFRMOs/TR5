'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
'Imports Autodesk.AutoCAD.EditorInput
'Imports Autodesk.AutoCAD.Geometry
'Imports Autodesk.AutoCAD.Runtime
'Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity

'Public Class PolylineFinderBetweenStations
'    <CommandMethod("FindPolylinesBetweenStations")>
'    Public Sub FindPolylinesBetweenStations()
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim ed As Editor = doc.Editor

'        ' Solicitar al usuario la primera estación
'        Dim ppr1 As PromptPointResult = ed.GetPoint("Seleccione la primera estación: ")
'        If ppr1.Status <> PromptStatus.OK Then
'            Return
'        End If
'        Dim stationPoint1 As Point3d = ppr1.Value

'        ' Solicitar al usuario la segunda estación
'        Dim ppr2 As PromptPointResult = ed.GetPoint("Seleccione la segunda estación: ")
'        If ppr2.Status <> PromptStatus.OK Then
'            Return
'        End If
'        Dim stationPoint2 As Point3d = ppr2.Value

'        ' Encontrar polilíneas entre las dos estaciones
'        Dim polylinesBetweenStations As List(Of ObjectId) = GetPolylinesBetweenStations(stationPoint1, stationPoint2)

'        ' Mostrar resultados
'        ed.WriteMessage($"{polylinesBetweenStations.Count} polilíneas encontradas entre las estaciones seleccionadas:")
'        For Each id As ObjectId In polylinesBetweenStations
'            ed.WriteMessage(vbCrLf & $"Polilínea ID: {id}")
'        Next
'    End Sub

'    Public Function GetPolylinesBetweenStations(stationPoint1 As Point3d, stationPoint2 As Point3d) As List(Of ObjectId)
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim db As Database = doc.Database
'        Dim polylines As New List(Of ObjectId)

'        Using tr As Transaction = db.TransactionManager.StartTransaction()
'            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
'            Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

'            For Each objId As ObjectId In btr
'                Dim entity As Entity = tr.GetObject(objId, OpenMode.ForRead)
'                If TypeOf entity Is Polyline Then
'                    Dim polyline As Polyline = CType(entity, Polyline)
'                    If IsPolylineBetweenStations(stationPoint1, stationPoint2, polyline) Then
'                        polylines.Add(objId)
'                    End If
'                End If
'            Next

'            tr.Commit()
'        End Using

'        Return polylines
'    End Function

'    Private Function IsPolylineBetweenStations(stationPoint1 As Point3d, stationPoint2 As Point3d, polyline As Polyline) As Boolean
'        Dim minX As Double = Math.Min(stationPoint1.X, stationPoint2.X)
'        Dim maxX As Double = Math.Max(stationPoint1.X, stationPoint2.X)
'        Dim minY As Double = Math.Min(stationPoint1.Y, stationPoint2.Y)
'        Dim maxY As Double = Math.Max(stationPoint1.Y, stationPoint2.Y)

'        For i As Integer = 0 To polyline.NumberOfVertices - 1
'            Dim vertexPoint As Point3d = polyline.GetPoint3dAt(i)
'            If vertexPoint.X >= minX AndAlso vertexPoint.X <= maxX AndAlso vertexPoint.Y >= minY AndAlso vertexPoint.Y <= maxY Then
'                Return True
'            End If
'        Next

'        Return False
'    End Function
'End Class
