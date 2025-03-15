'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
'Imports Autodesk.AutoCAD.EditorInput
'Imports Autodesk.AutoCAD.Geometry
'Imports Autodesk.AutoCAD.Runtime
'Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
'Imports AlignmentLabel = Autodesk.Civil.DatabaseServices.Label
'Imports System.Linq

'Public Class PolylineFinderBetweenStations
'    <CommandMethod("FindPolylinesBetweenStations")>
'    Public Sub FindPolylinesBetweenStations()
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim ed As Editor = doc.Editor

'        ' Solicitar la primera estación al usuario
'        Dim pdr1 As PromptDoubleResult = ed.GetDouble("Ingrese la primera estación: ")
'        If pdr1.Status <> PromptStatus.OK Then
'            ed.WriteMessage("Operación cancelada.")
'            Return
'        End If
'        Dim station1 As Double = pdr1.Value

'        ' Solicitar la segunda estación al usuario
'        Dim pdr2 As PromptDoubleResult = ed.GetDouble("Ingrese la segunda estación: ")
'        If pdr2.Status <> PromptStatus.OK Then
'            ed.WriteMessage("Operación cancelada.")
'            Return
'        End If
'        Dim station2 As Double = pdr2.Value

'        ' Encontrar los puntos correspondientes a las estaciones
'        Dim point1 As Point3d = GetPointFromStation(station1)
'        Dim point2 As Point3d = GetPointFromStation(station2)

'        If point1.Equals(Point3d.Origin) OrElse point2.Equals(Point3d.Origin) Then
'            ed.WriteMessage($"No se encontraron puntos para las estaciones {station1} y {station2}.")
'            Return
'        End If

'        ' Encontrar polilíneas cercanas a cada punto
'        Dim polylinesNearPoint1 As List(Of ObjectId) = GetEntitiesNearPoint(point1, 5.0) ' Radio de ejemplo
'        Dim polylinesNearPoint2 As List(Of ObjectId) = GetEntitiesNearPoint(point2, 5.0) ' Radio de ejemplo

'        ' Identificar polilíneas comunes en ambos conjuntos
'        Dim commonPolylines As List(Of ObjectId) = polylinesNearPoint1.Intersect(polylinesNearPoint2).ToList()

'        ' Mostrar resultados
'        ed.WriteMessage($"{commonPolylines.Count} polilíneas encontradas entre las estaciones {station1} y {station2}:")
'        For Each id As ObjectId In commonPolylines
'            ed.WriteMessage(vbCrLf & $"Polilínea ID: {id}")
'        Next
'    End Sub

'    Private Function GetPointFromStation(station As Double) As Point3d
'        Dim labelLocations As List(Of Point3d) = AlignmentLabelHelper.GetStationLabelLocation()

'        For Each point As Point3d In labelLocations
'            If Math.Abs(point.X - station) < 0.01 Then ' Ajustar precisión según sea necesario
'                Return point
'            End If
'        Next

'        Return Point3d.Origin ' Retorna origen si no se encuentra el punto
'    End Function

'    Private Function GetEntitiesNearPoint(referencePoint As Point3d, searchRadius As Double) As List(Of ObjectId)
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim db As Database = doc.Database
'        Dim entities As New List(Of ObjectId)

'        Using tr As Transaction = db.TransactionManager.StartTransaction()
'            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
'            Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

'            For Each objId As ObjectId In btr
'                Dim entity As Entity = tr.GetObject(objId, OpenMode.ForRead)
'                If TypeOf entity Is Polyline Then
'                    Dim polyline As Polyline = CType(entity, Polyline)
'                    If IsPointNearPolyline(referencePoint, searchRadius, polyline) Then
'                        entities.Add(objId)
'                    End If
'                End If
'            Next

'            tr.Commit()
'        End Using

'        Return entities
'    End Function

'    Private Function IsPointNearPolyline(referencePoint As Point3d, searchRadius As Double, polyline As Polyline) As Boolean
'        ' Verificar si el punto está dentro del radio de búsqueda de la polilínea
'        For i As Integer = 0 To polyline.NumberOfVertices - 1
'            Dim vertexPoint As Point3d = polyline.GetPoint3dAt(i)
'            If referencePoint.DistanceTo(vertexPoint) <= searchRadius Then
'                Return True
'            End If
'        Next

'        ' Verificar si la polilínea intercepta el punto
'        If polyline.GetClosestPointTo(referencePoint, Vector3d.ZAxis, False).DistanceTo(referencePoint) <= searchRadius Then
'            Return True
'        End If

'        Return False
'    End Function
'End Class

