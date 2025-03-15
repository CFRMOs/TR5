'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
'Imports Autodesk.AutoCAD.EditorInput
'Imports Autodesk.AutoCAD.Geometry
'Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.Civil

'Public Class AutoCADEntitiesFinder

'    <CommandMethod("FindEntitiesNearPoint")>
'    Public Sub FindEntitiesNearPoint()
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim ed As Editor = doc.Editor

'        Try
'            ' Solicitar al usuario el punto de referencia
'            Dim ppr As PromptPointResult = ed.GetPoint("Seleccione el punto de referencia: ")
'            If ppr.Status <> PromptStatus.OK Then
'                Return
'            End If
'            Dim referencePoint As Point3d = ppr.Value

'            ' Solicitar al usuario el radio de búsqueda
'            Dim pdo As New PromptDistanceOptions("Ingrese el radio de búsqueda: ") With {
'                .BasePoint = referencePoint,
'                .UseBasePoint = True
'            }
'            Dim pdr As PromptDoubleResult = ed.GetDistance(pdo)
'            If pdr.Status <> PromptStatus.OK Then
'                Return
'            End If
'            Dim searchRadius As Double = pdr.Value

'            ' Listar entidades encontradas
'            Dim entitiesFound As List(Of ObjectId) = GetEntitiesNearPoint(referencePoint, searchRadius)

'            ' Mostrar resultados
'            ed.WriteMessage($"{entitiesFound.Count} entidades encontradas cerca del punto {referencePoint}:")
'            For Each id As ObjectId In entitiesFound
'                ed.WriteMessage(vbCrLf & $"Entidad ID: {id}")
'            Next
'        Catch ex As Exception
'            ed.WriteMessage(vbCrLf & $"Error: {ex.Message}")
'        End Try
'    End Sub

'    Private Function GetEntitiesNearPoint(referencePoint As Point3d, searchRadius As Double) As List(Of ObjectId)
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim db As Database = doc.Database
'        Dim entities As New List(Of ObjectId)

'        Using tr As Transaction = db.TransactionManager.StartTransaction()
'            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
'            Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

'            For Each objId As ObjectId In btr
'                Dim entity As Entity = TryCast(tr.GetObject(objId, OpenMode.ForRead), Entity)
'                If entity IsNot Nothing AndAlso IsEntityNearPoint(referencePoint, searchRadius, entity) Then
'                    entities.Add(objId)
'                End If
'            Next

'            tr.Commit()
'        End Using

'        Return entities
'    End Function

'    Private Function IsEntityNearPoint(referencePoint As Point3d, searchRadius As Double, entity As Entity) As Boolean
'        Select Case entity.GetType()
'            Case GetType(Polyline)
'                Return IsPointNearPolyline(referencePoint, searchRadius, DirectCast(entity, Polyline))
'            Case GetType(FeatureLine)
'                Return IsPointNearFeatureLine(referencePoint, searchRadius, DirectCast(entity, FeatureLine))
'            Case Else
'                Return False
'        End Select
'    End Function

'    Private Function IsPointNearPolyline(referencePoint As Point3d, searchRadius As Double, polyline As Polyline) As Boolean
'        ' Verificar si el punto está dentro del radio de búsqueda de la polilínea
'        For i As Integer = 0 To polyline.NumberOfVertices - 1
'            If referencePoint.DistanceTo(polyline.GetPoint3dAt(i)) <= searchRadius Then
'                Return True
'            End If
'        Next

'        ' Verificar si la polilínea intercepta el punto
'        Return polyline.GetClosestPointTo(referencePoint, Vector3d.ZAxis, False).DistanceTo(referencePoint) <= searchRadius
'    End Function

'    Private Function IsPointNearFeatureLine(referencePoint As Point3d, searchRadius As Double, featureLine As FeatureLine) As Boolean
'        ' Verificar si el punto está dentro del radio de búsqueda de la línea de características
'        For Each point As Point3d In featureLine.GetPoints(FeatureLinePointType.AllPoints)
'            If referencePoint.DistanceTo(point) <= searchRadius Then
'                Return True
'            End If
'        Next

'        ' Verificar si la línea de características intercepta el punto
'        Return featureLine.GetClosestPointTo(referencePoint, Vector3d.ZAxis, False).DistanceTo(referencePoint) <= searchRadius
'    End Function

'End Class


