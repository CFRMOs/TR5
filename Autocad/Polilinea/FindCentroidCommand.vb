Imports System.Linq
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Module FindCentroidCommand

    <CommandMethod("FindCentroidAndPerpendicular")>
    Public Sub FindCentroidAndPerpendicular()
        ' Obtener el documento actual y el editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor

        ' Solicitar al usuario que seleccione una polilínea cerrada
        Dim opt As New PromptEntityOptions(vbLf & "Seleccione una polilínea cerrada:")
        opt.SetRejectMessage(vbLf & "Debe seleccionar una polilínea cerrada.")
        opt.AddAllowedClass(GetType(Polyline), False)

        Dim res As PromptEntityResult = acEd.GetEntity(opt)

        If res.Status <> PromptStatus.OK Then
            acEd.WriteMessage(vbLf & "Operación cancelada.")
            Return
        End If

        Dim polylineId As ObjectId = res.ObjectId

        ' Calcular el centroide y crear la línea perpendicular
        Using trans As Transaction = acDoc.TransactionManager.StartTransaction()
            Try
                Dim polyline As Polyline = TryCast(trans.GetObject(polylineId, OpenMode.ForRead), Polyline)

                If polyline Is Nothing OrElse Not polyline.Closed Then
                    acEd.WriteMessage(vbLf & "La entidad seleccionada no es una polilínea cerrada.")
                    Return
                End If

                Dim centroid As Point3d = CalculateCentroid(polyline, trans)
                acEd.WriteMessage(vbLf & "Centroide: X={0}, Y={1}, Z={2}", centroid.X, centroid.Y, centroid.Z)

                Dim closestPoint As Point3d = polyline.GetClosestPointTo(centroid, Vector3d.ZAxis, False)
                acEd.WriteMessage(vbLf & "Punto más cercano: X={0}, Y={1}, Z={2}", closestPoint.X, closestPoint.Y, closestPoint.Z)

                ' Crear la línea perpendicular
                Dim perpendicularLine As New Line(centroid, closestPoint)

                ' Extender la línea hasta los bordes de la polilínea
                Dim extendedLine As Line = ExtendLineToPolylineBorders(perpendicularLine, polyline, trans)

                ' Añadir la línea extendida al espacio del modelo
                Dim acBlkTbl As BlockTable = trans.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = trans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(extendedLine)
                trans.AddNewlyCreatedDBObject(extendedLine, True)

                trans.Commit()
            Catch ex As Exception
                acEd.WriteMessage(vbLf & "Ocurrió un error: " & ex.Message)
                trans.Abort()
            End Try
        End Using
    End Sub

    Private Function CalculateCentroid(polyline As Polyline, trans As Transaction) As Point3d
        If trans Is Nothing Then
            Throw New ArgumentNullException(NameOf(trans))
        End If

        Dim area As Double = 0
        Dim centroidX As Double = 0
        Dim centroidY As Double = 0

        ' Calcular el centroide usando la fórmula para polígonos
        For i As Integer = 0 To polyline.NumberOfVertices - 1
            Dim p1 As Point2d = polyline.GetPoint2dAt(i)
            Dim p2 As Point2d = polyline.GetPoint2dAt((i + 1) Mod polyline.NumberOfVertices)

            Dim a As Double = p1.X * p2.Y - p2.X * p1.Y
            area += a
            centroidX += (p1.X + p2.X) * a
            centroidY += (p1.Y + p2.Y) * a
        Next

        area /= 2
        centroidX /= (6 * area)
        centroidY /= (6 * area)

        Return New Point3d(centroidX, centroidY, 0)
    End Function

    Private Function ExtendLineToPolylineBorders(line As Line, polyline As Polyline, trans As Transaction) As Line
        If trans Is Nothing Then
            Throw New ArgumentNullException(NameOf(trans))
        End If

        Dim centroid As Point3d = line.StartPoint
        Dim direction As Vector3d = (line.EndPoint - line.StartPoint).GetNormal()

        ' Extender la línea en ambas direcciones
        Dim extendedLinePos As New LineSegment3d(centroid, centroid + direction * 10000)
        Dim extendedLineNeg As New LineSegment3d(centroid, centroid - direction * 10000)

        ' Encontrar intersecciones con la polilínea
        Dim posIntersection As Point3d? = Nothing
        Dim negIntersection As Point3d? = Nothing

        For i As Integer = 0 To polyline.NumberOfVertices - 1
            Dim p1 As Point3d = polyline.GetPoint3dAt(i)
            Dim p2 As Point3d = polyline.GetPoint3dAt((i + 1) Mod polyline.NumberOfVertices)
            Dim segment As New LineSegment3d(p1, p2)

            Dim posIntersections = extendedLinePos.IntersectWith(segment)
            If posIntersections IsNot Nothing AndAlso posIntersections.Count > 0 Then
                posIntersection = posIntersections(0)
            End If

            Dim negIntersections = extendedLineNeg.IntersectWith(segment)
            If negIntersections IsNot Nothing AndAlso negIntersections.Count > 0 Then
                negIntersection = negIntersections(0)
            End If
        Next

        ' Crear la línea extendida entre los puntos de intersección
        If posIntersection.HasValue AndAlso negIntersection.HasValue Then
            Return New Line(posIntersection.Value, negIntersection.Value)
        ElseIf posIntersection.HasValue Then
            Return New Line(centroid, posIntersection.Value)
        ElseIf negIntersection.HasValue Then
            Return New Line(centroid, negIntersection.Value)
        Else
            ' No se encontraron intersecciones, retornar una línea sin cambios
            Return New Line(centroid, line.EndPoint)
        End If
    End Function

End Module
