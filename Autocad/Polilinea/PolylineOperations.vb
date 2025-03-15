Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class PolylineOperations
    <CommandMethod("CreateStripAndOrderVertices")>
    Public Shared Sub CreateStripAndOrderVertices()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Solicitar la selección de las polilíneas
        Dim peo As New PromptEntityOptions(vbNewLine & "Seleccione la polilínea cerrada")
        peo.SetRejectMessage("Debe seleccionar una polilínea.")
        peo.AddAllowedClass(GetType(Polyline), False)
        Dim per As PromptEntityResult = ed.GetEntity(peo)
        If per.Status <> PromptStatus.OK Then Return

        Dim peoOpen As New PromptEntityOptions(vbNewLine & "Seleccione la polilínea abierta")
        peoOpen.SetRejectMessage("Debe seleccionar una polilínea.")
        peoOpen.AddAllowedClass(GetType(Polyline), False)
        Dim perOpen As PromptEntityResult = ed.GetEntity(peoOpen)
        If perOpen.Status <> PromptStatus.OK Then Return

        ' Solicitar el ancho de la franja
        Dim ppo As New PromptDoubleOptions(vbNewLine & "Ingrese el ancho de la franja") With {
            .AllowZero = False,
            .AllowNegative = False
        }
        Dim ppr As PromptDoubleResult = ed.GetDouble(ppo)
        If ppr.Status <> PromptStatus.OK Then Return
        Dim width As Double = ppr.Value

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim closedPoly As Polyline = tr.GetObject(per.ObjectId, OpenMode.ForRead)
            Dim openPoly As Polyline = tr.GetObject(perOpen.ObjectId, OpenMode.ForRead)

            ' Crear los offsets
            Dim offsetPolys As Tuple(Of Polyline, Polyline) = CreateOffsets(openPoly, width)
            Dim offsetPoly1 As Polyline = offsetPolys.Item1
            Dim offsetPoly2 As Polyline = offsetPolys.Item2

            ' Identificar los vértices en un listado por separado
            Dim vertices1 As List(Of Point3d) = IdentifyVertices(closedPoly, offsetPoly1)
            Dim vertices2 As List(Of Point3d) = IdentifyVertices(closedPoly, offsetPoly2)

            ' Identificar puntos de intersección entre la polilínea abierta y la cerrada
            Dim intersectionPoints As List(Of Point3d) = IdentifyIntersectionPoints(openPoly, closedPoly)

            If intersectionPoints.Count = 0 Then
                ed.WriteMessage("No se encontraron puntos de intersección entre las polilíneas.")
                Return
            End If

            ' Ordenar los vértices por separado asegurando que no se intersecten
            Dim orderedVertices1 As List(Of Point3d) = OrderVertices(vertices1, intersectionPoints)
            Dim orderedVertices2 As List(Of Point3d) = OrderVertices(vertices2, intersectionPoints)

            ' Crear las polilíneas de los offsets ordenados sin combinar
            Dim orderedOffsetPolyline1 As Polyline = CreatePolylineClosedOrOpen(orderedVertices1, False)
            Dim orderedOffsetPolyline2 As Polyline = CreatePolylineClosedOrOpen(orderedVertices2, False)

            ' Unir las polilíneas offset en una sola polilínea cerrada
            Dim combinedPolyline As Polyline = CombineOffsetPolylines(orderedOffsetPolyline1, orderedOffsetPolyline2)

            ' Agregar la polilínea resultante al espacio de trabajo
            Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            btr.AppendEntity(combinedPolyline)
            tr.AddNewlyCreatedDBObject(combinedPolyline, True)

            tr.Commit()
        End Using

        ed.WriteMessage("Operación completada.")
    End Sub

    Private Shared Function CreateOffsets(openPoly As Polyline, width As Double) As Tuple(Of Polyline, Polyline)
        Dim offsetPolys1 As DBObjectCollection = openPoly.GetOffsetCurves(width / 2)
        Dim offsetPolys2 As DBObjectCollection = openPoly.GetOffsetCurves(-width / 2)

        Dim offsetPoly1 As Polyline = CType(offsetPolys1(0), Polyline)
        Dim offsetPoly2 As Polyline = CType(offsetPolys2(0), Polyline)

        Return New Tuple(Of Polyline, Polyline)(offsetPoly1, offsetPoly2)
    End Function

    Private Shared Function IdentifyVertices(closedPoly As Polyline, offsetPoly As Polyline) As List(Of Point3d)
        Dim vertices As New List(Of Point3d)

        ' Añadir los vértices del offset que están dentro de la polilínea cerrada
        AddVerticesInside(closedPoly, offsetPoly, vertices)

        ' Añadir puntos de intersección
        AddIntersectionPoints(vertices, offsetPoly, closedPoly)

        Return vertices
    End Function

    Private Shared Sub AddVerticesInside(closedPoly As Polyline, offsetPoly As Polyline, ByRef vertices As List(Of Point3d))
        For i As Integer = 0 To offsetPoly.NumberOfVertices - 1
            Dim pt As Point3d = offsetPoly.GetPoint3dAt(i)
            If IsPointInPolygon(pt, closedPoly) Then
                vertices.Add(pt)
            End If
        Next
    End Sub

    Private Shared Function IsPointInPolygon(point As Point3d, poly As Polyline) As Boolean
        Dim numVerts As Integer = poly.NumberOfVertices
        Dim j As Integer = numVerts - 1
        Dim oddNodes As Boolean = False
        Dim x As Double = point.X
        Dim y As Double = point.Y

        For i As Integer = 0 To numVerts - 1
            Dim verti As Point2d = poly.GetPoint2dAt(i)
            Dim vertj As Point2d = poly.GetPoint2dAt(j)

            If (verti.Y < y And vertj.Y >= y Or vertj.Y < y And verti.Y >= y) And (verti.X <= x Or vertj.X <= x) Then
                If verti.X + (y - verti.Y) / (vertj.Y - verti.Y) * (vertj.X - verti.X) < x Then
                    oddNodes = Not oddNodes
                End If
            End If
            j = i
        Next

        Return oddNodes
    End Function

    Private Shared Sub AddIntersectionPoints(ByRef vertices As List(Of Point3d), offsetPoly As Polyline, closedPoly As Polyline)
        Dim intersectPoints As New Point3dCollection()
        offsetPoly.IntersectWith(closedPoly, Intersect.OnBothOperands, intersectPoints, IntPtr.Zero, IntPtr.Zero)

        For Each pt As Point3d In intersectPoints
            vertices.Add(pt)
        Next
    End Sub

    Private Shared Function IdentifyIntersectionPoints(openPoly As Polyline, closedPoly As Polyline) As List(Of Point3d)
        Dim intersectPoints As New Point3dCollection()
        openPoly.IntersectWith(closedPoly, Intersect.OnBothOperands, intersectPoints, IntPtr.Zero, IntPtr.Zero)

        Dim points As New List(Of Point3d)
        For Each pt As Point3d In intersectPoints
            points.Add(pt)
        Next

        Return points
    End Function

    Public Shared Function OrderVertices(vertices As List(Of Point3d), intersectionPoints As List(Of Point3d)) As List(Of Point3d)
        Dim orderedVertices As New List(Of Point3d)
        If vertices.Count < 3 Then Return vertices

        ' Añadir el punto inicial de los puntos de intersección como inicio de la polilínea
        orderedVertices.Add(intersectionPoints(0))

        ' Ordenar los vértices para que las aristas no se intersecten
        While vertices.Count > 0
            Dim lastVertex As Point3d = orderedVertices(orderedVertices.Count - 1)
            Dim nextVertex As Point3d = vertices(0)
            Dim minDistance As Double = lastVertex.DistanceTo(nextVertex)
            Dim minIndex As Integer = 0

            For i As Integer = 1 To vertices.Count - 1
                Dim distance As Double = lastVertex.DistanceTo(vertices(i))
                If distance < minDistance AndAlso Not IntersectsPolyline(lastVertex, vertices(i), orderedVertices) Then
                    minDistance = distance
                    nextVertex = vertices(i)
                    minIndex = i
                End If
            Next

            orderedVertices.Add(nextVertex)
            vertices.RemoveAt(minIndex)
        End While

        ' Añadir el punto final de los puntos de intersección como final de la polilínea
        If intersectionPoints.Count > 0 Then
            orderedVertices.Add(intersectionPoints(1))
        End If

        Return orderedVertices
    End Function

    Public Shared Function IntersectsPolyline(pt1 As Point3d, pt2 As Point3d, vertices As List(Of Point3d)) As Boolean
        For i As Integer = 0 To vertices.Count - 2
            If IntersectsSegment(pt1, pt2, vertices(i), vertices(i + 1)) Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Shared Function IntersectsSegment(a As Point3d, b As Point3d, c As Point3d, d As Point3d) As Boolean
        ' Verificar si los segmentos [a, b] y [c, d] se intersectan
        Dim denom As Double = ((b.X - a.X) * (d.Y - c.Y)) - ((b.Y - a.Y) * (d.X - c.X))
        If denom = 0 Then Return False

        Dim ua As Double = (((c.X - a.X) * (d.Y - c.Y)) - ((c.Y - a.Y) * (d.X - c.X))) / denom
        Dim ub As Double = (((c.X - a.X) * (b.Y - a.Y)) - ((c.Y - a.Y) * (b.X - a.X))) / denom

        Return ua >= 0 AndAlso ua <= 1 AndAlso ub >= 0 AndAlso ub <= 1
    End Function

    Public Shared Function CreatePolylineClosedOrOpen(vertices As List(Of Point3d), closed As Boolean) As Polyline
        Dim poly As New Polyline()
        For i As Integer = 0 To vertices.Count - 1
            Dim pt As Point3d = vertices(i)
            poly.AddVertexAt(i, New Point2d(pt.X, pt.Y), 0, 0, 0)
        Next
        poly.Closed = closed

        Return poly
    End Function
    ' Función para extender una línea infinitamente y encontrar la intersección con una polilínea
    Public Shared Function FindIntersectionWithPolyline(PtStation As Point3d, vector As Vector3d, polyline As Polyline) As Point3d
        ' Crear una línea que se extiende en la dirección del vector (usaremos una extensión grande en lugar de infinito)
        Dim lineStart As Point3d = PtStation - (vector * 10000) ' Extender hacia atrás 10,000 unidades
        Dim lineEnd As Point3d = PtStation + (vector * 10000) ' Extender hacia adelante 10,000 unidades
        Dim extendedLine As New Line(lineStart, lineEnd)

        ' Iniciar una transacción para buscar la intersección
        Using acTrans As Transaction = polyline.Database.TransactionManager.StartTransaction()
            Try
                ' Encontrar la intersección entre la línea extendida y la polilínea
                Dim intersectionPoints As New Point3dCollection()
                polyline.IntersectWith(extendedLine, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero)

                ' Si se encuentra una intersección, devolver el primer punto de intersección
                If intersectionPoints.Count > 0 Then
                    Return intersectionPoints(0)
                Else
                    ' Si no se encuentra intersección, devolver un punto nulo o un valor por defecto
                    Return Point3d.Origin ' O podrías devolver un valor nulo personalizado
                End If

            Catch ex As Exception
                ' Manejar cualquier error
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error al calcular la intersección: " & ex.Message)
                Return Point3d.Origin ' O manejar el error de otra forma
            Finally
                acTrans.Commit()
            End Try
        End Using
    End Function
    Private Shared Function CombineOffsetPolylines(offsetPoly1 As Polyline, offsetPoly2 As Polyline) As Polyline
        Dim combinedVertices As New List(Of Point3d)

        ' Añadir vértices del primer offset
        For i As Integer = 0 To offsetPoly1.NumberOfVertices - 1
            combinedVertices.Add(offsetPoly1.GetPoint3dAt(i))
        Next

        ' Añadir vértices del segundo offset en orden inverso
        For i As Integer = offsetPoly2.NumberOfVertices - 1 To 0 Step -1
            combinedVertices.Add(offsetPoly2.GetPoint3dAt(i))
        Next

        ' Crear la polilínea combinada
        Dim combinedPolyline As New Polyline()
        For i As Integer = 0 To combinedVertices.Count - 1
            combinedPolyline.AddVertexAt(i, New Point2d(combinedVertices(i).X, combinedVertices(i).Y), 0, 0, 0)
        Next

        combinedPolyline.Closed = True
        Return combinedPolyline
    End Function
End Class
