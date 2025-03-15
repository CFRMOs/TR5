Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class SplitPolyline
    <CommandMethod("SplitClosedPolyline")>
    Public Sub SplitClosedPolyline()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Seleccionar la polilínea cerrada
        Dim closedPolylinePrompt As New PromptEntityOptions("Seleccione una polilínea cerrada:")
        closedPolylinePrompt.SetRejectMessage("Debe seleccionar una polilínea cerrada.")
        closedPolylinePrompt.AddAllowedClass(GetType(Polyline), False)
        Dim closedPolylineResult As PromptEntityResult = ed.GetEntity(closedPolylinePrompt)

        If closedPolylineResult.Status <> PromptStatus.OK Then
            Return
        End If

        ' Seleccionar la polilínea abierta
        Dim openPolylinePrompt As New PromptEntityOptions("Seleccione una polilínea abierta:")
        openPolylinePrompt.SetRejectMessage("Debe seleccionar una polilínea abierta.")
        openPolylinePrompt.AddAllowedClass(GetType(Polyline), False)
        Dim openPolylineResult As PromptEntityResult = ed.GetEntity(openPolylinePrompt)

        If openPolylineResult.Status <> PromptStatus.OK Then
            Return
        End If

        Using tr As Transaction = doc.TransactionManager.StartTransaction()
            Try
                Dim closedPolyline As Polyline = tr.GetObject(closedPolylineResult.ObjectId, OpenMode.ForWrite)
                Dim openPolyline As Polyline = tr.GetObject(openPolylineResult.ObjectId, OpenMode.ForRead)

                If Not closedPolyline.Closed Then
                    ed.WriteMessage("La polilínea seleccionada no está cerrada.")
                    Return
                End If

                ' Encontrar puntos de intersección
                Dim intersectionPoints As New Point3dCollection()
                closedPolyline.IntersectWith(openPolyline, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero)

                If intersectionPoints.Count < 2 Then
                    ed.WriteMessage("La polilínea abierta no intersecta la polilínea cerrada en al menos dos puntos.")
                    Return
                End If

                ' Añadir vértices en los puntos de intersección si es necesario
                For Each puntoDivisor As Point3d In intersectionPoints
                    AgregarVertice(closedPolyline, puntoDivisor)
                Next

                ' Crear nuevas polilíneas a partir de los puntos de intersección
                Dim polyline1 As Polyline = CreateSubPolyline(closedPolyline, openPolyline, intersectionPoints, True)
                Dim polyline2 As Polyline = CreateSubPolyline(closedPolyline, openPolyline, intersectionPoints, False)

                ' Agregar las nuevas polilíneas al dibujo
                Dim btr As BlockTableRecord = tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite)
                btr.AppendEntity(polyline1)
                tr.AddNewlyCreatedDBObject(polyline1, True)
                btr.AppendEntity(polyline2)
                tr.AddNewlyCreatedDBObject(polyline2, True)

                ' Eliminar la polilínea original
                closedPolyline.Erase()

                tr.Commit()
            Catch ex As Exception
                ed.WriteMessage("Error: " & ex.Message)
            End Try
        End Using
    End Sub

    Private Function CreateSubPolyline(closedPolyline As Polyline, openPolyline As Polyline, intersectionPoints As Point3dCollection, first As Boolean) As Polyline
        If openPolyline Is Nothing Then
            Throw New ArgumentNullException(NameOf(openPolyline))
        End If

        Dim resultPolyline As New Polyline()
        Dim startPoint As Point3d = intersectionPoints(0)
        Dim endPoint As Point3d = intersectionPoints(1)

        Dim addPoint As Action(Of Point3d) = Sub(pt)
                                                 resultPolyline.AddVertexAt(resultPolyline.NumberOfVertices, New Point2d(pt.X, pt.Y), 0, 0, 0)
                                             End Sub

        Dim adding As Boolean = False

        For i As Integer = 0 To closedPolyline.NumberOfVertices - 1
            Dim pt As Point3d = closedPolyline.GetPoint3dAt(i)

            If pt.IsEqualTo(startPoint) Or pt.IsEqualTo(endPoint) Then
                If adding Then
                    addPoint(pt)
                    adding = False
                Else
                    adding = True
                    addPoint(pt)
                End If
            ElseIf adding Then
                addPoint(pt)
            End If
        Next

        If first Then
            addPoint(startPoint)
        Else
            addPoint(endPoint)
        End If

        resultPolyline.Closed = True
        Return resultPolyline
    End Function

    Private Sub AgregarVertice(ByVal pline As Polyline, ByVal point As Point3d)
        If pline Is Nothing Then Exit Sub

        ' Obtener el punto en la curva más cercano al punto especificado
        point = pline.GetClosestPointTo(point, False)

        ' Obtener el parámetro en la curva en el punto más cercano
        Dim parameter As Double = pline.GetParameterAtPoint(point)

        ' Obtener el índice del segmento en el que se encuentra el punto
        Dim index As Integer = CInt(parameter)

        ' No agregar un nuevo vértice si el punto está en un vértice existente
        If parameter = index Then
            Return
        End If

        ' Obtener el bulge del segmento donde se encuentra el punto
        Dim bulge As Double = pline.GetBulgeAt(index)

        ' Crear un plano OCS (Object Coordinate System) para la polilínea
        Dim plane As New Plane(Point3d.Origin, pline.Normal)

        If bulge = 0.0 Then
            ' Segmento lineal
            pline.AddVertexAt(index + 1, point.Convert2d(plane), 0.0, 0.0, 0.0)
        Else
            ' Segmento de arco
            Dim angle As Double = Math.Atan(bulge) ' Cuarto del ángulo total del arco
            Dim angle1 As Double = angle * (parameter - index) ' Cuarto del primer ángulo del arco
            Dim angle2 As Double = angle - angle1 ' Cuarto del segundo ángulo del arco

            ' Agregar el nuevo vértice y establecer su bulge
            pline.AddVertexAt(index + 1, point.Convert2d(plane), Math.Tan(angle2), 0.0, 0.0)

            ' Establecer el bulge del primer segmento de arco
            pline.SetBulgeAt(index, Math.Tan(angle1))
        End If
    End Sub
End Class
