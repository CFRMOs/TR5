Imports System.Drawing
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Module PolilineaDivider
    Public Sub AgregarVertice(ByVal pline As Polyline, ByVal point As Point3d)
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

    Public Sub DividirPolilineaEnPunto(ByVal puntoDivisor As Point3d, Id As ObjectId)
        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then
            Return
        End If

        ' Obtener el editor de AutoCAD
        Dim editor As Editor = doc.Editor

        ' Empezar una transacción
        Using trans As Transaction = doc.Database.TransactionManager.StartTransaction()
            Try
                '' Pedir al usuario que seleccione la polilínea a dividir
                'Dim promptResult As PromptEntityResult = editor.GetEntity("Selecciona una polilínea para dividir")
                'If promptResult.Status <> PromptStatus.OK Then
                '    Return
                'End If

                ' Abrir la entidad seleccionada para lectura
                Dim acObj As Object = trans.GetObject(Id, OpenMode.ForWrite)
                If acObj Is Nothing OrElse TypeOf acObj IsNot Polyline Then
                    editor.WriteMessage("La entidad seleccionada no es una polilínea válida.")
                    Return
                End If

                ' Convertir la entidad a una polilínea
                Dim polyline As Polyline = CType(acObj, Polyline)

                ' Agregar un nuevo vértice en la polilínea
                AgregarVertice(polyline, puntoDivisor)

                Dim lisPl As List(Of Polyline) = DividirPolilineaEnPunto(polyline, puntoDivisor)
                ' Guardar los cambios
                polyline.DowngradeOpen()
                trans.Commit()
                editor.WriteMessage("La polilínea ha sido dividida correctamente.")
            Catch ex As Exception
                editor.WriteMessage("Error al dividir la polilínea: " & ex.Message)
            Finally
                trans.Dispose()
            End Try
        End Using
    End Sub
    Private Function DividirPolilineaEnPunto(ByVal polyline As Polyline, ByVal point As Point3d) As List(Of Polyline)
        Dim newPolylines As New List(Of Polyline)()

        ' Obtener el parámetro de la polilínea en el punto especificado
        Dim param As Double = polyline.GetParameterAtPoint(point)

        ' Dividir la polilínea en dos partes en el punto especificado
        If param > 0 AndAlso param < polyline.EndParam Then
            Dim polyline1 As New Polyline()
            Dim polyline2 As New Polyline()

            polyline1.Layer = polyline.Layer
            polyline2.Layer = polyline.Layer

            For i As Integer = 0 To polyline.NumberOfVertices - 1
                Dim vertex As Point3d = polyline.GetPoint3dAt(i)
                Dim Vertex2d As New Point2d(vertex.X, vertex.Y)

                If polyline.GetParameterAtPoint(vertex) < param Then
                    polyline1.AddVertexAt(polyline1.NumberOfVertices, Vertex2d, 0, 0, 0)
                ElseIf polyline.GetParameterAtPoint(vertex) > param Then
                    polyline2.AddVertexAt(polyline2.NumberOfVertices, Vertex2d, 0, 0, 0)
                End If
                If polyline.GetParameterAtPoint(vertex) = param Then
                    polyline2.AddVertexAt(polyline2.NumberOfVertices, Vertex2d, 0, 0, 0)
                    polyline1.AddVertexAt(polyline1.NumberOfVertices, Vertex2d, 0, 0, 0)
                End If
            Next

            newPolylines.Add(polyline1)
            newPolylines.Add(polyline2)
            CrearPolyline_2D(polyline1)
            CrearPolyline_2D(polyline2)
            ' Reemplazar la polilínea original con la primera sección
            polyline.Erase()
        Else
            ' No se puede dividir la polilínea en el punto especificado
            Throw New InvalidOperationException("No se puede dividir la polilínea en el punto especificado.")
        End If

        Return newPolylines
    End Function
    Public Sub CrearPolyline_2D(acPoly As Polyline)
        '' Get the current document and database, and start a transaction
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                         OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                            OpenMode.ForWrite)

            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acPoly)
            acTrans.AddNewlyCreatedDBObject(acPoly, True)
            '' Save the new objects to the database
            acTrans.Commit()
        End Using
    End Sub




End Module

Public Class NewPolilineaDivider
    Public Shared Function DividirPolilineaEnPunto(ByVal puntoDivisor As Point3d, Id As ObjectId) As List(Of Polyline)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        If acDoc Is Nothing Then Return New List(Of Polyline)

        Dim editor As Editor = acDoc.Editor

        ' Iniciar transacción
        Using trans As Transaction = acDoc.Database.TransactionManager.StartTransaction()
            Try
                ' Abrir la polilínea para escribir
                Dim polyline As Polyline = trans.GetObject(Id, OpenMode.ForWrite)
                If polyline Is Nothing Then
                    editor.WriteMessage("La entidad seleccionada no es una polilínea.")
                    Return New List(Of Polyline)
                End If

                ' Dividir la polilínea en dos nuevas
                Dim param As Double = polyline.GetParameterAtPoint(puntoDivisor)

                ' Crear la primera polilínea hasta el punto de división
                Dim polyline1 As Polyline = NuevaPolilinea.CrearNuevaPolilinea(polyline, 0, CInt(Math.Floor(param)), puntoDivisor, param)

                ' Crear la segunda polilínea desde el punto de división hasta el final
                Dim polyline2 As Polyline = NuevaPolilinea.CrearNuevaPolilinea(polyline, CInt(Math.Floor(param)) + 1, polyline.NumberOfVertices - 1, puntoDivisor, param)

                polyline1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0)
                polyline2.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0)

                ' Insertar el punt  o de división en ambas polilíneas
                NuevaPolilinea.InsertarPuntoDivisor(polyline1, polyline2, polyline, puntoDivisor)

                ' Añadir las nuevas polilíneas al dibujo
                InsertarPolilineasEnDibujo(trans, polyline1, polyline2, acDoc.Database)

                ' Borrar la polilínea original
                'polyline.Erase()

                ' Guardar los cambios
                trans.Commit()
                Dim Result As New List(Of Polyline) From {polyline1, polyline2}
                Return Result
                editor.WriteMessage("La polilínea ha sido dividida correctamente.")
            Catch ex As Exception
                editor.WriteMessage("Error al dividir la polilínea: " & ex.Message)
                Return New List(Of Polyline)
            Finally
                trans.Dispose()
            End Try
        End Using
    End Function
    ' Inserta el punto de división en ambas polilíneas en la posición correcta



    Private Shared Sub InsertarPuntoDivisor2(ByVal polyline1 As Polyline, ByVal polyline2 As Polyline, ByVal puntoDivisor As Point3d)
        Dim point2d As New Point2d(puntoDivisor.X, puntoDivisor.Y)

        ' Añadir el punto de división como el último vértice de la primera polilínea
        polyline1.AddVertexAt(polyline1.NumberOfVertices, point2d, 0, 0, 0)

        ' Añadir el punto de división como el primer vértice de la segunda polilínea
        polyline2.AddVertexAt(0, point2d, 0, 0, 0)
    End Sub


    ' Inserta las nuevas polilíneas en el espacio de trabajo del dibujo
    Private Shared Sub InsertarPolilineasEnDibujo(ByVal trans As Transaction, ByVal polyline1 As Polyline, ByVal polyline2 As Polyline, ByVal db As Database)
        Dim blkTbl As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim blkTblRec As BlockTableRecord = trans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

        ' Añadir la primera polilínea al dibujo
        blkTblRec.AppendEntity(polyline1)
        trans.AddNewlyCreatedDBObject(polyline1, True)

        ' Añadir la segunda polilínea al dibujo
        blkTblRec.AppendEntity(polyline2)
        trans.AddNewlyCreatedDBObject(polyline2, True)
    End Sub

End Class
Public Class NuevaPolilinea
    ' Crea una nueva polilínea copiando vértices y bulges desde un índice inicial a un índice final
    Public Shared Function CrearNuevaPolilinea(ByVal polyline As Polyline, ByVal startIdx As Integer, ByVal endIdx As Integer, ByVal puntoDivisor As Point3d, ByVal param As Double) As Polyline
        Dim nuevaPolyline As New Polyline With {
            .Layer = polyline.Layer
        }
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim editor As Editor = acDoc.Editor

        ' Añadir vértices y bulges a la nueva polilínea
        For i As Integer = startIdx To endIdx
            Dim vertex As Point3d = polyline.GetPoint3dAt(i)
            Dim vertex2d As New Point2d(vertex.X, vertex.Y)
            Dim bulge As Double = polyline.GetBulgeAt(i)
            nuevaPolyline.AddVertexAt(nuevaPolyline.NumberOfVertices, vertex2d, bulge, 0, 0)
        Next

        Return nuevaPolyline
    End Function
    ' Inserta el punto de división en ambas polilíneas en la posición correcta
    Public Shared Sub InsertarPuntoDivisor(ByVal polyline1 As Polyline, ByVal polyline2 As Polyline, ByVal polyline As Polyline, ByVal puntoDivisor As Point3d)
        Dim point2d As New Point2d(puntoDivisor.X, puntoDivisor.Y)

        ' Añadir el punto de división como el último vértice de la primera polilínea
        polyline1.AddVertexAt(polyline1.NumberOfVertices, point2d, 0, 0, 0)

        ' Obtener el índice del segmento dividido
        Dim indexLastVertexOriginal As Integer = polyline1.NumberOfVertices - 2  ' Obtener el índice original del penúltimo vértice antes del divisor

        ' Recalcular el bulge para el segmento dividido
        Dim recalculatedBulge1 As Double = NuevaPolilinea.RecalcularBulge(polyline, indexLastVertexOriginal, puntoDivisor, True)

        ' Aplicar el bulge recalculado a polyline1 (último segmento antes del divisor)
        polyline1.SetBulgeAt(indexLastVertexOriginal, recalculatedBulge1)

        ' Añadir el punto de división como el primer vértice de la segunda polilínea
        polyline2.AddVertexAt(0, point2d, 0, 0, 0)

        ' Aplicar el mismo bulge recalculado a polyline2 (primer segmento después del divisor)
        If polyline2.NumberOfVertices > 1 Then
            Dim recalculatedBulge2 As Double = NuevaPolilinea.RecalcularBulge(polyline, indexLastVertexOriginal, puntoDivisor, False)
            polyline2.SetBulgeAt(0, recalculatedBulge2) ' Usamos el bulge recalculado también en el primer segmento de polyline2
        End If
    End Sub
    ' Recalcula el bulge para la nueva sección de la polilínea después de la división
    ' Recalcula el bulge para un segmento dividido
    Public Shared Function RecalcularBulge(ByVal polyline As Polyline, ByVal index As Integer, ByVal puntoDivisor As Point3d, ByVal paraElPrimerTramo As Boolean) As Double
        Dim pt1 As Point3d = polyline.GetPoint3dAt(index)
        Dim pt2 As Point3d = polyline.GetPoint3dAt(index + 1)

        ' Calcular la longitud de la cuerda original
        Dim chordLength As Double = pt1.DistanceTo(pt2)

        ' Calcular la longitud del nuevo segmento según si es el primer o el segundo tramo
        Dim nuevaLongitud As Double
        If paraElPrimerTramo Then
            ' Para el primer tramo (desde el punto inicial hasta el divisor)
            nuevaLongitud = pt1.DistanceTo(puntoDivisor)
        Else
            ' Para el segundo tramo (desde el divisor hasta el segundo punto original)
            nuevaLongitud = puntoDivisor.DistanceTo(pt2)
        End If

        ' Ajustar el bulge proporcionalmente a la nueva longitud de la cuerda
        Dim bulgeOriginal As Double = polyline.GetBulgeAt(index)
        Dim nuevoBulge As Double = bulgeOriginal * (nuevaLongitud / chordLength)

        Return nuevoBulge
    End Function

End Class

