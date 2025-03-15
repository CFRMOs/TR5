Imports System.Linq
Imports System.Windows.Documents
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
'CGPointHelper.AddCGPoint
Public Class CGPointHelper
    Public Shared Function AddCGPoint(Eat As Double, Nor As Double, elevation1 As Double, Optional Description As String = "") As Boolean
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim points As CogoPointCollection = CivilApplication.ActiveDocument.CogoPoints
                Dim PTGlobal As New Point3d(Eat, Nor, elevation1)

                Dim pointId As ObjectId = points.Add(PTGlobal, False)
                Dim cogopoint As CogoPoint = CType(acTrans.GetObject(pointId, OpenMode.ForWrite), CogoPoint)


                cogopoint.RawDescription = Description
                acTrans.Commit()
                Return True
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage($"Exception: {ex.Message}")
                Return False
            End Try
        End Using
    End Function

    'crear una function que devuelva los puntos en el primer grupo de puntos 
    Public Shared Function ObtenerPrimerGrupoDePuntos(gruposDePuntos As List(Of List(Of Point3d))) As List(Of Point3d)
        If gruposDePuntos IsNot Nothing AndAlso gruposDePuntos.Count > 0 Then
            ' Devuelve el primer grupo de puntos
            Return gruposDePuntos(0)
        Else
            ' Si no hay grupos o la lista es nula, devuelve una lista vacía o maneja el error según sea necesario
            Return New List(Of Point3d)()
        End If
    End Function

    Public Shared Function OrderVertices(vertices As List(Of Point3d)) As Dictionary(Of String, Point3d)
        Dim orderedVertices As New Dictionary(Of String, Point3d)
        If vertices Is Nothing OrElse vertices.Count < 3 Then Return orderedVertices

        ' Diccionario para almacenar los vértices y sus claves
        Dim DicVertices As New Dictionary(Of String, Point3d)

        'Order vertices By Easting firt
        vertices = vertices.OrderBy(Function(v) v.X).ToList()


        For i As Integer = 0 To vertices.Count - 1
            DicVertices.Add($"P-{i + 1}", vertices(i))
        Next


        ' Añadir el primer vértice como punto inicial de referencia (IPoint)
        orderedVertices.Add("P-1", DicVertices("P-1"))
        Dim IPoint As Point3d = DicVertices("P-1")
        DicVertices.Remove("P-1")

        ' Contador de vértices para controlar si se puede realizar la segunda verificación
        Dim vertexCount As Integer = orderedVertices.Count

        ' Comenzamos el bucle de ordenación
        While DicVertices.Count > 0
            Dim nearestVertex As Point3d = Nothing
            ' Ordenar los puntos restantes según su distancia al punto de referencia (IPoint)
            ' Obtener los 5 puntos más cercanos ordenados por distancia desde IPoint
            Dim nearestPoints = DicVertices _
                            .OrderBy(Function(kvp) IPoint.DistanceTo(kvp.Value)) _
                             .Take(5) ' Seleccionar los 5 puntos más cercanos

            ' Crear un diccionario para almacenar los puntos más cercanos junto con su clave
            Dim ListNearestPoints As New Dictionary(Of String, Point3d)

            ' Listado secundario para controlar cantidad de vertices cercanos y lograr que el polígono sea cerrado
            If vertexCount >= 2 Then
                ListNearestPoints.Add("P-1", orderedVertices("P-1"))
            End If

            ' Añadir los puntos más cercanos al diccionario
            For Each kvp In nearestPoints
                ListNearestPoints.Add(kvp.Key, kvp.Value) ' Clave y coordenadas del punto
                'P1ListNearestPoints.Add(kvp.Key, kvp.Value) ' También agregarlo a la lista para la verificación del cierre del polígono
            Next
            ' Recorrer el diccionario de puntos más cercanos
            For Each kvp In ListNearestPoints
                ' Imprimir las claves y coordenadas de los puntos más cercanos
                Console.WriteLine($"Clave: {kvp.Key}, Coordenadas: ({kvp.Value.X}, {kvp.Value.Y}, {kvp.Value.Z})")
                ' Verificar intersecciones con los puntos ya ordenados
                If kvp.Key <> "P-1" Then 'kvp.Key <> "P-1" Then
                    If Not IntersectsPolyline(IPoint, kvp.Value, ListNearestPoints.Values.ToList()) Then
                        ' Crear el segmento y añadirlo al modelo de AutoCAD
                        'Dim lista1 As New List(Of Point3d)(New Point3d() {IPoint, kvp.Value})
                        'AddToModal(CrearPL(lista1))
                        'verificacion de los tramos ya verificados en primera instacion con el nuevo tramo para evitar intersecciones 
                        'If Not IntersectsPolyline2(IPoint, kvp.Value, orderedVertices.Values.ToList()) Then
                        nearestVertex = kvp.Value
                            vertexCount += 1 ' Aumentar el número de vértices después de añadir uno nuevo
                            Exit For ' Detener el bucle al encontrar un vértice válido
                        'End If

                    End If
                End If
            Next

            ' Si se encontró un vértice válido que no interfiere
            If nearestVertex <> Point3d.Origin Then
                ' Actualizamos el punto de referencia para la siguiente iteración
                IPoint = nearestVertex

                Dim key = DicVertices.FirstOrDefault(Function(kvp) kvp.Value = nearestVertex).Key
                ' Añadir el vértice al diccionario de vértices ordenados
                orderedVertices.Add(key, nearestVertex)
                ' Remover el vértice ya ordenado
                DicVertices.Remove(key)
            Else
                ' Si no se encontró ningún vértice válido, evitar bucles infinitos
                Console.WriteLine("No se encontró un vértice válido que no cause intersecciones.")
                Exit While
            End If
        End While

        Return orderedVertices
    End Function


    Public Shared Function IntersectsPolyline(pt1 As Point3d, pt2 As Point3d, vertices As List(Of Point3d)) As Boolean
        For i As Integer = 0 To vertices.Count - 2
            For j As Integer = 0 To vertices.Count - 2
                'valores verdaderos indica solapes de segmentos 
                'Dim lista1 As New List(Of Point3d)(New Point3d() {IPoint, kvp.Value})
                'AddToModal(CrearPL(lista1))
                'Dim lista2 As New List(Of Point3d)(New Point3d() {vertices(i), vertices(j)})
                'AddToModal(CrearPL(lista2))

                If IntersectsSegment(pt1, pt2, vertices(i), vertices(j)) Then

                    Return True
                End If
            Next
        Next
        Return False
    End Function
    Public Shared Function IntersectsPolyline2(pt1 As Point3d, pt2 As Point3d, vertices As List(Of Point3d)) As Boolean
        Dim lista1 As New List(Of Point3d)(New Point3d() {pt1, pt2})
        Dim pl1 As Polyline = CrearPL(lista1)
        AddToModal(pl1)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        AcadZoomManager.SelectedZoom(pl1.Handle.ToString(), doc)

        For i As Integer = 0 To vertices.Count - 2
            If i + 1 <= vertices.Count Then
                If IntersectsSegment(pt1, pt2, vertices(i), vertices(i + 1)) Then

                    Dim lista2 As New List(Of Point3d)(New Point3d() {vertices(i), vertices(i + 1)})
                    Dim pl2 As Polyline = CrearPL(lista2)
                    AddToModal(pl2)
                    AcadZoomManager.SelectedZoom(pl1.Handle.ToString(), doc)
                    Return True
                End If
            End If
        Next
        Return False
    End Function
    Public Shared Function IntersectsSegment(a As Point3d, b As Point3d, c As Point3d, d As Point3d) As Boolean
        Try
            ' Calcula el determinante o valor del denominador de las fórmulas de intersección entre segmentos.
            Dim denom As Double = (b.X - a.X) * (d.Y - c.Y) - (b.Y - a.Y) * (d.X - c.X)

            ' Si el denominador es 0, significa que los segmentos son paralelos (o coincidentes).
            If denom = 0 Then Return False

            ' Calcula la posición de intersección relativa para el primer y segundo segmento.
            Dim ua As Double = ((c.X - a.X) * (d.Y - c.Y) - (c.Y - a.Y) * (d.X - c.X)) / denom
            Dim ub As Double = ((c.X - a.X) * (b.Y - a.Y) - (c.Y - a.Y) * (b.X - a.X)) / denom

            ' Verificar si los segmentos se intersectan dentro de sus respectivos límites
            ' Si ambos ua y ub están en el rango [0, 1], entonces los segmentos se intersectan
            If ua >= 0 AndAlso ua <= 1 AndAlso ub >= 0 AndAlso ub <= 1 Then
                ' Verificar si los segmentos se tocan en un extremo
                Dim interseccion As New Point3d(a.X + ua * (b.X - a.X), a.Y + ua * (b.Y - a.Y), a.Z + ua * (b.Z - a.Z))

                ' Si el punto de intersección coincide con los puntos finales de los segmentos, no es una intersección válida
                If interseccion.Equals(b) OrElse interseccion.Equals(c) Then
                    Return False ' Los segmentos solo se tocan en los extremos, no es una intersección completa
                End If
                ' Los segmentos se cruzan de manera válida
                Return True
            End If
            ' No hay intersección
            Return False
        Catch ex As Exception
            ' Manejo de excepciones
            Console.WriteLine($"Error: {ex.Message}")
            Return False
        End Try
    End Function



    Public Shared Function Crearsegmentospl(Sortedpoint As List(Of Point3d)) As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim Polyline As Polyline = PolylineOperations.CreatePolylineClosedOrOpen(Sortedpoint, True)
                ' Agregar la polilínea resultante al espacio de trabajo
                Dim blkTbl As BlockTable = CType(trans.GetObject(doc.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim blkTblRec As BlockTableRecord = CType(trans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                blkTblRec.AppendEntity(Polyline)
                trans.AddNewlyCreatedDBObject(Polyline, True)
                AcadZoomManager.SelectedZoom(Polyline.Handle.ToString(), doc)
                Return Polyline
                ' Completar la transacción
                trans.Commit()
            Catch ex As Exception
                ' Manejar cualquier error
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
                Return Nothing
            Finally
                trans.Dispose()
            End Try
        End Using

    End Function
    Public Shared Sub AddToModal(cPolyline As Polyline)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim blkTbl As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim blkTblRec As BlockTableRecord = CType(trans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                blkTblRec.AppendEntity(cPolyline)

                trans.AddNewlyCreatedDBObject(cPolyline, True)

                'AcadZoomManager.SelectedZoom(cPolyline.TabName.ToString(), doc)

                trans.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
            Finally
                trans.Dispose()
            End Try
        End Using
    End Sub
    ' Función para crear una polilínea a partir de una lista de puntos
    Public Shared Function CrearPL(Points As List(Of Point3d)) As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Crear la polilínea
                Dim cPolyline As Polyline = PolylineOperations.CreatePolylineClosedOrOpen(Points, False)
                ' Completar la transacción
                trans.Commit()
                Return cPolyline
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
                Return Nothing
            Finally
                trans.Dispose()
            End Try
        End Using
    End Function

End Class
