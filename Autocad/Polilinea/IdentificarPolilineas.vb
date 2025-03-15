Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity ' Resolviendo la ambigüedad

Public Class IdentificarPolilineas
    ' Declaraciones a nivel de clase
    Private ReadOnly allIntersections As New List(Of IntersectionResult)
    Private ReadOnly entityTypes As New HashSet(Of Type) From {GetType(Polyline), GetType(FeatureLine)}
    Private classifiedEntities As New List(Of Entity)
    Private ReadOnly processedEntities As New HashSet(Of ObjectId)
    'Private ReadOnly Tipos As New List(Of String) From {"CU-T1", "CU-T2", "CU-BO1"}

    <CommandMethod("IdentificarPolilineasPorEntidadOptimizado")>
    Public Sub MainCommandPorEntidadOptimizado()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        ' Seleccionar el alineamiento
        Dim alignmentId As ObjectId = SelectAlignment(ed)
        If alignmentId.IsNull Then Return

        Dim length As Double = 50 ' Longitud de la polilínea temporal
        Dim intervalo As Double = 20.0 ' Intervalo de 20 metros

        Try
            ' Clasificar las entidades por tipo
            classifiedEntities = ClassifyEntitiesByType()

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim alignment As Alignment = TryCast(tr.GetObject(alignmentId, OpenMode.ForRead), Alignment)
                If alignment Is Nothing Then
                    ed.WriteMessage("No se pudo encontrar el alineamiento.")
                    Return
                End If

                ' Iterar sobre las entidades clasificadas
                For Each entity As Entity In classifiedEntities
                    Dim StartPT, EndPT As Point3d
                    Dim Len, Area, StPtStation, EDPtStation As Double
                    Dim STStation_OffSet_Side As String = String.Empty
                    Dim EDStation_OffSet_Side As String = String.Empty

                    Select Case TypeName(entity)
                        Case "Polyline"
                            Dim PL As Polyline = CType(entity, Polyline)
                            CStationOffsetLabel.ProcessPolyline(alignment, PL, StartPT, EndPT, Len, Area, StPtStation, EDPtStation, STStation_OffSet_Side, EDStation_OffSet_Side)
                        Case "FeatureLine"
                            Dim FL As FeatureLine = CType(entity, FeatureLine)
                            StartPT = FL.StartPoint
                            EndPT = FL.EndPoint
                            Len = FL.Length2D
                            Area = FL.Area
                            CStationOffsetLabel.GetexStation(alignment, StPtStation, EDPtStation, StartPT, EndPT, STStation_OffSet_Side, EDStation_OffSet_Side)
                    End Select

                    Dim bestIntersections As New List(Of IntersectionResult)
                    Dim maxIntersections As Integer = 0

                    ' Iterar sobre las estaciones de la entidad cada 20 metros
                    For station As Double = StPtStation To EDPtStation Step intervalo
                        ' Crear la polilínea temporal para la estación actual
                        Dim tempPL As Polyline = CrearTempPL(alignmentId, station, length)
                        If tempPL Is Nothing Then Continue For

                        ' Obtener las polilíneas que intersectan para la estación actual
                        Dim intersections As List(Of IntersectionResult) = GetPLsInter(tempPL, alignmentId)
                        If intersections.Count > maxIntersections Then
                            maxIntersections = intersections.Count
                            bestIntersections = intersections
                        End If
                    Next

                    ' Añadir las mejores intersecciones al resultado global
                    allIntersections.AddRange(bestIntersections)

                    ' Marcar la entidad como procesada
                    processedEntities.Add(entity.ObjectId)
                Next

                tr.Commit()
            End Using
        Catch ex As Exception
            ed.WriteMessage("Error: " & ex.Message)
        End Try

        ' Imprimir los resultados
        PrintIntersections(ed, allIntersections)
    End Sub

    Private Function SelectAlignment(ed As Editor) As ObjectId
        Dim options As New PromptEntityOptions("Seleccione el alineamiento:")
        options.SetRejectMessage("Debe seleccionar un alineamiento.")
        options.AddAllowedClass(GetType(Alignment), False)

        Dim result As PromptEntityResult = ed.GetEntity(options)
        If result.Status = PromptStatus.OK Then
            Return result.ObjectId
        End If
        Return ObjectId.Null
    End Function

    'Private Function SelectPL(ed As Editor) As ObjectId
    '    Dim options As New PromptEntityOptions("Seleccione el Polyline:")
    '    options.SetRejectMessage("Debe seleccionar un Polyline.")
    '    options.AddAllowedClass(GetType(Polyline), False)

    '    Dim result As PromptEntityResult = ed.GetEntity(options)
    '    If result.Status = PromptStatus.OK Then
    '        Return result.ObjectId
    '    End If
    '    Return ObjectId.Null
    'End Function

    Private Function CrearTempPL(alignmentId As ObjectId, station As Double, length As Double) As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim alignment As Alignment = TryCast(tr.GetObject(alignmentId, OpenMode.ForRead), Alignment)
                If alignment Is Nothing Then
                    ed.WriteMessage("No se pudo encontrar el alineamiento.")
                    Return Nothing
                End If

                ' Asegurar que la estación esté dentro del rango del alineamiento
                If station < alignment.StartingStation OrElse station > alignment.EndingStation Then
                    ed.WriteMessage("La estación especificada está fuera del rango del alineamiento.")
                    Return Nothing
                End If

                ' Crear una instancia de CStationOffsetLabel
                Dim refPtStation As Double = station
                Dim PGL_Elev, StatOffSet As Double
                Dim stationOffsetSide As String = String.Empty

                Dim East As Double, Nor As Double

                alignment.PointLocation(station, StatOffSet, East, Nor)

                ' Obtener la ubicación del punto en la estación
                CStationOffsetLabel.GETStationByPoint(alignment, New Point3d(East, Nor, 0), refPtStation, PGL_Elev, StatOffSet, stationOffsetSide)

                Dim pointOnAlignment As New Point3d(East, Nor, StatOffSet)

                ' Verificar si se obtuvo el punto correctamente
                If pointOnAlignment.Equals(Point3d.Origin) Then
                    ed.WriteMessage("No se pudo obtener el punto en la estación especificada.")
                    Return Nothing
                End If

                ' Obtener la dirección del alineamiento en la estación especificada usando dos puntos
                Dim dist As Double = 10.0 ' Distancia para calcular la dirección
                Dim pointBefore As Point3d = alignment.GetPointAtDist(Math.Max(station - dist, alignment.StartingStation))
                Dim pointAfter As Point3d = alignment.GetPointAtDist(Math.Min(station + dist, alignment.EndingStation))
                Dim alignmentDirection As Vector3d = (pointAfter - pointBefore).GetNormal()

                ' Crear la polilínea perpendicular
                Dim perpendicularDirection As Vector3d = alignmentDirection.RotateBy(Math.PI / 2, Vector3d.ZAxis)
                Dim startPoint As Point3d = pointOnAlignment.Add(perpendicularDirection.MultiplyBy(-length / 2))
                Dim endPoint As Point3d = pointOnAlignment.Add(perpendicularDirection.MultiplyBy(length / 2))

                Dim tempPL As New Polyline()
                tempPL.AddVertexAt(0, New Point2d(startPoint.X, startPoint.Y), 0, 0, 0)
                tempPL.AddVertexAt(1, New Point2d(endPoint.X, endPoint.Y), 0, 0, 0)

                ' Añadir la polilínea al espacio de modelo para visualización
                Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                'btr.AppendEntity(tempPL)
                'tr.AddNewlyCreatedDBObject(tempPL, True)
                'tr.Commit()
                Return tempPL
            Catch ex As Exception
                tr.Abort()
                ed.WriteMessage(vbCrLf & ex.Message)
                Return Nothing
            Finally
                tr.Dispose()
            End Try
        End Using
    End Function

    Private Function GetPLsInter(tempPL As Polyline, AlignId As ObjectId) As List(Of IntersectionResult)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Dim result As New List(Of IntersectionResult)

        Try
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim Align As Alignment = TryCast(tr.GetObject(AlignId, OpenMode.ForRead), Alignment)
                Dim indexLeft As Integer = 0
                Dim indexRight As Integer = 0

                For Each AcEnT As Entity In classifiedEntities
                    If IsClosedEnt(AcEnT) = False Then
                        Dim intersectionPoints As New Point3dCollection()
                        tempPL.IntersectWith(AcEnT, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero)
                        If intersectionPoints.Count > 0 Then
                            Dim Station As Double = 0
                            Dim Side As String = GetSide(Align, intersectionPoints(0), Station)
                            If Side = "Left" Then
                                result.Add(New IntersectionResult(AcEnT, indexLeft, Side, AcEnT.Layer))
                                indexLeft += 1
                            ElseIf Side = "Right" Then
                                result.Add(New IntersectionResult(AcEnT, indexRight, Side, AcEnT.Layer))
                                indexRight += 1
                            End If
                        End If
                    End If
                Next
                tr.Commit()
            End Using
        Catch ex As Exception
            ed.WriteMessage("Error: " & ex.Message)
        End Try

        Return result
    End Function

    Private Function ClassifyEntitiesByType() As List(Of Entity)
        Dim classifiedEntities As New List(Of Entity)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Try
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                For Each objId As ObjectId In ms
                    Dim entity As Entity = TryCast(tr.GetObject(objId, OpenMode.ForRead), Entity)
                    If entity IsNot Nothing Then
                        If entityTypes.Contains(entity.GetType()) Then
                            If Not IsClosedEnt(entity) Then
                                classifiedEntities.Add(entity)
                            End If
                        End If
                    End If
                Next
                tr.Commit()
            End Using
        Catch ex As Exception
            ed.WriteMessage("Error: " & ex.Message)
        End Try

        Return classifiedEntities
    End Function

    Public Function IsClosedEnt(AcEnt As Entity) As Boolean
        If TypeOf AcEnt Is FeatureLine Then
            Return CType(AcEnt, FeatureLine).Closed
        ElseIf TypeOf AcEnt Is Polyline Then
            Return CType(AcEnt, Polyline).Closed
        End If
        Return False
    End Function

    Private Sub PrintIntersections(ed As Editor, intersections As List(Of IntersectionResult))
        For Each result In intersections
            ed.WriteMessage($"Entity TabName: {result.Entity.Handle}, Posición Lateral: {result.Index}, Lado: {result.Side}, Layer: {result.Layer}{vbCrLf}")
        Next
    End Sub

    Private Function GetSide(Alignment As Alignment, ByRef StartPT As Point3d, ByRef StPtStation As Double, Optional ByRef OffSet As Double = 0, Optional ByRef StPGL_Elev As Double = 0) As String
        Dim STStation_OffSet_Side As String = String.Empty
        CStationOffsetLabel.GETStationByPoint(Alignment, StartPT, StPtStation, StPGL_Elev, OffSet, STStation_OffSet_Side)
        Return STStation_OffSet_Side
    End Function
End Class

Public Class IntersectionResult
    Public Property Entity As Entity
    Public Property Index As Integer
    Public Property Side As String
    Public Property Layer As String

    Public Sub New(entity As Entity, index As Integer, side As String, layer As String)
        Me.Entity = entity
        Me.Index = index
        Me.Side = side
        Me.Layer = layer
    End Sub
End Class
