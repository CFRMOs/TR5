Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil
Imports AlignmentLabel = Autodesk.Civil.DatabaseServices.Label
Imports C3DAlignment = Autodesk.Civil.DatabaseServices.Alignment
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
'AlignmentLabelHelper.GetLabelForLineOrCurve
Public Class AlignmentLabelHelper
    'funccio para identificar una label de una estacion o proxima a esta si existe y seleccionarla 
    Public Shared Sub CFindLabelStation(Station As Double, ByRef CAcadHelp As ACAdHelpers)
        FindLabelStation(Station, CAcadHelp.Alignment, CAcadHelp.ThisDrawing)
    End Sub
    Public Shared Sub FindLabelStation(Station As Double, Alignment As Alignment, ThisDrawing As Document)
        ' Get station label
        Dim stationLabel As AlignmentLabel = AlignmentLabelHelper.GetLabelAtStation(Station, Alignment)

        If stationLabel IsNot Nothing Then
            AcadZoomManager.SelectedZoom(stationLabel.Handle.ToString(), ThisDrawing)
        Else
            CStationOffsetLabel.GotoStation(Station, Alignment, ThisDrawing)
        End If
    End Sub

    ' Función que obtiene la ubicación de las etiquetas de estación
    Public Shared Function GetStationLabelLocation() As List(Of Point3d)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim editor As Editor = doc.Editor
        Dim db As Database = doc.Database
        Dim labelLocations As New List(Of Point3d)

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

            For Each objId As ObjectId In btr
                Dim entity As Entity = TryCast(trans.GetObject(objId, OpenMode.ForRead), Entity)
                If TypeOf entity Is AlignmentLabel Then
                    Dim alignmentLabel As AlignmentLabel = CType(entity, AlignmentLabel)
                    Dim location As Point3d = alignmentLabel.AnchorInfo.Location
                    labelLocations.Add(location)
                End If
            Next
            trans.Commit()
        End Using

        Return labelLocations
    End Function

    ' Función que obtiene una etiqueta de estación en una estación específica
    Public Shared Function GetLabelAtStation(station As Double, c3d_Alignment As C3DAlignment) As AlignmentLabel
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

            For Each objId As ObjectId In btr
                Dim entity As Entity = TryCast(trans.GetObject(objId, OpenMode.ForRead), Entity)
                If TypeOf entity Is AlignmentLabel Then
                    Dim alignmentLabel As AlignmentLabel = CType(entity, AlignmentLabel)
                    Dim labelPoint As Point3d = alignmentLabel.AnchorInfo.Location
                    Dim labelStation As Double = GETStationByP(c3d_Alignment, labelPoint)

                    ' Contar los dígitos decimales de station
                    Dim decimalDigits As Integer = CountDecimalDigits(station)

                    ' Redondear labelStation al mismo número de dígitos decimales que station
                    labelStation = Math.Round(labelStation, decimalDigits)

                    If labelStation = station Then
                        Return alignmentLabel
                    End If
                End If
            Next
            trans.Commit()
        End Using

        Return Nothing
    End Function
    Public Shared Function GetLabelForLineOrCurve(lineId As ObjectId) As GeneralSegmentLabel
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

            ' Obtener la entidad de línea o curva (Polyline, FeatureLine, etc.)
            Dim lineEntity As Entity = TryCast(trans.GetObject(lineId, OpenMode.ForRead), Entity)
            If lineEntity Is Nothing Then
                doc.Editor.WriteMessage(vbLf & "No se pudo encontrar la línea o curva con el ObjectId proporcionado.")
                Return Nothing
            End If

            ' Recorrer todas las entidades en el espacio de modelo para buscar etiquetas de segmentos
            For Each objId As ObjectId In btr
                Dim entity As Entity = TryCast(trans.GetObject(objId, OpenMode.ForRead), Entity)
                If TypeOf entity Is GeneralSegmentLabel Then
                    Dim segmentLabel As GeneralSegmentLabel = CType(entity, GeneralSegmentLabel)
                    'segmentLabel.AcadObject
                    ' Obtener el punto de anclaje de la etiqueta
                    Dim labelPoint As Point3d = segmentLabel.AnchorInfo.Location

                    ' Verificar si el punto de la etiqueta está en la línea o curva
                    If IsPointOnLine(labelPoint, lineEntity) Then
                        ' Retorna la etiqueta si está asociada a la línea o curva
                        'AcadZoomManager.SelectedZoom(segmentLabel.TabName.ToString(), doc)
                        Return segmentLabel
                    End If
                End If
            Next

            trans.Commit()
        End Using

        Return Nothing
    End Function

    ' Función auxiliar para verificar si el punto de la etiqueta está en la línea o curva
    Private Shared Function IsPointOnLine(point As Point3d, lineEntity As Entity) As Boolean
        If TypeOf lineEntity Is Polyline Then
            Dim polyline As Polyline = CType(lineEntity, Polyline)
            Dim closestPoint As Point3d = polyline.GetClosestPointTo(point, False)

            ' Verificar si el punto más cercano está lo suficientemente cerca del punto de la etiqueta
            'Return point.DistanceTo(closestPoint) < Tolerance
            Return point.DistanceTo(closestPoint) < 0.001
        End If

        ' Si es otro tipo de línea, agregar lógica aquí
        Return False
    End Function


    ' Función para contar los dígitos decimales de un número
    Private Shared Function CountDecimalDigits(number As Double) As Integer
        Dim numberStr As String = number.ToString(System.Globalization.CultureInfo.InvariantCulture)
        If numberStr.Contains(".") Then
            Return numberStr.Split("."c)(1).Length
        End If
        Return 0
    End Function

    ' Función que obtiene entidades entre dos etiquetas de estación
    Public Shared Function GetEntitiesBetweenStations(startLabel As AlignmentLabel, endLabel As AlignmentLabel, typeEnt As Type, c3d_Alignment As C3DAlignment) As List(Of Entity)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim entities As New List(Of Entity)

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

            For Each objId As ObjectId In btr
                Dim entity As Entity = TryCast(trans.GetObject(objId, OpenMode.ForRead), Entity)
                If entity IsNot Nothing AndAlso entity.GetType() Is typeEnt Then
                    If IsEntityBetweenLabels(entity, startLabel, endLabel, c3d_Alignment) Then
                        entities.Add(entity)
                    End If
                End If
            Next

            trans.Commit()
        End Using

        Return entities
    End Function

    ' Nueva función que obtiene entidades entre dos etiquetas de estación usando puntos
    Public Shared Function GetEntitiesBetweenStationsUsingPoints(startLabel As AlignmentLabel, endLabel As AlignmentLabel, typeEnt As Type) As List(Of Entity)
        Dim startPoint As Point3d = startLabel.AnchorInfo.Location
        Dim endPoint As Point3d = endLabel.AnchorInfo.Location

        ' Obtener entidades cercanas al punto de inicio
        Dim startEntities As List(Of ObjectId) = GetEntitiesNearPoint(startPoint, 0.01)
        ' Obtener entidades cercanas al punto de finalización
        Dim endEntities As List(Of ObjectId) = GetEntitiesNearPoint(endPoint, 0.01)

        ' Filtrar las entidades que están en ambos listados
        Dim commonEntities As New List(Of Entity)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            For Each startEntityId In startEntities
                If endEntities.Contains(startEntityId) Then
                    Dim entity As Entity = TryCast(trans.GetObject(startEntityId, OpenMode.ForRead), Entity)
                    If entity IsNot Nothing AndAlso entity.GetType() Is typeEnt Then
                        commonEntities.Add(entity)
                    End If
                End If
            Next
            trans.Commit()
        End Using

        Return commonEntities
    End Function

    ' Función que verifica si una entidad está entre dos etiquetas de estación
    Private Shared Function IsEntityBetweenLabels(entity As Entity, startLabel As AlignmentLabel, endLabel As AlignmentLabel, c3d_Alignment As C3DAlignment) As Boolean
        Dim startStation As Double = GETStationByP(c3d_Alignment, startLabel.AnchorInfo.Location)
        Dim endStation As Double = GETStationByP(c3d_Alignment, endLabel.AnchorInfo.Location)
        Dim entityStartStation As Double = GETStationByP(c3d_Alignment, entity.GeometricExtents.MinPoint)

        Return entityStartStation >= startStation AndAlso entityStartStation <= endStation
    End Function

    ' Función que obtiene entidades cercanas a un punto especificado
    Public Shared Function GetEntitiesNearPoint(referencePoint As Point3d, searchRadius As Double) As List(Of ObjectId)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim entities As New List(Of ObjectId)

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

            For Each objId As ObjectId In btr
                Dim entity As Entity = TryCast(tr.GetObject(objId, OpenMode.ForRead), Entity)
                If entity IsNot Nothing AndAlso IsEntityNearPoint(referencePoint, searchRadius, entity) Then
                    entities.Add(objId)
                End If
            Next

            tr.Commit()
        End Using

        Return entities
    End Function

    ' Función que obtiene la estación de un punto especificado
    Public Shared Function GETStationByP(ByRef c3d_Alignment As C3DAlignment, ByRef mLoc As Point3d) As Double
        Dim RefPtStation As Double = 0
        Dim PGL_Elev As Double = 0
        Dim StatOffSet As Double = 0
        Dim mStation_OffSet_Side As String = ""

        ' Utiliza la función GETStationByPoint de UserC_StationOffsetLabel
        CStationOffsetLabel.GETStationByPoint(c3d_Alignment, mLoc, RefPtStation, PGL_Elev, StatOffSet, mStation_OffSet_Side)
        Return RefPtStation
    End Function

    ' Función que verifica si una entidad está cerca de un punto especificado
    Private Shared Function IsEntityNearPoint(referencePoint As Point3d, searchRadius As Double, entity As Entity) As Boolean
        Select Case entity.GetType()
            Case GetType(Polyline)
                Return IsPointNearPolyline(referencePoint, searchRadius, DirectCast(entity, Polyline))
            Case GetType(FeatureLine)
                Return IsPointNearFeatureLine(referencePoint, searchRadius, DirectCast(entity, FeatureLine))
            Case Else
                Return False
        End Select
    End Function

    ' Función que verifica si un punto está cerca de una polilínea
    Private Shared Function IsPointNearPolyline(referencePoint As Point3d, searchRadius As Double, polyline As Polyline) As Boolean
        ' Verificar si el punto está dentro del radio de búsqueda de la polilínea
        For i As Integer = 0 To polyline.NumberOfVertices - 1
            If referencePoint.DistanceTo(polyline.GetPoint3dAt(i)) <= searchRadius Then
                Return True
            End If
        Next

        ' Verificar si la polilínea intercepta el punto
        Return polyline.GetClosestPointTo(referencePoint, Vector3d.ZAxis, False).DistanceTo(referencePoint) <= searchRadius
    End Function

    ' Función que verifica si un punto está cerca de una línea de características
    Private Shared Function IsPointNearFeatureLine(referencePoint As Point3d, searchRadius As Double, featureLine As FeatureLine) As Boolean
        ' Verificar si el punto está dentro del radio de búsqueda de la línea de características
        For Each point As Point3d In featureLine.GetPoints(FeatureLinePointType.AllPoints)
            If referencePoint.DistanceTo(point) <= searchRadius Then
                Return True
            End If
        Next

        ' Verificar si la línea de características intercepta el punto
        Return featureLine.GetClosestPointTo(referencePoint, Vector3d.ZAxis, False).DistanceTo(referencePoint) <= searchRadius
    End Function
    Public Shared Function GetStInline(ByVal Station As Double, ByVal PL As Polyline, ByVal Aling As Alignment) As Point3d
        ' Verificar que el alineamiento no sea nulo
        If Aling Is Nothing Then
            Throw New ArgumentNullException(NameOf(Aling), "El alineamiento no puede ser nulo.")
        End If

        ' Verificar que la polilínea no sea nula
        If PL Is Nothing Then
            Throw New ArgumentNullException(NameOf(PL), "La polilínea no puede ser nula.")
        End If

        ' Declarar variables para coordenadas
        Dim offset As Double = 0
        Dim East As Double = 0
        Dim North As Double = 0

        ' Obtener el punto en el alineamiento para la estación dada
        Aling.PointLocation(Station, offset, East, North)

        ' Convertir coordenadas Este y Norte a un punto 3D en AutoCAD
        Dim pointOnAlignment As New Point3d(East, North, 0)

        ' Obtener los puntos cercanos a la estación dada para calcular el vector tangente
        Dim stationAhead As Double = Station + 0.01 ' Estación un poco más adelante
        Dim EastAhead As Double = 0
        Dim NorthAhead As Double = 0
        Aling.PointLocation(stationAhead, offset, EastAhead, NorthAhead)

        ' Calcular el vector tangente aproximado entre las dos estaciones
        Dim pointAhead As New Point3d(EastAhead, NorthAhead, 0)
        Dim tangentVector As Vector3d = (pointAhead - pointOnAlignment).GetNormal()

        ' Calcular el vector perpendicular
        Dim perpendicularVector As Vector3d = tangentVector.RotateBy(Math.PI / 2, Vector3d.ZAxis)

        ' Crear una línea desde el punto en el alineamiento en la dirección del vector perpendicular
        Dim line As New Line(pointOnAlignment, pointOnAlignment + perpendicularVector)

        ' Encontrar el punto más cercano en la polilínea a esta línea
        Dim closestPoint As Point3d = PL.GetClosestPointTo(line.StartPoint, False)

        Return closestPoint
    End Function



End Class
