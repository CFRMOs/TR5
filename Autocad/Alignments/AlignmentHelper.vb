Imports Autodesk.AutoCAD.Geometry
Imports System.Linq

Public Class AlignmentHelper
    Private ReadOnly SideMaxDistance As Double = 70
    Public List_ALingName As List(Of String)
    Public List_ALing As New List(Of Alignment) ' From GetAllAlignments(List_ALingName)
    Public Alignment As Alignment
    Public Sub New(Optional ListALing As List(Of Alignment) = Nothing)
        If ListALing Is Nothing Then
            List_ALing = GetAllAlignments(List_ALingName)
        Else
            List_ALing = ListALing
        End If
    End Sub

    Public Sub AlimentByProximida(Optional ByRef Handle As String = Nothing, Optional cSelect As Boolean = False)
        If CLHandle.CheckIfExistHd(Handle) Then CheckAlimentByProximida(CLHandle.GetEntityByStrHandle(Handle), cSelect)
    End Sub

    Public Sub CheckAlimentByProximida(Optional ByRef ent As Autodesk.AutoCAD.DatabaseServices.Entity = Nothing, Optional cSelect As Boolean = False)
        If ent Is Nothing AndAlso cSelect Then CSelectionHelper.GetLayerByEnt(ent)
        If ent Is Nothing Then
            Exit Sub
        End If
        If List_ALing Is Nothing OrElse List_ALing?.Count = 1 Then
            Alignment = List_ALing?(0)
            Exit Sub
        End If
        For Each Al As Alignment In List_ALing
            'chequear que los vertices  de cuneta esten esten a no mas de 30m del alineamiento 
            Dim PL As Autodesk.AutoCAD.DatabaseServices.Polyline = CType(ent, Autodesk.AutoCAD.DatabaseServices.Polyline)
            Dim Offset As Double = 0
            For Each pt As Point2d In CollectPLPoints(PL)
                Dim vr As New Point3d(pt.X, pt.Y, 0)
                Dim Station As Double = 0
                Dim Elevation As Double = 0
                Dim Side As String = String.Empty
                CStationOffsetLabel.GETStationByPoint(Al, vr, Station, Elevation, Offset, Side)
                If Math.Abs(Offset) > SideMaxDistance Then
                    'Alignment = Al
                    GoTo NextAL
                End If
            Next

            If Math.Abs(Offset) < SideMaxDistance AndAlso CheckApprovedStationEquation(Al, ent) Then
                Alignment = Al
                If cSelect Then
                    SelectByEntity(CLHandle.GetEntityByStrHandle(Al.Handle.ToString))
                End If
                Exit For
            End If
NextAL:
        Next
    End Sub

    Public Function CheckApprovedStationEquation(AL As Alignment, Ent As Autodesk.AutoCAD.DatabaseServices.Entity) As Boolean
        ' Convertir la entidad en una polilínea.
        Dim PL As Autodesk.AutoCAD.DatabaseServices.Polyline = CType(Ent, Autodesk.AutoCAD.DatabaseServices.Polyline)

        ' Obtener los puntos inicial y final de la polilínea.
        Dim StartPoint As Point2d = PL.GetPoint2dAt(0)
        Dim EndPoint As New Point2d(PL.EndPoint.X, PL.EndPoint.Y)

        ' Convertir los puntos en Point3d para trabajar con ellos en el alineamiento.
        Dim StartVr As New Point3d(StartPoint.X, StartPoint.Y, 0)
        Dim EndVr As New Point3d(EndPoint.X, EndPoint.Y, 0)

        ' Variables para almacenar estaciones y offsets.
        Dim StartStation As Double = 0
        Dim EndStation As Double = 0
        Dim Offset As Double = 0
        Dim Elevation As Double = 0
        Dim Side As String = String.Empty

        ' Obtener la estación inicial y final de la polilínea en el alineamiento.
        CStationOffsetLabel.GETStationByPoint(AL, StartVr, StartStation, Elevation, Offset, Side)
        CStationOffsetLabel.GETStationByPoint(AL, EndVr, EndStation, Elevation, Offset, Side)


        ' Verificar si la polilínea está completamente dentro del rango aprobado por las ecuaciones.

        If IsPolylineWithinApprovedStationEquation(AL, StartStation, EndStation) AndAlso Not (EndStation = 0 And StartStation = 0) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CheckCunetasAlignment(ByVal PL As Autodesk.AutoCAD.DatabaseServices.Polyline) As Alignment

        ' Variables para almacenar estaciones y offsets.


        Dim DicOffS As New Dictionary(Of Double, Alignment)

        For Each Al As Alignment In List_ALing

            ' Obtener los puntos inicial y final de la polilínea.
            Dim StartPoint As Point2d = PL.GetPoint2dAt(0)
            Dim EndPoint As New Point2d(PL.EndPoint.X, PL.EndPoint.Y)

            ' Convertir los puntos en Point3d para trabajar con ellos en el alineamiento.
            Dim StartVr As New Point3d(StartPoint.X, StartPoint.Y, 0)
            Dim EndVr As New Point3d(EndPoint.X, EndPoint.Y, 0)

            ' Variables para almacenar estaciones y offsets.
            Dim StartStation As Double = 0
            Dim EndStation As Double = 0
            Dim Offset1 As Double = 0
            Dim Offset2 As Double = 0
            Dim Elevation As Double = 0
            Dim Side As String = String.Empty

            ' Obtener la estación inicial y final de la polilínea en el alineamiento.
            CStationOffsetLabel.GETStationByPoint(Al, StartVr, StartStation, Elevation, Offset1, Side)
            'Dim ptSStation As Point3d = CStationOffsetLabel.GetPoint3dByStation(StartStation, Al)

            CStationOffsetLabel.GETStationByPoint(Al, EndVr, EndStation, Elevation, Offset1, Side)

            If IsPolylineWithinApprovedStationEquation(Al, StartStation, EndStation) Then
                If Not (EndStation = 0 And StartStation = 0) AndAlso
                    Math.Abs(Offset1) < SideMaxDistance AndAlso
                    Math.Abs(Offset2) < SideMaxDistance Then
                    Return Al
                End If
            Else
                'crear un metodo para
                If Not (EndStation = 0 And StartStation = 0) AndAlso
                                         Math.Abs(Offset1) < SideMaxDistance AndAlso
                                             Math.Abs(Offset2) < SideMaxDistance Then
                    If Not DicOffS.ContainsKey(Offset2) Then DicOffS.Add(Offset2, Al)
                End If
            End If
        Next
        If DicOffS.Count <> 0 Then
            'devolver el aligment con el offset mas cercano
            Dim k As Double = DicOffS.Keys.Min
            Return DicOffS.Values(k)
        End If
        Return Nothing
    End Function

    ' Función auxiliar para verificar si ambas estaciones están dentro del rango aprobado.
    Private Function IsPolylineWithinApprovedStationEquation(ByVal Al As Alignment, ByVal StartStation As Double, ByVal EndStation As Double) As Boolean
        Dim startInRange As Boolean = False
        Dim endInRange As Boolean = False
        SelectByEntity(CType(Al, Autodesk.AutoCAD.DatabaseServices.Entity))
        CStationOffsetLabel.GetMxMnOPPL(StartStation, EndStation)
        If Al.StationEquations.Count = 0 Then Return True
        ' Comprobar si ambas estaciones están dentro del mismo tramo delimitado por las ecuaciones.

        'qemax y qemin
        ' Obtener el valor mínimo y máximo de RawStationBack usando LINQ
        Dim minRawStationBack As Double = Al.StationEquations.Select(Function(eq) eq.RawStationBack).Min()
        Dim maxRawStationBack As Double = Al.StationEquations.Select(Function(eq) eq.RawStationBack).Max()

        If Al.StationEquations.Count = 1 Then
            If StartStation >= Al.StartingStation AndAlso StartStation <= minRawStationBack Then
                startInRange = True
            End If
        Else
            If StartStation >= minRawStationBack Then
                startInRange = True
            End If
        End If

        If Al.StationEquations.Count = 1 AndAlso startInRange Then Return True

        If EndStation <= Al.EndingStation AndAlso EndStation <= maxRawStationBack Then
            endInRange = True
        End If

        ' Ambas estaciones deben estar dentro del rango aprobado para que se considere válido.
        Return startInRange AndAlso endInRange
    End Function

    Public Sub CheckAlignmentWithStationEquation(ByRef CAcadHelp As ACAdHelpers, Optional ent As Autodesk.AutoCAD.DatabaseServices.Entity = Nothing)

        If ent Is Nothing Then CSelectionHelper.GetLayerByEnt(ent)

        ' Chequeo de proximidad a un alineamiento, tomando en cuenta las ecuaciones de estación.
        For Each Al As Alignment In CAcadHelp.List_ALing
            ' Convertir la entidad en una polilínea.
            Dim PL As Autodesk.AutoCAD.DatabaseServices.Polyline = CType(ent, Autodesk.AutoCAD.DatabaseServices.Polyline)

            ' Obtener los puntos inicial y final de la polilínea.
            Dim StartPoint As Point2d = PL.GetPoint2dAt(0)
            Dim EndPoint As New Point2d(PL.EndPoint.X, PL.EndPoint.Y)

            ' Convertir los puntos en Point3d para trabajar con ellos en el alineamiento.
            Dim StartVr As New Point3d(StartPoint.X, StartPoint.Y, 0)
            Dim EndVr As New Point3d(EndPoint.X, EndPoint.Y, 0)

            ' Variables para almacenar estaciones y offsets.
            Dim StartStation As Double = 0
            Dim EndStation As Double = 0
            Dim Offset As Double = 0
            Dim Elevation As Double = 0
            Dim Side As String = String.Empty

            ' Obtener la estación inicial y final de la polilínea en el alineamiento.
            CStationOffsetLabel.GETStationByPoint(Al, StartVr, StartStation, Elevation, Offset, Side)
            'Dim ptSStation As Point3d = CStationOffsetLabel.GetPoint3dByStation(StartStation, Al)

            CStationOffsetLabel.GETStationByPoint(Al, EndVr, EndStation, Elevation, Offset, Side)


            ' Verificar si la polilínea está completamente dentro del rango aprobado por las ecuaciones.
            If CheckApprovedStationEquation(Al, ent) Then
                CAcadHelp.Alignment = Al
                Exit For ' Si se encuentra un alineamiento válido, salir del bucle.
            End If
        Next
    End Sub

    Public Sub CheckByStations(entity As Autodesk.AutoCAD.DatabaseServices.Entity, side As String, startStation As Double, endStation As Double, Optional Tolerancia As Double = 0.001)
        Alignment = DetermAlignmentByStations(entity:=entity, side:=side, startStation:=startStation, endStation:=endStation, Tolerancia)
    End Sub

    'crear una funcion que con con los estacionamientos y entidad y lado determine Cual es el alignment correspondiente 
    ''' <summary>
    ''' Determines the correct alignment from List_ALing based on the provided entity, side, and stations.
    ''' </summary>
    ''' <param name="entity">The AutoCAD entity, such as a polyline.</param>
    ''' <param name="side">The side specified, "Left" or "Right".</param>
    ''' <param name="startStation">The start station provided for comparison.</param>
    ''' <param name="endStation">The end station provided for comparison.</param>
    ''' <returns>
    ''' The corresponding Alignment if found, or Nothing if no suitable alignment is found.
    ''' </returns>
    Public Function DetermAlignmentByStations(entity As Autodesk.AutoCAD.DatabaseServices.Entity, side As String, startStation As Double, endStation As Double, Optional Tolerancia As Double = 0.001) As Alignment
        ' Convert the entity to a Polyline
        Dim polyline As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(entity, Autodesk.AutoCAD.DatabaseServices.Polyline)
        If polyline Is Nothing Then
            Console.WriteLine("The entity is not a valid polyline.")
            Return Nothing
        End If

        ' Iterate over all alignments in List_ALing
        For Each alignment As Alignment In List_ALing
            ' Get the start and end points of the polyline
            Dim startPoint As Point2d = polyline.GetPoint2dAt(0)
            Dim endPoint As Point2d = polyline.GetPoint2dAt(polyline.NumberOfVertices - 1)

            ' Convert points to Point3d for compatibility with alignment methods
            Dim startVr As New Point3d(startPoint.X, startPoint.Y, 0)
            Dim endVr As New Point3d(endPoint.X, endPoint.Y, 0)

            ' Variables to store calculated stations and offsets
            Dim calculatedStartStation As Double = 0
            Dim calculatedEndStation As Double = 0
            Dim offset1 As Double = 0
            Dim offset2 As Double = 0
            Dim elevation As Double = 0
            Dim actualSide As String = String.Empty

            ' Get stations and offsets for the start and end points of the polyline
            CStationOffsetLabel.GETStationByPoint(alignment, startVr, calculatedStartStation, elevation, offset1, actualSide)
            CStationOffsetLabel.GETStationByPoint(alignment, endVr, calculatedEndStation, elevation, offset2, actualSide)

            ' Check if the given stations match the calculated stations within a reasonable tolerance
            If Math.Abs(calculatedStartStation - startStation) < Tolerancia AndAlso Math.Abs(calculatedEndStation - endStation) < Tolerancia Then
                ' Check if the side matches
                Dim determinedSide As String = GetSideFromOffsets(offset1, offset2)
                If String.Equals(determinedSide, side, StringComparison.OrdinalIgnoreCase) Then
                    ' Return the matching alignment
                    Return alignment
                End If
            End If
        Next

        ' Return Nothing if no suitable alignment is found
        Return Nothing
    End Function

    ''' <summary>
    ''' Determines the side ("Left" or "Right") based on the offsets.
    ''' </summary>
    ''' <param name="offset1">The first offset.</param>
    ''' <param name="offset2">The second offset.</param>
    ''' <returns>The side as "Left", "Right", or "Error" if the alignment intersects the entity.</returns>
    Private Function GetSideFromOffsets(offset1 As Double, offset2 As Double) As String
        If offset1 < 0 AndAlso offset2 < 0 Then
            Return "Left"
        ElseIf offset1 > 0 AndAlso offset2 > 0 Then
            Return "Right"
        Else
            Return "Error" ' Indicates the alignment intersects or does not align properly with the entity
        End If
    End Function

End Class