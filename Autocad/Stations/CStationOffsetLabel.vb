Imports System.Data.Entity.Core.Metadata.Edm
Imports System.Diagnostics.Eventing.Reader
Imports System.Linq
Imports System.Windows.Media
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.Civil.ApplicationServices
Imports Aligment = Autodesk.Civil.DatabaseServices.Alignment
Public Class CStationOffsetLabel
    Function GetStationLabelInfo(ByVal acTrans As Transaction, C3D_Label As StationOffsetLabel) As List(Of Object)

        Dim explodedObjects As New DBObjectCollection
        Dim Alignment As Aligment
        Dim mLoc As Point3d
        Dim RefPtStation As Double
        Dim OffSet As Double

        Dim Side As String = vbNull
        Dim PGL_Elev As Double

        If Not IsDBNull(C3D_Label) Then

            Alignment = acTrans.GetObject(C3D_Label.FeatureId, OpenMode.ForRead)

            mLoc = C3D_Label.AnchorInfo.Location

            GETStationByPoint(Alignment, mLoc, RefPtStation, PGL_Elev, OffSet, Side)
        End If

        Dim SLData As New List(Of Object) From {
            RefPtStation,
            OffSet,
            PGL_Elev,
            Side,
            mLoc.X,
            mLoc.Y
        }

        Return SLData
    End Function

    ''' <summary>
    ''' Obtiene la estación, elevación, desplazamiento y lado (izquierda, derecha o centro) de un punto en relación con un alineamiento.
    ''' </summary>
    ''' <param name="Alignment">El objeto Alignment de Civil 3D.</param>
    ''' <param name="mLoc">El punto 3D para el cual se calculará la estación y otros parámetros.</param>
    ''' <param name="RefPtStation">Referencia a la estación calculada en el punto proporcionado.</param>
    ''' <param name="PGL_Elev">Referencia a la elevación calculada en el punto proporcionado.</param>
    ''' <param name="OffSet">Referencia al desplazamiento lateral calculado desde el alineamiento.</param>
    ''' <param name="Side">Referencia al lado calculado ("Left", "Right" o "Mid").</param>
    ''' <returns>Devuelve la estación ajustada, considerando ecuaciones de estación si existen.</returns>
    ''' <exception cref="Exception">Devuelve 0 en caso de error.</exception>
    Public Shared Function GETStationByPoint(Alignment As Alignment, ByRef mLoc As Point3d,
                                         ByRef RefPtStation As Double, ByRef PGL_Elev As Double,
                                         ByRef OffSet As Double, ByRef Side As String) As Double

        Dim stationAdjusted As Double

        Try
            ' Calcula la estación y elevación en el punto de referencia
            Alignment.StationOffset(mLoc.X, mLoc.Y, RefPtStation, PGL_Elev)

            ' Calcula el desplazamiento para determinar el lado (izquierda o derecha)
            Alignment.StationOffset(mLoc.X, mLoc.Y, RefPtStation, OffSet)

            ' Ajuste de la estación en caso de que existan ecuaciones
            If Alignment.StationEquations.Count > 0 Then
                ' Recorremos cada ecuación para realizar ajustes, si es necesario
                For Each equation In Alignment.StationEquations
                    ' Cambia 'StationBack' y 'StationAhead' por los nombres correctos de las propiedades según el SDK
                    If RefPtStation >= equation.StationBack Then
                        ' Ajusta la estación usando la diferencia entre estaciones adelante y atrás
                        stationAdjusted = RefPtStation + (equation.StationAhead - equation.StationBack)
                        ' Utiliza la estación ajustada solo si cambió
                        RefPtStation = stationAdjusted
                    Else
                        Exit For
                    End If
                Next
            End If
            ' Determina el lado basado en el desplazamiento calculado
            If (OffSet < 0) Then
                Side = "Left"
            ElseIf (OffSet > 0) Then
                Side = "Right"
            Else
                Side = "Mid"
            End If

            Return RefPtStation
        Catch ex As Exception
            Return 0 ' En caso de error, devuelve 0 y captura el error.
        End Try
    End Function


    ' Go to station based on input
    Public Shared Sub CGotoStation(ByVal Station As Double, ByRef CAcadHelp As ACAdHelpers)
        GotoStation(Station, CAcadHelp.Alignment, CAcadHelp.ThisDrawing)
    End Sub
    Public Shared Sub GotoStation(ByVal Station As Double, ByRef Alignment As Aligment, ThisDrawing As Document)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Try


            Dim offset, Eat, Nor As Double

            If Alignment Is Nothing OrElse Alignment.EndingStation < Station Then 'Exit Sub
                Dim tramos As List(Of Alignment) = GetAlignmentTramos(Alignment)
                Exit Sub
            End If

            Alignment.PointLocation(Station, offset, Eat, Nor)

            Dim Diff As Double = 10.0#

            Dim min As New Point3d(Eat - Diff, Nor - Diff, 0)
            Dim max As New Point3d(Eat + Diff, Nor + Diff, 0)

            AcadZoomManager.ZoomToExtends(min, max)

            Dim stationLabel As Label = AlignmentLabelHelper.GetLabelAtStation(Station, Alignment)

            AcadZoomManager.SelectedZoom(stationLabel.Handle.ToString(), ThisDrawing)

        Catch ex As Exception
            acEd.WriteMessage(("Exception: " & ex.Message))
            Exit Sub
        Finally

        End Try
    End Sub
    ''' <summary>
    ''' Identifica los alineamientos que pueden considerarse "tramos" de un alineamiento dado, 
    ''' basándose únicamente en la continuidad de estaciones.
    ''' </summary>
    ''' <param name="mainAlignment">El alineamiento principal para el cual se buscan tramos.</param>
    ''' <returns>Una lista de alineamientos que cumplen con la condición de continuidad de estaciones con el alineamiento principal.</returns>
    Private Shared Function GetAlignmentTramos(ByVal mainAlignment As Alignment) As List(Of Alignment) ', ByVal ThisDrawing As Document
        Dim tramos As New List(Of Alignment)
        Dim tramosNAme As New List(Of String)

        ' Obtener todos los alineamientos del documento
        Dim allAlignments As List(Of Alignment) = GetAllAlignments(tramosNAme)

        'Tolerancias de estacionamientos 
        Dim Tolerancia As Double = 2

        ' Iterar sobre cada alineamiento para encontrar tramos basados en estaciones
        For Each align As Alignment In allAlignments
            ' Saltar el alineamiento principal (no compararlo consigo mismo)
            If align.Handle = mainAlignment.Handle Then
                Continue For
            End If

            ' Verificar si el alineamiento actual es un tramo:
            ' - Si el alineamiento comienza donde termina el alineamiento principal
            ' - O si el alineamiento termina donde comienza el alineamiento principal
            If Math.Abs(align.StartingStation - mainAlignment.EndingStation) <= Tolerancia OrElse
                Math.Abs(align.EndingStation - mainAlignment.StartingStation) <= Tolerancia Then
                tramos.Add(align)
            End If
        Next

        Return tramos
    End Function

    Public Shared Function GetPoint3dByStation(ByVal Station As Double, Alignment As Aligment) As Point3d
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Try


            Dim offset, Eat, Nor, Elev As Double

            If Alignment Is Nothing Then
                Return New Point3d(0, 0, 0)
            End If

            Alignment.PointLocation(Station, offset, Eat, Nor)
            'calcular la elevacion para conformar el point3d
            Alignment.StationOffset(Eat, Nor, Station, Elev)
            Return New Point3d(Eat, Nor, Elev)
        Catch ex As Exception
            acEd.WriteMessage(("Exception: " & ex.Message))
            Return New Point3d(0, 0, 0)
        Finally

        End Try
    End Function


    ' Función para verificar si una entidad es soportada para el analisis
    Public Shared Function IsSupportedEntity(entity As Entity) As Boolean
        Dim supportedTypes As String() = {"FeatureLine", "Polyline", "Polyline2d", "Polyline3d", "Parcel", "TinSurface"}
        Return supportedTypes.Contains(TypeName(entity))
    End Function

    ' Función para procesar una polilínea y extraer información relevante
    ' Process polyline entities
    Public Shared Sub ProcessPolyline(Alignment As Alignment, ByVal PL As Polyline, ByRef StartPT As Point3d, ByRef EndPT As Point3d, ByRef Len As Double, ByRef Area As Double, ByRef StPtStation As Double, ByRef EDPtStation As Double, ByRef Side1 As String, ByRef Side2 As String)
        If PL.Closed Then
            GetMxMnBorder(PL, Alignment, StPtStation, EDPtStation, Side1)
            Side2 = Side1
        Else
            StartPT = PL.StartPoint
            EndPT = PL.EndPoint
            GetexStation(Alignment, StPtStation, EDPtStation, StartPT, EndPT, Side1, Side2)
        End If
        Len = PL.Length
        Area = PL.Area
    End Sub
    ' Get station information by points
    Public Shared Sub GetexStation(ByVal Alignment As Alignment, ByRef StPtStation As Double, ByRef EDPtStation As Double, StartPT As Point3d, EndPT As Point3d, ByRef STStation_OffSet_Side As String, ByRef EDStation_OffSet_Side As String)

        Dim StPGL_Elev, StStatOffSet, EDPGL_Elev, EDStatOffSet As Double

        GETStationByPoint(Alignment, StartPT, StPtStation, StPGL_Elev, StStatOffSet, STStation_OffSet_Side)
        GETStationByPoint(Alignment, EndPT, EDPtStation, EDPGL_Elev, EDStatOffSet, EDStation_OffSet_Side)
        GetMxMnOPPL(StPtStation, EDPtStation)
    End Sub
    Public Shared Sub GetMxMnBorder(PolyL As Polyline, c3d_Alignment As Alignment, ByRef MXstation As Double, ByRef Mnstation As Double, Optional ByRef mStation_OffSet_Side As String = vbNullString)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim acEd As Editor = doc.Editor

        Dim iResult As New List(Of Object)

        Dim RefPtStation As Double
        Dim PGL_Elev As Double
        Dim StatOffSet As Double
        'Dim Side As String = vbNullString

        Dim acPts As Point2dCollection = CollectPLPoints(PolyL)

        Dim StationList As New List(Of Double)

        For Each Pt3d As Point2d In acPts
            GETStationByPoint(c3d_Alignment, New Point3d(Pt3d.X, Pt3d.Y, 0), RefPtStation, PGL_Elev, StatOffSet, mStation_OffSet_Side)
            StationList.Add(RefPtStation)
        Next

        'Calculo de Maximos y minimos de los estacionamientos definidos por los puntos de los bordes y el alineamiento 
        MXstation = StationList.Max()
        Mnstation = StationList.Min()

    End Sub

    Public Shared Sub GetMxMnOPPL(ByRef StPtStation As Double, ByRef EDPtStation As Double)
        Dim LTPT As New List(Of Double) From {
            StPtStation,
            EDPtStation
        }
        StPtStation = LTPT.Min()
        EDPtStation = LTPT.Max()
    End Sub
    ' Función para obtener las estaciones y el lado de una polilínea usando STSL.GETStationByPoint
    Public Shared Sub GetStationByPoint(c3d_Alignment As Alignment, startPt As Point3d, endPt As Point3d, ByRef startStation As Double, ByRef endStation As Double, ByRef startSide As String, ByRef endSide As String)

        Dim startElev As Double
        Dim startOffset As Double
        GETStationByPoint(c3d_Alignment, startPt, startStation, startElev, startOffset, startSide)

        Dim endElev As Double
        Dim endOffset As Double
        GETStationByPoint(c3d_Alignment, endPt, endStation, endElev, endOffset, endSide)

        GetMxMnOPPL(startStation, endStation)
    End Sub
    ''implementacion de ProcessEntity en clases sin accesos a librerias de autocad 



    ''' <summary>
    ''' Identifica los alineamientos que pueden considerarse "tramos" de un alineamiento dado, 
    ''' basándose únicamente en la continuidad de estaciones.
    ''' </summary>
    ''' <param name="HandleAcEnt">Handel de la entidad analizada, la cual puede ser un poligono.</param>
    ''' <param name="Layer">La capa a la que pertenece esta estidad.</param>
    Public Shared Sub StrProcessEntity(ByRef HandleAcEnt As String, ByRef Layer As String, ByRef MinX As Double,
                                       ByRef MinY As Double, ByRef MaxX As Double, ByRef MaxY As Double,
                                       ByRef startStation As Double, ByRef endStation As Double, ByRef Side As String,
                                       ByRef Len As Double, ByRef Area As Double, HandleAligment As String,
                                       Optional ByRef FileName As String = vbNullString,
                                       Optional ByRef FilePath As String = vbNullString,
                                       Optional ByRef Closed As Boolean = False)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        FilePath = doc.Name
        FileName = Right(doc.Name, FilePath.Length - InStrRev(doc.Name, "\"))

        Dim acEnt As Entity = CType(CLHandle.GetEntityByStrHandle(HandleAcEnt), Entity)
        'En la entidades de autocad que pueden ser cerradas o abierta quiero determinar si es una o la otra y devolver falso o verdadero 
        Closed = EntidadEnClosed(acEnt)

        Dim EntAli As Entity = CLHandle.GetEntityByStrHandle(HandleAligment)

        If EntAli Is Nothing Then Exit Sub

        Dim Alignment As Aligment = CType(EntAli, Aligment)

        Dim StartPT, EndPT As Point3d

        Dim Side1 As String = String.Empty
        Dim Side2 As String = String.Empty

        ProcessEntity(acEnt, StartPT, EndPT, startStation, endStation, Side1, Side2, Len, Area, Alignment)

        MinX = StartPT.X
        MinY = StartPT.Y
        MaxX = EndPT.X
        MaxY = EndPT.Y
        Layer = acEnt?.Layer

        If Side1 <> Side2 Then
            Side = "Error Posible cruce entre cuneta y eje de la via"
        Else
            Side = Side1
        End If

    End Sub

    Public Shared Function EntidadEnClosed(ByVal acEnt As Entity) As Boolean
        'en la entidades de autocad que pueden ser cerradas o abierta quiero determinar si es una o la otra y devolver falso o verdadero 

        Select Case TypeName(acEnt)
            Case "Polyline"
                Return DirectCast(acEnt, Polyline).Closed
            Case "Polyline2d"
                Return DirectCast(acEnt, Polyline2d).Closed
            Case "Polyline3d"
                Return DirectCast(acEnt, Polyline3d).Closed
                'Case "FeatureLine"
                'Return DirectCast(acEnt, Polyline3d).Closed
            Case Else
                Return False
        End Select

    End Function

    '' Función para procesar diferentes tipos de entidades y extraer información relevante en relaciona un alignmet 
    Public Shared Sub ProcessEntity(acEnt As Entity, ByRef startPt As Point3d, ByRef endPt As Point3d, ByRef startStation As Double, ByRef endStation As Double, ByRef Side1 As String, ByRef Side2 As String, ByRef Len As Double, ByRef Area As Double, Aligment As Aligment)
        If acEnt Is Nothing Then Exit Sub
        Select Case TypeName(acEnt)
            Case "Polyline"

                Dim PL As Polyline = CType(acEnt, Polyline)
                CStationOffsetLabel.ProcessPolyline(Aligment, PL, startPt, endPt, Len, Area, startStation, endStation, Side1, Side2)

            Case "Polyline2d"

                Dim PL0 As Polyline2d = CType(acEnt, Polyline2d)
                Dim PL As Polyline = CPoLy2dToPL(PL0)
                CStationOffsetLabel.ProcessPolyline(Aligment, PL, startPt, endPt, Len, Area, startStation, endStation, Side1, Side1)

            Case "Polyline3d"

                Dim PL0 As Polyline3d = CType(acEnt, Polyline3d)
                Dim PL As Polyline = CPoLy3dToPL(PL0)
                CStationOffsetLabel.ProcessPolyline(Aligment, PL, startPt, endPt, Len, Area, startStation, endStation, Side1, Side2)

            Case "FeatureLine"

                Dim FL As FeatureLine = CType(acEnt, FeatureLine)
                startPt = FL.StartPoint
                endPt = FL.EndPoint
                Len = FL.Length2D
                Area = FL.Area
                CStationOffsetLabel.GetexStation(Aligment, startStation, endStation, startPt, endPt, Side1, Side2)

            Case "Line"

                Dim PL As Polyline = ConvertLineToPolyline(CType(acEnt, Line))
                CStationOffsetLabel.ProcessPolyline(Aligment, PL, startPt, endPt, Len, Area, startStation, endStation, Side1, Side2)

            Case "Parcel"

                Dim PR As Parcel = CType(acEnt, Parcel)

                ' Crear la polilínea de la parcela
                'CreatePolylineFromParcel(PR, doc.Database, ed)
                Dim PL As Polyline = GetSegmentParcels(acEnt)
                CStationOffsetLabel.ProcessPolyline(Aligment, PL, startPt, endPt, Len, Area, startStation, endStation, Side1, Side2)
                Area = PR.Area

            Case "TinSurface"

                Dim ClSF As New GPOCommands()

                Dim surface As TinSurface = CType(acEnt, TinSurface)

                Dim PolyCol As DBObjectCollection = ClSF.EntExtractBorder(surface.Id)

                For Each obj As Polyline3d In PolyCol
                    Dim PL As Polyline = CPoLy3dToPL(obj)
                    CStationOffsetLabel.ProcessPolyline(Aligment, PL, startPt, endPt, Len, Area, startStation, endStation, Side1, Side2)
                    Area += surface.Area
                Next
            Case Else
                Throw New InvalidOperationException("Unsupported entity type.")
        End Select

        If startStation = 0 Then
            GetexStation(Aligment, startStation, endStation, startPt, endPt, Side1, Side2)
        End If

        GetMxMnOPPL(startStation, endStation)
    End Sub
End Class
