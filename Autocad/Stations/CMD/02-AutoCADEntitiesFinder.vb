Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports AlignmentLabel = Autodesk.Civil.DatabaseServices.Label
Imports C3DAlignment = Autodesk.Civil.DatabaseServices.Alignment
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity

Public Class AutoCADEntitiesFinder

    <CommandMethod("FindEntitiesNearPoint")>
    Public Sub FindEntitiesNearPoint()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Try
            ' Solicitar al usuario el punto de referencia
            Dim ppr As PromptPointResult = ed.GetPoint("Seleccione el punto de referencia: ")
            If ppr.Status <> PromptStatus.OK Then
                Return
            End If
            Dim referencePoint As Point3d = ppr.Value

            ' Solicitar al usuario el radio de búsqueda
            Dim pdo As New PromptDistanceOptions("Ingrese el radio de búsqueda: ") With {
                .BasePoint = referencePoint,
                .UseBasePoint = True
            }
            Dim pdr As PromptDoubleResult = ed.GetDistance(pdo)
            If pdr.Status <> PromptStatus.OK Then
                Return
            End If
            Dim searchRadius As Double = pdr.Value

            ' Listar entidades encontradas
            Dim entitiesFound As List(Of ObjectId) = AlignmentLabelHelper.GetEntitiesNearPoint(referencePoint, searchRadius)

            ' Mostrar resultados
            ed.WriteMessage($"{entitiesFound.Count} entidades encontradas cerca del punto {referencePoint}:")
            For Each id As ObjectId In entitiesFound
                ed.WriteMessage(vbCrLf & $"Entidad ID: {id}")
            Next
        Catch ex As Exception
            ed.WriteMessage(vbCrLf & $"Error: {ex.Message}")
        End Try
    End Sub

    <CommandMethod("FindPolylinesBetweenStations")>
    Public Sub FindPolylinesBetweenStations()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Try
            ' Solicitar al usuario el alignment
            Dim alignmentResult As PromptEntityResult = ed.GetEntity("Seleccione el alignment: ")
            If alignmentResult.Status <> PromptStatus.OK Then Return

            ' Obtener el objeto alignment
            Dim alignmentId As ObjectId = alignmentResult.ObjectId
            Dim alignment As C3DAlignment = GetAlignmentById(alignmentId)

            ' Solicitar al usuario la estación inicial
            Dim initialStation As Double = PromptForStation("Seleccione la estación inicial: ")
            If initialStation = Double.NaN Then Return

            ' Solicitar al usuario la estación final
            Dim finalStation As Double = PromptForStation("Seleccione la estación final: ")
            If finalStation = Double.NaN Then Return

            ' Obtener etiquetas de estación
            Dim initialLabel As AlignmentLabel = AlignmentLabelHelper.GetLabelAtStation(initialStation, alignment)
            Dim finalLabel As AlignmentLabel = AlignmentLabelHelper.GetLabelAtStation(finalStation, alignment)

            If initialLabel Is Nothing OrElse finalLabel Is Nothing Then
                ed.WriteMessage("No se encontraron etiquetas de estación para las estaciones especificadas." & vbCrLf)
                Return
            End If

            ' Obtener polilíneas entre las estaciones
            Dim polylines As List(Of Entity) = AlignmentLabelHelper.GetEntitiesBetweenStationsUsingPoints(initialLabel, finalLabel, GetType(Polyline))

            ' Mostrar resultados
            ed.WriteMessage($"{polylines.Count} polilíneas encontradas entre las estaciones {initialStation} y {finalStation}:" & vbCrLf)
            For Each polyline As Entity In polylines
                ed.WriteMessage($"Entidad ID: {polyline.ObjectId}" & vbCrLf)
                AcadZoomManager.SelectedZoom(polyline.Handle.ToString(), doc)
            Next
        Catch ex As Exception
            ed.WriteMessage($"Error: {ex.Message}" & vbCrLf)
        End Try
    End Sub

    <CommandMethod("FindStations")>
    Public Sub FindStations()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Try
            ' Solicitar al usuario el alignment
            Dim alignmentResult As PromptEntityResult = ed.GetEntity("Seleccione el alignment: ")
            If alignmentResult.Status <> PromptStatus.OK Then Return

            ' Obtener el objeto alignment
            Dim alignmentId As ObjectId = alignmentResult.ObjectId
            Dim alignment As C3DAlignment = GetAlignmentById(alignmentId)

            ' Solicitar al usuario la estación
            Dim station As Double = PromptForStation("Seleccione la estación: ")
            If station = Double.NaN Then Return

            ' Obtener etiqueta de estación
            Dim stationLabel As AlignmentLabel = AlignmentLabelHelper.GetLabelAtStation(station, alignment)
            If stationLabel Is Nothing Then
                ed.WriteMessage("No se encontró ninguna etiqueta para la estación especificada." & vbCrLf)
                Return
            End If

            'Seleccionar y hacer zoom a la etiqueta
            AcadZoomManager.SelectedZoom(stationLabel.Handle.ToString(), doc)
        Catch ex As Exception
            ed.WriteMessage($"Error: {ex.Message}" & vbCrLf)
        End Try
    End Sub

    ' Función que solicita al usuario ingresar una estación
    Public Shared Function PromptForStation(prompt As String) As Double
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim pdr As PromptDoubleResult = ed.GetDouble(prompt)
        If pdr.Status <> PromptStatus.OK Then
            Return Double.NaN
        End If
        Return pdr.Value
    End Function

    ' Función que obtiene el objeto alignment por su ID
    Private Function GetAlignmentById(alignmentId As ObjectId) As C3DAlignment
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim alignment As C3DAlignment = TryCast(trans.GetObject(alignmentId, OpenMode.ForRead), C3DAlignment)
            trans.Commit()
            Return alignment
        End Using
    End Function

End Class
