Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Internal

Public Class BoundaryProcessor
    ''' <summary>
    ''' Comando que genera el área común entre dos polilíneas cerradas.
    ''' </summary>
    <CommandMethod("BOUNDARY_INTERSECTION")>
    Public Sub ExecuteBoundaryIntersection()
        ExecuteBoundaryOperation(BooleanOperationType.BoolIntersect)
    End Sub

    ''' <summary>
    ''' Comando que genera el área de la primera polilínea limitada a la segunda.
    ''' </summary>
    <CommandMethod("BOUNDARY_DIFFERENCE")>
    Public Sub ExecuteBoundaryDifference()
        ExecuteBoundaryOperation(BooleanOperationType.BoolSubtract)
    End Sub

    ''' <summary>
    ''' Ejecuta la operación booleana especificada entre dos polilíneas.
    ''' </summary>
    Private Sub ExecuteBoundaryOperation(operationType As BooleanOperationType)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Solicitar selección de la polilínea externa (Polyline1)
                Dim peo1 As New PromptEntityOptions("\nSeleccione la primera polilínea cerrada: ")
                peo1.SetRejectMessage("Solo polilíneas cerradas.")
                peo1.AddAllowedClass(GetType(Polyline), True)
                Dim ppr1 As PromptEntityResult = ed.GetEntity(peo1)
                If ppr1.Status <> PromptStatus.OK Then Return

                ' Solicitar selección de la polilínea límite (Polyline2)
                Dim peo2 As New PromptEntityOptions("\nSeleccione la segunda polilínea cerrada: ")
                peo2.SetRejectMessage("Solo polilíneas cerradas.")
                peo2.AddAllowedClass(GetType(Polyline), True)
                Dim ppr2 As PromptEntityResult = ed.GetEntity(peo2)
                If ppr2.Status <> PromptStatus.OK Then Return

                ' Obtener las polilíneas seleccionadas
                Dim pl1 As Polyline = TryCast(tr.GetObject(ppr1.ObjectId, OpenMode.ForRead), Polyline)
                Dim pl2 As Polyline = TryCast(tr.GetObject(ppr2.ObjectId, OpenMode.ForRead), Polyline)
                If pl1 Is Nothing OrElse pl2 Is Nothing Then Return

                ' Ejecutar la operación booleana solicitada
                Dim resultPolyline As Polyline = ProcessBooleanOperation(pl1, pl2, db, tr, operationType)
                If resultPolyline IsNot Nothing Then
                    ' Agregar la polilínea resultante al dibujo
                    Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    btr.AppendEntity(resultPolyline)
                    tr.AddNewlyCreatedDBObject(resultPolyline, True)
                End If

                tr.Commit()
            Catch ex As Exception
                ed.WriteMessage("\nError: " & ex.Message)
                tr.Abort()
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' Aplica la operación booleana especificada entre dos polilíneas cerradas.
    ''' </summary>
    Private Function ProcessBooleanOperation(pl1 As Polyline, pl2 As Polyline, db As Database, tr As Transaction, operationType As BooleanOperationType) As Polyline
        ' Convertir ambas polilíneas en regiones para operar con ellas
        Dim reg1 As Region = ConvertPolylineToRegion(pl1, db, tr)
        Dim reg2 As Region = ConvertPolylineToRegion(pl2, db, tr)
        If reg1 Is Nothing OrElse reg2 Is Nothing Then Return Nothing

        ' Clonar la región principal para evitar modificar la original
        Dim reg1Clone As Region = reg1.Clone()
        Dim reg2Clone As Region = reg2.Clone()

        ' Aplicar la operación booleana solicitada
        reg1Clone.BooleanOperation(operationType, reg2Clone)

        ' Convertir la región resultante nuevamente en polilínea
        Return ConvertRegionToPolyline(reg1Clone, db, tr)
    End Function

    ''' <summary>
    ''' Convierte una polilínea en una región para operaciones booleanas.
    ''' </summary>
    Private Function ConvertPolylineToRegion(pl As Polyline, db As Database, tr As Transaction) As Region
        If pl Is Nothing Then Throw New ArgumentNullException(NameOf(pl))
        If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
        If tr Is Nothing Then Throw New ArgumentNullException(NameOf(tr))

        ' Descomponer la polilínea en sus segmentos individuales
        Dim curves As New DBObjectCollection()
        pl.Explode(curves)

        ' Crear una región a partir de las curvas obtenidas
        Dim regionCollection As DBObjectCollection = Region.CreateFromCurves(curves)
        If regionCollection.Count > 0 Then
            Return TryCast(regionCollection(0), Region)
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Convierte una región en una polilínea manteniendo la geometría.
    ''' </summary>
    Private Function ConvertRegionToPolyline(region As Region, db As Database, tr As Transaction) As Polyline
        If region Is Nothing Then Throw New ArgumentNullException(NameOf(region))
        If db Is Nothing Then Throw New ArgumentNullException(NameOf(db))
        If tr Is Nothing Then Throw New ArgumentNullException(NameOf(tr))

        ' Descomponer la región en curvas básicas
        Dim curves As New DBObjectCollection()
        region.Explode(curves)

        ' Construir la polilínea a partir de las curvas extraídas
        Dim pl As New Polyline()
        Dim index As Integer = 0
        For Each obj As Object In curves
            Dim line As Line = TryCast(obj, Line)
            If line IsNot Nothing Then
                pl.AddVertexAt(index, New Point2d(line.StartPoint.X, line.StartPoint.Y), 0, 0, 0)
                index += 1
            End If
        Next
        pl.Closed = True ' Asegurar que la polilínea esté cerrada
        Return pl
    End Function
End Class