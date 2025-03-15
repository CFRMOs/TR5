Imports System.Diagnostics
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.Runtime
Imports Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline
Public Class CCunetasAreaPlanta
    <CommandMethod("CMDCrearAreaEnPlanta")>
    Public Sub SCrearAreaEnPlanta()
        Dim pl As Polyline = CrearAreaEnPlanta()
    End Sub
    Public Function CrearAreaEnPlanta() As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Selección de la polilínea
                Dim peoOpen As New PromptEntityOptions(vbNewLine & "Seleccione la polilínea abierta")
                peoOpen.SetRejectMessage("Debe seleccionar una polilínea.")
                peoOpen.AddAllowedClass(GetType(Polyline), False)
                Dim perOpen As PromptEntityResult = ed.GetEntity(peoOpen)

                If perOpen.Status <> PromptStatus.OK Then Return Nothing

                ' Obtener la polilínea seleccionada
                Dim SelectedPolyline As Polyline = CType(acTrans.GetObject(perOpen.ObjectId, OpenMode.ForRead), Polyline)
                'ancho de cunetas
                Dim Layer As String = vbNullString

                Dim Ancho As Double = ObtenerAnchoPorType(SelectedPolyline.Layer, Layer)

                ' Generar offsets
                Dim TPL As Tuple(Of Polyline, Polyline) = CreateOffsets(SelectedPolyline, Ancho)
                If TPL Is Nothing Then
                    ed.WriteMessage("No se pudo generar los offsets correctamente.")
                    Return Nothing
                End If

                ' Combinar las polilíneas offset
                Dim combinedPolyline As Polyline = CombineOffsetPolylines(TPL.Item1, TPL.Item2)

                CLayerHelpers.CreateLayerIfNotExistsRAcEnt(Layer)

                LayerColorChanger.SetGPLYColor(Layer)

                combinedPolyline.Layer = Layer

                ' Agregar la polilínea resultante al dibujo
                Dim btr As BlockTableRecord = acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                btr.AppendEntity(combinedPolyline)

                acTrans.AddNewlyCreatedDBObject(combinedPolyline, True)

                acTrans.Commit()

                Return combinedPolyline

            Catch ex As Exception
                ed.WriteMessage("Error: " & ex.Message)
                Return Nothing
            End Try
        End Using
    End Function
    Private Shared Function ObtenerAnchoPorType(layerName As String, Optional ByRef Layer As String = vbNullString) As Double
        Select Case True
            Case layerName.Contains("CU-BO1")
                Layer = "CU-BO1"
                Return 0.55
            Case layerName.Contains("CU-T1")
                Layer = "CU-T1"
                Return 1.024
            Case layerName.Contains("CU-T2")
                Layer = "CU-T2"
                Return 1.0207
            Case layerName.Contains("ZJ-DR1")
                Layer = "ZJ-DR1"
                Return 0.9
            Case layerName.Contains("ZJ-DR2")
                Layer = "ZJ-DR2"
                Return 1.2
            Case layerName.Contains("ZJ-DR3")
                Layer = "ZJ-DR3"
                Return 1.8
            Case layerName.Contains("TIPO-LIBRITO")
                Layer = "ZJ-DR3"
                Return 0.8
            Case Else
                Return 0.55 ' Si la capa no coincide, devuelve 0
        End Select
    End Function
    Private Function CreateOffsets(openPoly As Polyline, width As Double) As Tuple(Of Polyline, Polyline)
        Try
            Dim offsetPolys1 As DBObjectCollection = openPoly.GetOffsetCurves(width / 2)
            Dim offsetPolys2 As DBObjectCollection = openPoly.GetOffsetCurves(-width / 2)

            ' Verificar que hay elementos en las colecciones antes de acceder a ellos
            If offsetPolys1.Count = 0 OrElse offsetPolys2.Count = 0 Then
                Return Nothing
            End If

            Dim offsetPoly1 As Polyline = TryCast(offsetPolys1(0), Polyline)
            Dim offsetPoly2 As Polyline = TryCast(offsetPolys2(0), Polyline)

            If offsetPoly1 Is Nothing OrElse offsetPoly2 Is Nothing Then
                Return Nothing
            End If

            Return New Tuple(Of Polyline, Polyline)(offsetPoly1, offsetPoly2)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function CombineOffsetPolylines(offsetPoly1 As Polyline, offsetPoly2 As Polyline) As Polyline
        Dim combinedPolyline As New Polyline()
        Dim index As Integer = 0

        ' Añadir vértices del primer offset con su bulge
        For i As Integer = 0 To offsetPoly1.NumberOfVertices - 1
            Dim pt As Point2d = offsetPoly1.GetPoint2dAt(i)
            Dim bulge As Double = offsetPoly1.GetBulgeAt(i)
            combinedPolyline.AddVertexAt(index, pt, bulge, 0, 0)
            index += 1
        Next

        ' Añadir vértices del segundo offset en orden inverso con bulge invertido
        Dim reverseIndex As Integer = offsetPoly2.NumberOfVertices - 1
        For i As Integer = 0 To reverseIndex
            Dim pt As Point2d = offsetPoly2.GetPoint2dAt(reverseIndex - i)
            Dim bulge As Double = offsetPoly2.GetBulgeAt(reverseIndex - i)
            Debug.Assert(index <> 72)
            ' Invertimos el bulge para mantener la dirección correcta del arco
            combinedPolyline.AddVertexAt(index, pt, -bulge, 0, 0)
            index += 1
        Next

        ' Cerrar la polilínea
        combinedPolyline.Closed = True
        Return combinedPolyline
    End Function


End Class
' Clase auxiliar para la línea temporal
Public Class TransientLine
    Implements IDisposable

    Private ReadOnly gfx As TransientManager
    Private ReadOnly line As Line

    Public Sub New(seg As LineSegment3d)
        gfx = TransientManager.CurrentTransientManager
        line = New Line(seg.StartPoint, seg.EndPoint) With {.ColorIndex = 1} ' Rojo
    End Sub

    Public Sub Display()
        gfx.AddTransient(line, TransientDrawingMode.Main, 128, Nothing)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        gfx.EraseTransient(line, Nothing)
    End Sub
End Class