Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.LayerManager
Imports Autodesk.AutoCAD.Runtime

Public Class Class2
    <CommandMethod("DLF")>
    Public Sub DeleteLayerFilter()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ListLayerFilters()

        Try
            ' Obtener los filtros de capa existentes
            ' (añadiremos y los estableceremos de nuevo)

            Dim lft As LayerFilterTree = db.LayerFilters
            Dim lfc As LayerFilterCollection = lft.Root.NestedFilters

            ' Solicitar el índice del filtro a eliminar
            Dim pio As New PromptIntegerOptions(vbCrLf & vbCrLf & "Enter index of filter to delete") With {
                .LowerLimit = 1,
                .UpperLimit = lfc.Count
            }

            Dim pir As PromptIntegerResult = ed.GetInteger(pio)

            If pir.Status <> PromptStatus.OK Then
                Return
            End If

            ' Obtener el filtro seleccionado
            Dim lf As LayerFilter = lfc(pir.Value - 1)

            ' Si es posible eliminarlo, hacerlo
            If Not lf.AllowDelete Then
                ed.WriteMessage(vbCrLf & "Layer filter cannot be deleted.")
            Else
                lfc.Remove(lf)
                db.LayerFilters = lft

                ListLayerFilters()
            End If
        Catch ex As Exception
            ed.WriteMessage(vbCrLf & "Exception: {0}", ex.Message)
        End Try
    End Sub
End Class
