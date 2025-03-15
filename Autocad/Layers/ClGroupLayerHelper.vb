Imports System.Linq
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.LayerManager

Public Class ClGroupLayerHelper
    Public Shared Function CrearGruposCapas(ListMediciones As String(), TiposCunetas As String()) As Dictionary(Of String, String())
        Dim CapasMedicion As String()
        Dim DCapas As New Dictionary(Of String, String())
        For Each m As String In ListMediciones
            CapasMedicion = New String(TiposCunetas.Length - 1) {}
            For i As Integer = 0 To TiposCunetas.Length - 1
                CapasMedicion(i) = m & "-" & TiposCunetas(i)
            Next
            CrearGrupoYCapas(m, CapasMedicion)
            DCapas.Add(m, CapasMedicion)
        Next
        Return DCapas
    End Function
    Public Sub CreateLayerGroup(nombresCapas As String(), LayerFNAme As String)
        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Dim LayFCol As LayerFilterCollection = AcCurD.LayerFilters.Root.NestedFilters

        If LayerGroupExists(LayerFNAme) Then
            CLayerHelpers.GetextraLayer(nombresCapas, LayerFNAme)
            DeleteLayerFilterbyName(LayerFNAme)
        Else
            CLayerHelpers.GetextraLayer(nombresCapas, LayerFNAme)
        End If
        'Crear un LAyer Filter 
        'Dim lf As LayerFilter = New LayerFilter()
        Dim Lg As New LayerGroup With {
            .Name = LayerFNAme
        }

        Dim ListObj As New ObjectIdCollection()


        For Each x As String In nombresCapas
            CLayerHelpers.CreateLayerIfNotExists(x)
        Next

        Dim trans As Transaction = AcCurD.TransactionManager.StartTransaction()

        Using trans
            Try
                Dim LYTable As LayerTable = trans.GetObject(AcCurD.LayerTableId, OpenMode.ForWrite)
                ' Obtener las capas que ya están en el grupo
                Dim existingLayerIds As New List(Of ObjectId)()
                For Each layerId As ObjectId In Lg.LayerIds
                    existingLayerIds.Add(layerId)
                Next

                For Each LYId As ObjectId In LYTable
                    Dim LYObj As LayerTableRecord = trans.GetObject(LYId, OpenMode.ForWrite)

                    If nombresCapas.Contains(LYObj.Name) And Not existingLayerIds.Contains(LYId) Then
                        ListObj.Add(LYId)
                    End If
                Next

                'create a 
                Dim Lft As LayerFilterTree = AcCurD.LayerFilters
                Dim Lfc As LayerFilterCollection = Lft.Root.NestedFilters
                If Lg.LayerIds.Count = 0 Then
                    For Each id As ObjectId In ListObj
                        Lg.LayerIds.Add(id)
                    Next
                    Lfc.Add(Lg)
                    AcCurD.LayerFilters = Lft
                End If
                AcEd.WriteMessage("\n\{0}\group created containing {1} layers.\n", Lg.Name, ListObj.Count)
                trans.Commit()
            Catch ex As Exception
                ' Manejar cualquier error
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
            Finally
                If trans.IsDisposed = False Then
                    trans.Dispose()
                End If
            End Try
        End Using
    End Sub
End Class
