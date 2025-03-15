Imports System.Windows.Forms
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
Imports TinSurface = Autodesk.Civil.DatabaseServices.TinSurface
Public Class DateViewSet
    Public Shared Function ExisteHandle(handle As String, DGView As DataGridView, ByRef Optional Rowresult As DataGridViewRow = Nothing) As Boolean
        ' Verificamos si el handle es una cadena vacía
        If handle = "" Then Return True ' Manejo para un handle en blanco

        ' Recorremos cada fila del DataGridView
        For Each row As DataGridViewRow In DGView.Rows
            ' Verificamos si el valor de la celda en la primera columna (columna de TabName) coincide
            If row.Cells(1).Value IsNot Nothing AndAlso row.Cells(1).Value.ToString() = handle Then
                Rowresult = row
                Return True ' Si se encuentra el TabName, regresamos True

            End If
        Next
        ' Si no se encuentra el TabName, regresamos False
        Return False
    End Function
    Public Sub GetDataByLayer(selectedLayer As String, DataGridView As DataGridView, CAcadHelp As ACAdHelpers)
        If selectedLayer = String.Empty Then Exit Sub
        Dim form As Form = DataGridView.FindForm()
        form.Update()
        DateViewSet.SetViewCellEnt(CAcadHelp.ThisDrawing, CAcadHelp.Alignment, form, selectedLayer, DataGridView)
        form.Show()
        form.Activate()
    End Sub
    Public Shared Sub CAddAnEntity(ByRef CAcadHelp As ACAdHelpers, DataGridView As DataGridView, Optional acEntHandle As String = "", Optional Cancel As Boolean = False)
        ' Si el handle está vacío y Cancel es True, salimos de la subrutina
        If acEntHandle = "" AndAlso Cancel Then Exit Sub

        ' Si el TabName ya existe en el DataGridView, no se agrega nada y salimos
        If acEntHandle <> "" AndAlso ExisteHandle(acEntHandle, DataGridView) Then Exit Sub

        ' Si llegamos aquí, el TabName no existe, así que llamamos a la función para agregar la entidad
        AddAnEntities(CAcadHelp.ThisDrawing, CAcadHelp.Alignment, DataGridView, acEntHandle)
    End Sub


    Public Shared Sub AddAnEntities(ThisDrawing As Document, Alignment As Alignment, DataGridView As DataGridView, Optional acEntHandle As String = "")
        Dim acEnt As Entity = Nothing
        Dim PromptResult As PromptEntityResult = Nothing
        Dim selectedLayer As String = String.Empty

        Dim Justonce As Boolean = False

        If acEntHandle <> "" Then Justonce = True

        ConsTruLData(PromptResult, selectedLayer, acEntHandle)
        If acEntHandle = String.Empty Then Exit Sub
        'DataGridView.Update()
        DataGridViewHelper.StrAddEntityInfo(ThisDrawing, Alignment, acEntHandle, DataGridView)
        'DataGridView.Show()
        acEntHandle = String.Empty

        If PromptResult?.Status = PromptStatus.Cancel Or Justonce Then Exit Sub

        With DataGridView.FindForm
            Do Until PromptResult.Status = PromptStatus.Cancel
                ConsTruLData(PromptResult, selectedLayer, acEntHandle)
                If acEntHandle = String.Empty Then Exit Sub
                '.Update()
                DataGridViewHelper.StrAddEntityInfo(ThisDrawing, Alignment, acEntHandle, DataGridView)
                '.Show()
                acEntHandle = String.Empty

                If PromptResult.Status = PromptStatus.Cancel Then Exit Sub
            Loop
            '.Activate()
        End With
    End Sub
    Public Shared Sub ConsTruLData(ByRef PromptResult As PromptEntityResult, ByRef selectedLayer As String, ByRef acEntHandle As String)

        Dim LData As List(Of (String, String)) = CSelectionHelper.GetDataByEnt(PromptResult:=PromptResult, acEntHandle)
        If LData?.Count <> 0 Then
            selectedLayer = LData(0).Item1
            acEntHandle = LData(1).Item1
        End If

    End Sub

    Public Shared Sub ChangeByLayer(hdString As String, LayerName As String, Optional DGView As DataGridView = Nothing, Optional ByRef CAcadHelp As ACAdHelpers = Nothing) ', Optional ThisDrawing As Document = Nothing, Optional Alignment As Alignment = Nothing)

        ' Exit the subroutine if hdString is empty
        If String.IsNullOrEmpty(hdString) Then Exit Sub

        CLayerHelpers.ChangeLayersAcEnt(hdString, LayerName)

        If DGView Is Nothing Or CAcadHelp Is Nothing Then Exit Sub

        Dim RowsResult As DataGridViewRow = Nothing

        If DateViewSet.StrHandleExist(hdString, DGView, RowsResult) Then

            RowsResult.Selected = True

            DGView.FirstDisplayedScrollingRowIndex = RowsResult.Index

            UpdateDataDGV.StrUpdateSelectedEntityInfo(CAcadHelp.ThisDrawing, hdString, RowsResult, CAcadHelp.Alignment)
        Else
            DataGridViewHelper.StrAddEntityInfo(CAcadHelp.ThisDrawing, CAcadHelp.Alignment, hdString, DGView)
        End If
    End Sub
    Public Shared Sub SetViewCellbyBaseLayer(ThisDrawing As Document, Alignment As Alignment, UFForm As PolylineMG, nombresCapas As String(), DataGridView As DataGridView)
        Dim selectedLayers As List(Of String) = CrearGrupoYCapas("Drenajes Longitudinales", nombresCapas)

        For Each selectedLayer In selectedLayers 'acEnt.Layer
            If UFForm.ComboBox2.SelectedValue = selectedLayer Then
                SetViewCellEnt(ThisDrawing, Alignment, UFForm, selectedLayer, DataGridView) ', selectedType)
            End If
        Next
    End Sub
    Public Shared Sub HandleDataGridView(DataGridView As DataGridView, ThisDrawing As Document, currentColumnIndex As Integer, newRowIndex As Integer)
        ' Ensure newRowIndex is within va--lid range
        If newRowIndex >= 0 AndAlso newRowIndex < DataGridView.Rows.Count Then
            DataGridView.ClearSelection()
            DataGridView.CurrentCell = DataGridView.Rows(newRowIndex).Cells(currentColumnIndex)
            DataGridView.Rows(newRowIndex).Cells(currentColumnIndex).Selected = True

            ' Get the handle value from the new selected row
            If DataGridView.Rows(newRowIndex).Cells("Handle").Value Is Nothing Then Exit Sub
            Dim hdString As String = DataGridView.Rows(newRowIndex).Cells("Handle").Value.ToString()

            'If hdString = vbNullString Then Exit Sub
            AcadZoomManager.SelectedZoom(hdString, ThisDrawing)
        End If
    End Sub
    Public Shared Sub IdDataFEnt(DataGridView As DataGridView)
        Dim acEnt As Entity = Nothing
        Dim RowsResult As DataGridViewRow = Nothing
        Dim selectedLayer As String = CSelectionHelper.GetLayerByEnt(acEnt)
        If DateViewSet.HandleExist(acEnt, DataGridView, RowsResult) Then
            RowsResult.Selected = True
            DataGridView.FirstDisplayedScrollingRowIndex = RowsResult.Index
        End If
    End Sub

    ' Configure DataGridView settings

    Public Shared Sub ConfigureDataGridView(DataGridView As DataGridView, columnFormats As List(Of String))
        DataGridView.ScrollBars = ScrollBars.Both
        DataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        DataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize

        ' Enable sorting for each column
        Dim i As Integer = 0
        For Each column As DataGridViewColumn In DataGridView.Columns
            column.SortMode = DataGridViewColumnSortMode.Automatic
            ' Aplicar formato de celda
            If i < DataGridView.Columns.Count Then
                column.DefaultCellStyle.Format = columnFormats(i)
                i += 1
            End If
        Next
    End Sub
    Public Shared Sub SetViewCellEnt(ThisDrawing As Document, Alignment As Alignment, UFForm As PolylineMG, selectedLayer As String, DataGridView As DataGridView) ', Optional selectedType As String = "")
        ' Obtener el editor activo de AutoCAD
        Dim acDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        'If selectedType = vbNullString Then selectedType = "Polyline"
        ' Recorrer los objetos del modelo
        Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim bt As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim ms As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead) ', BlockTableRecord)
            For Each objId As ObjectId In ms
                Dim acObj As Object = trans.GetObject(objId, OpenMode.ForRead)
                Dim acEntToAdd As Entity
                acEntToAdd = CType(acObj, Entity)
                If TypeOf acObj Is Polyline Or
                    TypeOf acObj Is FeatureLine Or
                    TypeOf acObj Is Line Or
                    TypeOf acObj Is Polyline3d Or
                    TypeOf acObj Is Parcel Or
                    TypeOf acObj Is TinSurface Then
                    ' Verificar si el objeto es una entidad y cumple las condiciones
                    If acEntToAdd.Layer.Equals(selectedLayer) Then 'AndAlso TypeOf obj Is Polyline Then
                        If HandleExist(acEntToAdd, UFForm.CunetasExistentesDGView) = False Then
                            DataGridViewHelper.AddEntityInfo(ThisDrawing, Alignment, acEntToAdd, DataGridView)
                        End If
                    End If
                End If
            Next
            trans.Commit()
            trans.Dispose()
        End Using
    End Sub
    Public Shared Function StrHandleExist(Handle As String, DataGV As DataGridView, ByRef Optional RowsResult As DataGridViewRow = Nothing) As Boolean
        Return HandleExist(CLHandle.GetEntityByStrHandle(Handle), DataGV, RowsResult)
    End Function
    Public Shared Function HandleExist(acEntToAdd As Entity, DataGV As DataGridView, ByRef Optional RowsResult As DataGridViewRow = Nothing) As Boolean
        For Each rows In DataGV.Rows
            If rows.Cells(1).Value IsNot Nothing AndAlso rows.cells(1).Value.ToString() = acEntToAdd.Handle.ToString() Then
                RowsResult = rows
                Return True
                Exit Function
            End If
        Next
        Return False
    End Function

    Public Shared Function HandleRelated(acEntToAdd As Entity, DataGV As DataGridView, ByRef Optional RowsResult As DataGridViewRow = Nothing) As Boolean
        For Each rows In DataGV.Rows
            If rows.Cells(0).Value IsNot Nothing AndAlso rows.cells(0).Value.ToString() = acEntToAdd.Handle.ToString() Then
                RowsResult = rows
                Return True
                Exit Function
            End If
        Next
        Return False
    End Function
    Public Shared Function GetSelectedEntity() As Entity
        Dim promptOptions As New PromptEntityOptions("Seleccionar un Polyline:")
        Dim acEnt As Entity = SelectEntityByType("Polyline", promptOptions)
        Return acEnt
    End Function
End Class

