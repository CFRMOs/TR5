Imports System.Linq
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
'CtrolsButtonsUT.VisibilityToggle
Public Class CtrolsButtonsUT
    'btHide y button12
    'metodo para ocultar/mostar una entidad seleccionada en el dgv 
    Public Shared Sub VisibilityToggle(DGWView As DataGridView, toggleVisibility As Boolean)
        Dim hdString As String
        Dim selectedCells As DataGridViewSelectedCellCollection = DGWView.SelectedCells
        For Each cell As DataGridViewCell In selectedCells
            hdString = DGWView.Rows(cell.RowIndex).Cells("Handle").Value.ToString()
            AutoCADUtilities.EntidadVisibility(CLHandle.GetEntityByStrHandle(hdString).Id, toggleVisibility)
        Next
    End Sub
    ' TabName button click event to change layer of selected entity
    Public Shared Sub CambiarLayer(CBox As ComboBox, ComMediciones As ComboBox, TabCtrol As TabControl, CAcadHelp As ACAdHelpers) ' Handles Button2.Click
        Dim selectedLayer As String = Trim(CBox.Text)
        'Dim hdString As String
        Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabCtrol) 'TabCtrol)

        Dim selectedCells As DataGridViewSelectedCellCollection = DGView.SelectedCells

        For Each cell As DataGridViewCell In selectedCells
            'hdString = DGView.Rows(cell.RowIndex).Cells("Handle").Value.ToString()
            Dim Cuneta As New CunetasHandleDataItem With {
                                                            .Handle = DGView.Rows(cell.RowIndex).Cells("Handle").Value.ToString()
                                                            }
            Cuneta.SetPropertiesFromDWG(Cuneta.Handle, Cuneta.AlignmentHDI)

            Cuneta.Layer = CBox.Text


            'CLayerHelpers.ChangeLayersAcEnt(hdString, CBox.Text)

            UpdateDataDGV.UpdateSelectedEntityInfo(CAcadHelp.ThisDrawing,
                                                   Cuneta.Polyline, DGView.Rows(cell.RowIndex),
                                                   Cuneta.Alignment)
            ' Si no existe en el diccionario, buscar en la tabla de Excel
            Dim ExRNGSe As New ExcelRangeSelector()
            Dim RNG As Microsoft.Office.Interop.Excel.Range = ExRNGSe.HandleFromExcelTable(TBName:="CunetasGeneral", handle:=Cuneta.Handle, "Layer")
            If RNG IsNot Nothing Then RNG.Value = Cuneta.Layer

        Next
        'UpdateDataDGV.UPDateDGView(DGView, CAcadHelp.ThisDrawing, CAcadHelp.Alignment)
        LayerColorChanger.SetcurrentMed(ComMediciones)
    End Sub

    'uso en button8
    'metodo para agregar una entidad analisada pero que se encuentra en un archivo anterior al actual 
    Public Shared Sub CopyEntityFromC(DGView As DataGridView, AcadHelp As ACAdHelpers, handleProcessor As HandleDataProcessor)
        ' Ensure at least one row is selected
        If DGView.SelectedCells.Count = 0 Then
            MessageBox.Show("Please select a cell first.")
            Return
        End If

        Try
            ' Proceed only if a row is selected
            Dim handler As New AutoCADHandler()
            Dim ProsesHandle As New List(Of String)()
            ' Safely retrieve TabName and FilePath from the selected row
            For Each Cell As DataGridViewCell In DGView.SelectedCells

                Dim Row As DataGridViewRow = DGView.Rows(Cell.RowIndex)

                Dim Handle As String = Row.Cells("Handle").Value.ToString()
                Dim ResultHandle As String = String.Empty
                If Not ProsesHandle.Contains(Handle) Then
                    Dim dwgPath As String = Row.Cells("FilePath").Value.ToString() 'DGView.SelectedRows(0).Cells("FilePath").Value.ToString()
                    If Not String.IsNullOrWhiteSpace(dwgPath) Or String.IsNullOrWhiteSpace(dwgPath) Then
                        ' Call to copy entity from closed DWG
                        handler.CopyEntityFromClosedDwg(dwgPath, Handle, AcadHelp.ThisDrawing, ResultHandle)

                        ProsesHandle.Add(Handle)
                        Dim Cuneta As CunetasHandleDataItem = Nothing
                        If handleProcessor.DicHandlesNoExistentes.ContainsKey(Handle) Then
                            Cuneta = handleProcessor.DicHandlesNoExistentes(Handle)
                        Else
                            Cuneta.Handle = Handle
                        End If


                        If Cuneta.Handle <> ResultHandle Then
                            Cuneta.Handle = ResultHandle
                            Cuneta.SetPropertiesFromDWG(Cuneta.Handle, Cuneta.AlignmentHDI)
                            Cuneta.Layer = Row.Cells("Layer").Value.ToString()
                            If Not String.IsNullOrEmpty(Cuneta.Handle) AndAlso Not String.IsNullOrEmpty(Row.Cells("Layer").Value.ToString()) Then
                                CLayerHelpers.ChangeLayersAcEnt(Cuneta.Handle, Cuneta.Layer)
                            End If
                            'CLHandle.GetEntityByStrHandle(Cuneta.Handle).Layer = Cuneta.Layer
                        End If
                        handleProcessor.DicHandlesNoExistentes.Remove(Handle)
                        handleProcessor.DicHandlesExistentes.Add(Cuneta.Handle, Cuneta)
                    End If
                End If
            Next
        Catch ex As Exception
            ' Log or display the exception message
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    'tabControlDGView
    Public Shared Sub AddAnEntty(TabCtrol As TabControl, CAcadHelp As ACAdHelpers,
                                 Optional Layer As String = vbNullString,
                                 Optional CleanDGV As Boolean = False,
                                 Optional handle As String = "")
        If TabCtrol.SelectedIndex <= 2 Then

            Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabCtrol)

            If CleanDGV Then DGView.Rows.Clear()

            Dim AlignmentHandle As String = CAcadHelp.Alignment?.Handle.ToString()

            If String.IsNullOrEmpty(Layer) Then

                If String.IsNullOrEmpty(handle) Then handle = CSelectionHelper.SelectEntityHandle().ToString()


                If handle IsNot Nothing AndAlso handle <> "0" Then AddHandle(handle, DGView, AlignmentHandle)

            Else
                For Each Hd As String In CLHandle.GetListofhandlebyLayer(Layer)

                    AddHandle(Hd, DGView, AlignmentHandle)

                Next
            End If
        End If
    End Sub
    Private Shared Sub AddHandle(Handle As String, DGView As DataGridView, AlignmentHnadle As String)
        Dim row As DataGridViewRow = Nothing
        If DateViewSet.ExisteHandle(Handle, DGView, row) = False Then

            If CLHandle.CheckIfExistHd(Handle) Then
                Dim Cuneta As New CunetasHandleDataItem

                Cuneta.SetPropertiesFromDWG(Handle, AlignmentHnadle)
                Cuneta.AddToDGView(DGView)
            Else
                ' Si no existe en el diccionario, buscar en la tabla de Excel
                Dim ExRNGSe As New ExcelRangeSelector()
                Dim RNG As Excel.Range = Nothing
                Dim Table As Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")
                RNG = ExRNGSe.LookInRange(Table:=Table, handle:=Handle, "Handle")
                If RNG IsNot Nothing Then
                    Dim Cuneta As New CunetasHandleDataItem(listObject:=Table)

                    Cuneta.SetPropertiesFromTableRng(RNG)
                    Cuneta.AddToDGView(DGView)
                End If

            End If
        Else
            row.Selected = True
            DGView.FirstDisplayedScrollingRowIndex = row.Index
        End If
    End Sub

    Public Shared Sub DividePL(Alignment As Alignment, Tbox As TextBox)
        Dim DPL As New CMDPolylineC
        If Tbox.Text = "" Then
            DPL.Dividitestpl()
        Else
            Dim Station As Double
            Double.TryParse(Tbox.Text.Replace("+", "").Replace(",", "."), Station) 'CDbl(charsToReplace.Aggregate(t, Function(current, ch) current.Replace(ch, "")))
            DPL.Dividitestpl(Alignment:=Alignment, Station:=Station)
        End If
    End Sub
    Public Shared Sub CFeatureLine(Optional DGView As DataGridView = Nothing, ByRef Optional handleProcessor As HandleDataProcessor = Nothing) ', Optional CBox As ComboBox = Nothing) 'CBox As ComboBox)
        If DGView IsNot Nothing Then
            If DGView.SelectedCells.Count = 0 Then
                MessageBox.Show("Please select a cell first.")
                Return
            End If
            Try
                ' Proceed only if a row is selected
                Dim handler As New AutoCADHandler()
                Dim ProsesHandle As New List(Of String)()
                ' Safely retrieve TabName and FilePath from the selected row
                For Each Cell As DataGridViewCell In DGView.SelectedCells
                    Dim Row As DataGridViewRow = DGView.Rows(Cell.RowIndex)

                    Dim Handle As String = Row.Cells("Handle").Value.ToString()
                    ' Call to copy entity from closed DWG
                    If handleProcessor.DicHandlesExistentes.ContainsKey(Handle) Then
                        Dim Cuneta As CunetasHandleDataItem = handleProcessor.DicHandlesExistentes(Handle)
                        Cuneta.ConvertToPolyline()
                        Cuneta.AddToDGView(DGView, Row.Index)
                        handleProcessor.DicHandlesExistentes.Remove(Handle)
                        handleProcessor.DicHandlesExistentes.Add(Cuneta.Handle, Cuneta)
                    End If

                Next
            Catch ex As Exception
                ' Log or display the exception message
                MessageBox.Show("An error occurred: " & ex.Message)
            End Try
        Else
            Dim Handlestr As String = CLHandle.GetEntityByHandle(CSelectionHelper.SelectEntityHandle()).Handle.ToString()
            Dim Cuneta As New CunetasHandleDataItem With {
                .Handle = Handlestr
            }

            If Not String.IsNullOrEmpty(Handlestr) Then
                Cuneta.ConvertToPolyline(False)
            End If
            'If Not String.IsNullOrEmpty(CBox.SelectedValue.ToString()) Then
            '	CLayerHelpers.ChangeLayersAcEnt(Cuneta.TabName, CBox.SelectedValue.ToString())
            'End If
        End If

    End Sub
    '1-btGetDataLCunetaFExcel: 
    Public Shared Sub GetData(Bt1 As Button, Bt2 As Button)
        Dim ExRNGSe As New ExcelRangeSelector()
        Try
            Dim handler As New AutoCADHandler()
            Dim Headers As New List(Of String) From {"Handle", "FilePath", "AlignmentHDI"}
            Dim Data As List(Of Object) = ExRNGSe.SelectRowOnTbl("CunetasGeneral", Headers) '"D:\Desktop\Typsa - Las Placetas\MED-HLP-ACC4-DRE-012-R4.dwg" ' Reemplaza con la ruta real
            Dim dwgPath As String = Data(0) 'DIRECCION DEL ARCHIVO 
            Dim handle As String = String.Empty ' Reemplaza con el handle real

            If Bt1.Focus() Then
                handle = Data(1) 'handle
            ElseIf Bt2.Focus() Then
                handle = Data(2) 'handle alignment
            End If

            handler.CopyEntityFromClosedDwg(dwgPath, handle)

        Catch ex As Exception
            Diagnostics.Debug.WriteLine("Error en CopyFromClosedDwgCommand: " & ex.Message)
        End Try
    End Sub

End Class
Public Class CommondUtl
    Public EventsHandler As New List(Of EventHandlerClass)
    Public Sub New(TabPanol As TabControl, TabCtrlViews As TabControl, CAcadHelp As ACAdHelpers, ComboBox1 As ComboBox, ComMediciones As ComboBox, handleProcessor As HandleDataProcessor)

        Dim Panel As New TabsMenu(TabPanol) With {.TabName = "Utilidades Generales"}

        Dim btGetDataLAlignmentFExcel As Button = Panel.CreateButtonAndGet("GetAlingmentHDFromExcel")

        Dim btGetDataLCunetaFExcel As Button = Panel.CreateButtonAndGet("Get Cuneta From Excel")

        Dim btShow As Button = Panel.CreateButtonAndGet("Motrar")

        Dim btHide As Button = Panel.CreateButtonAndGet("Ocultar")

        Dim DeleteENT As Button = Panel.CreateButtonAndGet("Delete")

        Dim CunetasExistentesDGView As DataGridView = GetControlsDGView.GetDGViewByTabName(TabCtrlViews, "Cunetas Existentes")

        Panel.ButtonActions = New Dictionary(Of String, Action) From {{"Cambiar Layer", Sub() CtrolsButtonsUT.CambiarLayer(ComboBox1, ComMediciones, TabCtrlViews, CAcadHelp)},
                                                                {"Add To Active DGView", Sub() CtrolsButtonsUT.AddAnEntty(TabCtrlViews, CAcadHelp)},
                                                                {"Add Ficture Line", Sub() CtrolsButtonsUT.AddAnEntty(TabCtrlViews, CAcadHelp, CSelectionHelper.GetLayerByEnt())},
                                                                {"Ad Lines", Sub() CtrolsButtonsUT.AddAnEntty(TabCtrlViews, CAcadHelp, CSelectionHelper.GetLayerByEnt(), True)},
                                                                {"Convet to Polyline", Sub() CtrolsButtonsUT.CFeatureLine()},
                                                                {"Copy Entity From File Externo", Sub() CtrolsButtonsUT.CopyEntityFromC(GetControlsDGView.GetDGViewByTabName(TabCtrlViews, "Cunetas Referencias Externas"), CAcadHelp, handleProcessor)},
                                                                {"Convert Selected Entities", Sub() CtrolsButtonsUT.CFeatureLine(CunetasExistentesDGView, handleProcessor)},
                                                                {"Identificar en DGView", Sub() DateViewSet.IdDataFEnt(GetControlsDGView.GetDGView(TabCtrlViews))},
                                                                {"Eliminar Entida", Sub() GetControlsDGView.Delete(TabCtrlViews, CAcadHelp)}
        }

        EventsHandler.Add(New EventHandlerClass({btGetDataLAlignmentFExcel, btGetDataLCunetaFExcel}, "click", Sub() CtrolsButtonsUT.GetData(btGetDataLCunetaFExcel, btGetDataLAlignmentFExcel)))

        'metodo para ocultar/mostar una entidad seleccionada en el dgv 
        EventsHandler.Add(New EventHandlerClass({btShow, btHide}, "click", Sub() CtrolsButtonsUT.VisibilityToggle(GetControlsDGView.GetDGView(TabCtrlViews),
                                    IIf(btShow.Focused, True, False))))

    End Sub
End Class
