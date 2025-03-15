Imports System.Linq
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.GraphicsSystem
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel


Public Class DataGridViewEventHandler
    Private ReadOnly dataGridView As DataGridView
    Private ReadOnly acadHelp As ACAdHelpers
    Private ReadOnly parentForm As Form
    Private ReadOnly handleProcessor As HandleDataProcessor ' Add the handleProcessor reference
    Private ReadOnly ExcelManager As New ExcelManager
    Private ReadOnly fileSearcher As FileSearcher
    Private ReadOnly contextMenuHandler As ContextMenuHandler
    ' Updated constructor
    Public Sub New(dataGridView As DataGridView, acadHelp As ACAdHelpers, parentForm As Form, Optional ByRef handleProcessor As HandleDataProcessor = Nothing)
        Me.dataGridView = dataGridView
        Me.acadHelp = acadHelp
        Me.parentForm = parentForm
        Dim carpetaBase As String = "C:\Users\typsa\Desktop\1.0-Mediciones\Notas de Campo CH Las Placetas - Accesos"
        Me.fileSearcher = New FileSearcher(carpetaBase)
        'Me.ExcelManager.GetExcelApp() ' Initialize the ExcelApp

        ' Asignar manejadores de eventos
        AddHandler dataGridView.CellContentClick, AddressOf DataGridView1_Click
        AddHandler dataGridView.KeyDown, AddressOf DataGridView1_KeyDown
        AddHandler dataGridView.Resize, AddressOf DataGridView1_Resize
        AddHandler parentForm.Resize, AddressOf PolylineMG_Resize
        AddHandler dataGridView.MouseDown, AddressOf DataGridView1_MouseDown
        ' en los casos donde handleProcessor no se necesita 
        'solo cuneta 
        ' tener pendiente la integracionde de areas de acera y losas representadas en polylineas 
        If handleProcessor IsNot Nothing Then
            Me.handleProcessor = handleProcessor ' Initialize the handleProcessor
            AddHandler dataGridView.CellValueChanged, AddressOf DataGridView1_CellValueChanged

            ' Asignar manejadores para la edición de celdas
            AddHandler dataGridView.CellBeginEdit, AddressOf DataGridView1_CellBeginEdit
            AddHandler dataGridView.CellEndEdit, AddressOf DataGridView1_CellEndEdit
        End If

    End Sub
    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs)
        ' Verifica si es CTRL + Clic Derecho
        If e.Button = MouseButtons.Right AndAlso Control.ModifierKeys = Keys.Control Then
            Dim hti As DataGridView.HitTestInfo = dataGridView.HitTest(e.X, e.Y)

            ' Si se hace clic sobre una celda válida
            If hti.RowIndex >= 0 And hti.ColumnIndex >= 0 Then
                ' Seleccionar la celda donde se hizo clic derecho
                dataGridView.ClearSelection()
                dataGridView.Rows(hti.RowIndex).Cells(hti.ColumnIndex).Selected = True

                Dim Handle As String = dataGridView.Rows(hti.RowIndex).Cells("Handle").Value.ToString()

                Dim ExRNGSe As New ExcelRangeSelector()

                Using excelOptimizer As New ExcelOptimizer()
                    excelOptimizer.TurnEverythingOff()

                    Dim ExcelFilters As New ExcelFilters

                    Dim Table As ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")

                    ExcelFilters.GuardarFiltros(Table)


                    Dim RNG As Range = ExRNGSe.LookInRange(Table, Handle, "Handle")

                    If RNG Is Nothing Then
                        handleProcessor.UpDateByhandleFromCDI(Handle)
                    Else
                        UpdateDataDGV.UpdateSelectedEntityInfo(acadHelp.ThisDrawing, CLHandle.GetEntityByStrHandle(Handle), dataGridView.Rows(hti.RowIndex), acadHelp.Alignment)
                    End If

                    ExcelFilters.AplicarFiltros(Table)
                    excelOptimizer.TurnEverythingOn()
                    excelOptimizer.Dispose()
                End Using
            ElseIf e.Button = MouseButtons.Right Then 'AndAlso Not Control.ModifierKeys = Keys.Control
                Dim hitTestInfo As DataGridView.HitTestInfo = dataGridView.HitTest(e.X, e.Y)
                If hitTestInfo.RowIndex >= 0 AndAlso hitTestInfo.ColumnIndex >= 0 Then
                    Dim celda As DataGridViewCell = dataGridView.Rows(hitTestInfo.RowIndex).Cells(hitTestInfo.ColumnIndex)
                    Dim CodigoExtractor As New CodigoExtractor
                    Dim valorCelda As String = If(CodigoExtractor.ExtraerCodigo(celda.Value, "NCA-\d{5}-[A-Z]\d+"), "").ToString()
                    Dim rutaArchivo As String = fileSearcher.BuscarArchivo(valorCelda)

                    If Not String.IsNullOrEmpty(rutaArchivo) Then
                        contextMenuHandler.MostrarMenu(rutaArchivo)
                    Else
                        MessageBox.Show("No se encontró el archivo para: " & valorCelda, "Archivo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End If
            End If
        End If
    End Sub
    ' Manejar el clic en el contenido de la celda del DataGridView
    Private Sub DataGridView1_Click(sender As Object, e As DataGridViewCellEventArgs)
        Dim hdString As String = dataGridView.CurrentRow.Cells("Handle").Value.ToString()
        If CLHandle.CheckIfExistHd(hdString) Then
            AcadZoomManager.SelectedZoom(hdString, acadHelp.ThisDrawing)
            parentForm.Activate()
        Else
            ' Llama a GotoStation con la estación ya convertida
            Dim StartSation As Double = dataGridView.CurrentRow.Cells("StartStation").Value
            CStationOffsetLabel.GotoStation(StartSation, acadHelp.Alignment, acadHelp.ThisDrawing)
        End If

    End Sub

    ' Define el controlador del evento KeyDown
    ''' <summary>
    ''' Maneja los eventos KeyDown en el DataGridView para detectar desplazamientos y copia de datos.
    ''' </summary>
    ''' <param name="sender">El origen del evento.</param>
    ''' <param name="e">Los datos del evento de tecla.</param>
    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs)

        If dataGridView.Rows.Count = 0 Then Exit Sub ' Evita errores si no hay filas

        Dim currentRow As Integer = dataGridView.CurrentCell.RowIndex

        If e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down Then
            HandleDataGridViewKeyDown(e)
        ElseIf e.Control AndAlso e.KeyCode = Keys.C Then
            ' Detecta si el usuario ha presionado Ctrl + C para copiar
            If dataGridView.SelectedCells.Count > 0 Then
                ' Ordenar las celdas seleccionadas por fila y columna
                Dim selectedCells = dataGridView.SelectedCells.Cast(Of DataGridViewCell)().OrderBy(Function(c) c.RowIndex).ThenBy(Function(c) c.ColumnIndex)

                ' Obtener el rango de filas y columnas seleccionadas
                Dim minRow As Integer = selectedCells.Min(Function(c) c.RowIndex)
                Dim maxRow As Integer = selectedCells.Max(Function(c) c.RowIndex)
                Dim minCol As Integer = selectedCells.Min(Function(c) c.ColumnIndex)
                Dim maxCol As Integer = selectedCells.Max(Function(c) c.ColumnIndex)

                ' Crear un array bidimensional para almacenar los valores
                Dim clipboardData(maxRow - minRow, maxCol - minCol) As String

                ' Rellenar el array con valores seleccionados
                For Each cell As DataGridViewCell In selectedCells
                    Dim rowIndex As Integer = cell.RowIndex - minRow
                    Dim colIndex As Integer = cell.ColumnIndex - minCol

                    ' Convertimos el valor a Double si es posible, de lo contrario, mantenemos el valor original
                    Dim formattedValue As String = cell.Value.ToString()
                    Dim numericValue As Double
                    If Double.TryParse(formattedValue.Replace("+", ""), numericValue) Then
                        clipboardData(rowIndex, colIndex) = numericValue.ToString()
                    Else
                        clipboardData(rowIndex, colIndex) = formattedValue
                    End If
                Next

                ' Crear un texto para el portapapeles con el array bidimensional
                Dim clipboardText As New System.Text.StringBuilder()
                For i As Integer = 0 To maxRow - minRow
                    For j As Integer = 0 To maxCol - minCol
                        clipboardText.Append(If(clipboardData(i, j), "") & vbTab)
                    Next
                    clipboardText.Length -= 1 ' Eliminar el último tabulador de la línea
                    clipboardText.AppendLine()
                Next

                ' Establecemos el texto en el portapapeles
                Clipboard.SetText(clipboardText.ToString().Trim())
                e.SuppressKeyPress = True
            End If
        ElseIf e.Control AndAlso e.KeyCode = Keys.V Then ' Verifica si se presiona CTRL + V

            If Clipboard.ContainsText() Then
                Dim clipboardText As String = Clipboard.GetText().Trim()

                ' Verifica si hay una celda seleccionada
                If dataGridView.SelectedCells.Count > 0 Then
                    For Each cell As DataGridViewCell In dataGridView.SelectedCells
                        ' Pega solo en celdas editables
                        If Not cell.ReadOnly Then
                            cell.Value = clipboardText


                            Dim comentario As Object = cell.Value
                            Dim handle As Object = dataGridView.Rows(cell.RowIndex).Cells("Handle").Value

                            ' Ensure values are not Nothing before proceeding
                            If comentario IsNot Nothing AndAlso handle IsNot Nothing Then
                                'Await Task.Run(Sub()
                                ActualizarTablaEnExcel(handle.ToString(), comentario.ToString(), "Comentarios")
                                'End Sub)
                            End If

                        End If
                    Next
                End If


                ' Ajustar automáticamente la altura de las filas en función del contenido
                With dataGridView
                    .DefaultCellStyle.WrapMode = DataGridViewTriState.True
                    .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                    .AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells)
                End With

            End If
        ElseIf (e.Control AndAlso e.KeyCode = Keys.Up) Then

            ' Si está en la primera fila, mover a la última fila
            If currentRow = 0 Then
                dataGridView.CurrentCell = dataGridView.Rows(dataGridView.Rows.Count - 1).Cells(dataGridView.CurrentCell.ColumnIndex)
            Else
                ' Mover hacia arriba
                dataGridView.CurrentCell = dataGridView.Rows(currentRow - 1).Cells(dataGridView.CurrentCell.ColumnIndex)
            End If
            e.Handled = True
        ElseIf (e.Control AndAlso e.KeyCode = Keys.Down) Then
            ' Si está en la última fila, mover a la primera fila
            If currentRow = dataGridView.Rows.Count - 1 Then
                dataGridView.CurrentCell = dataGridView.Rows(0).Cells(dataGridView.CurrentCell.ColumnIndex)
            Else
                ' Mover hacia abajo
                dataGridView.CurrentCell = dataGridView.Rows(currentRow + 1).Cells(dataGridView.CurrentCell.ColumnIndex)
            End If
            e.Handled = True

        End If
    End Sub

    ' Manejar el redimensionamiento del DataGridView
    Private Sub DataGridView1_Resize(sender As Object, e As EventArgs)
        dataGridView.ScrollBars = System.Windows.Forms.ScrollBars.Both
    End Sub

    ' Ajustar el tamaño del DataGridView cuando se redimensiona el formulario
    Private Sub PolylineMG_Resize(sender As Object, e As EventArgs)
        ResizeDataGridView()
    End Sub

    ' Redimensionar el DataGridView para ajustarlo al formulario
    Private Sub ResizeDataGridView()
        Dim tabControl As TabControl = GetTabControlFromDataGridView(dataGridView)

        If dataGridView IsNot Nothing Then
            ' Ajustar el tamaño del DataGridView para que ocupe el 90% del ancho y el 80% de la altura del formulario
            dataGridView.Width = tabControl.ClientSize.Width

            ' Centrar el DataGridView en el formulario
            dataGridView.Left = (tabControl.ClientSize.Width - dataGridView.Width) / 2
        End If
    End Sub

    ' Manejar la navegación del DataGridView con las teclas de flecha
    Private Sub HandleDataGridViewKeyDown(e As KeyEventArgs)
        e.Handled = True

        Dim currentRowIndex As Integer = dataGridView.CurrentCell.RowIndex

        Dim currentColumnIndex As Integer = dataGridView.CurrentCell.ColumnIndex

        Dim newRowIndex As Integer = If(e.KeyCode = Keys.Up Or e.KeyCode = Keys.Enter, currentRowIndex - 1, currentRowIndex + 1)

        DateViewSet.HandleDataGridView(dataGridView, acadHelp.ThisDrawing, currentColumnIndex, newRowIndex)
    End Sub
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        'Dim total As Long = Operaciones.CalcularSumaColumna(dataGridView, "Longitud")
        ' Actualizar la Label con el resultado de la sumatoria
        'parentForm.Controls("Label5").Text = "Longitud Total: " & total.ToString("0.000") & " m."
        With dataGridView
            .DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            .AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells)
        End With
    End Sub
    Private Function GetTabControlFromDataGridView(dataGridView As DataGridView) As TabControl
        ' Verifica si el DataGridView tiene un padre, que debería ser la TabPage
        Dim parentTabPage As TabPage = TryCast(dataGridView.Parent, TabPage)

        ' Si el padre es una TabPage, obtenemos su contenedor, que será el TabControl
        If parentTabPage IsNot Nothing Then
            Return TryCast(parentTabPage.Parent, TabControl)
        End If

        ' Si no pertenece a una TabPage o TabControl, retorna Nothing
        Return Nothing
    End Function
    ' Detectar cuándo empieza la edición en la columna de comentarios
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs)
        ' Verificar si la columna es la de "Comentarios"
        If dataGridView.Columns(e.ColumnIndex).Name = "Comentarios" Then
            ' Aquí puedes ejecutar cualquier lógica cuando comienza la edición de una celda de comentarios
            ' Por ejemplo, puedes mostrar un mensaje o preparar la interfaz
            'MessageBox.Show("Editando Connum en la fila " & e.RowIndex.ToString())

        ElseIf dataGridView.Columns(e.ColumnIndex).Name = "IDNum" Then

        End If
    End Sub
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
        Try
            ' Check if the edited column is "Comentarios"
            If dataGridView.Columns(e.ColumnIndex).Name = "Comentarios" Then
                Dim comentario As Object = dataGridView.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Dim handle As Object = dataGridView.Rows(e.RowIndex).Cells("Handle").Value
                ' Ensure values are not Nothing before proceeding
                If comentario IsNot Nothing AndAlso handle IsNot Nothing Then
                    'Await Task.Run(Sub()
                    ActualizarTablaEnExcel(handle.ToString(), comentario.ToString(), "Comentarios")
                    'End Sub)
                End If
                ' Ajustar automáticamente la altura de las filas en función del contenido
                With dataGridView
                    .DefaultCellStyle.WrapMode = DataGridViewTriState.True
                    .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                    .AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells)
                End With

            ElseIf dataGridView.Columns(e.ColumnIndex).Name = "IDNum" Then

                Dim IDNum As Object = dataGridView.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Dim handle As Object = dataGridView.Rows(e.RowIndex).Cells("Handle").Value
                'Dim cunetasHandleDataItem As New CunetasHandleDataItem
                ActualizarTablaEnExcel(handle.ToString(), IDNum.ToString(), "IDNum")

            ElseIf dataGridView.Columns(e.ColumnIndex).Name = "Posicion Formatted" Then

                Dim IDNum As Object = dataGridView.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Dim handle As Object = dataGridView.Rows(e.RowIndex).Cells("Handle").Value
                'Dim cunetasHandleDataItem As New CunetasHandleDataItem
                ActualizarTablaEnExcel(handle.ToString(), IDNum.ToString(), "Ubicacion")

            ElseIf dataGridView.Columns(e.ColumnIndex).Name = "PLANO" Then

                Dim PLANO As Object = dataGridView.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Dim handle As Object = dataGridView.Rows(e.RowIndex).Cells("Handle").Value
                'Dim cunetasHandleDataItem As New CunetasHandleDataItem
                ActualizarTablaEnExcel(handle.ToString(), PLANO.ToString(), "PLANO")

            End If
        Catch ex As Exception
            ' Log or handle exceptions here
        End Try


    End Sub

    ' Método para actualizar los comentarios en Excel
    Private Sub ActualizarTablaEnExcel(handle As String, comentario As String, Header As String)
        'sinplificar su implementacion agregando declraciones prepetidas y necesarias para su unico uso y manejo 

        'Dim Texto As Object = dataGridView.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        'Dim handle As Object = dataGridView.Rows(e.RowIndex).Cells("Handle").Value

        ' Asegúrate de usar bloqueos de recursos si hay posibilidad de problemas de concurrencia
        'SyncLock Me
        If handleProcessor.HandleData.Count <> 0 AndAlso handleProcessor.DicHandlesExistentes.ContainsKey(handle) Then
            ' Si el handle ya existe en el diccionario, actualiza el comentario
            Dim cunetasHandleDataItem As CunetasHandleDataItem = handleProcessor.DicHandlesExistentes(handle)
            If cunetasHandleDataItem.RNGComentarios IsNot Nothing Then
                Dim OP As New ExcelOptimizer(Me.ExcelManager.GetExcelApp)
                OP.TurnEverythingOff()
                cunetasHandleDataItem.RNGComentarios.Value = comentario
                cunetasHandleDataItem.Comentarios = comentario ' Asegúrate de que este campo existe
                OP.TurnEverythingOn()
                OP.Dispose()
            Else
                'si no se encontro en al tabla se debera añadir la informacion nueva 
                handleProcessor.UpDateExcel(dataGridView.Parent, handle)
            End If
        Else
            ' Si no existe en el diccionario, buscar en la tabla de Excel
            Dim ExRNGSe As New ExcelRangeSelector()
            Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing
            RNG = ExRNGSe.HandleFromExcelTable(TBName:="CunetasGeneral", handle:=handle, Header)
            If RNG IsNot Nothing Then
                RNG.Value = comentario
            Else
                'si no se encontro en al tabla se debera añadir la informacion nueva 
                'pendiente optimizacion del proceso de updatedata
                Dim TabCtrol As TabControl = dataGridView.Parent.Parent
                handleProcessor.UpDateExcel(TabCtrol, handle)
                RNG = ExRNGSe.HandleFromExcelTable(TBName:="CunetasGeneral", handle:=handle, Header)
                RNG.Value = comentario
            End If
        End If
        'End SyncLock
    End Sub
    Private Sub AddValueToTableCell(ByRef RNGRow As Excel.Range, headerName As String, value As Object, Optional format As String = "@")

        Dim ExRNGSe As New ExcelRangeSelector()

        Dim TB As Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")

        If TB Is Nothing Then Exit Sub

        ' Encuentra el encabezado correspondiente en la tabla
        Dim headerCell As Range = TB.HeaderRowRange.Cells.Cast(Of Range)().FirstOrDefault(Function(c) c.Value = headerName)

        If headerCell IsNot Nothing Then
            ' Calcula el índice de la columna del encabezado
            Dim columnIndex As Integer = headerCell.Column - TB.HeaderRowRange(1).Column + 1

            ' Asigna el valor a la celda de la nueva fila
            RNGRow.Cells(1, columnIndex).Value = value

            ' Aplica el formato si es necesario
            RNGRow.Cells(1, columnIndex).NumberFormat = format
        End If
    End Sub
End Class
