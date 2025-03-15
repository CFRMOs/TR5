Imports System.Diagnostics
Imports System.Linq
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Public Class DataTransferManager
    Public WithEvents ExcelApp As ExcelAppTR
    Private ReadOnly TableReferences As New Dictionary(Of String, Tuple(Of ExcelAppTR, ListObject, DataGridView))

    ' Constructor
    Public Sub New(tabControlDGView As TabControl, columnData As List(Of (String, String, String)))
        ExcelApp = New ExcelAppTR()
        If tabControlDGView.TabPages.Count > 0 Then
            Dim DGView As DataGridView = GetControlsDGView.GetDGView(tabControlDGView)
            InitializeTablesAndDGV(columnData,
            Tuple.Create("Cunetas - General", "CunetasGeneral", DGView)) ',
            'Tuple.Create("Cunetas Diseños", "CunetasDiseñosTable", Diseños))
        End If

    End Sub

    '' Initialize tables and DataGridViews
    'Private Sub InitializeTablesAndDGV(DataTransferManager As DataTransferManager)
    '	DataTransferManager.InitializeTablesAndDGV(columnData,
    '		Tuple.Create("Cunetas - General", "CunetasGeneral", Me.CunetasExistentesDGView),
    '		Tuple.Create("Cunetas Diseños", "CunetasDiseñosTable", Diseños))
    'End Sub

    ' Method to initialize tables and DataGridViews
    Public Sub InitializeTablesAndDGV(columnData As List(Of (String, String, String)), ParamArray views() As Tuple(Of String, String, DataGridView))
        For Each view In views
            Dim sheetName = view.Item1
            Dim tableName = view.Item2
            Dim dgv = view.Item3

            If dgv IsNot Nothing Then
                Dim ExcelWorksheet As Worksheet = ExcelApp.GetSH(sheetName)
                If ExcelWorksheet IsNot Nothing Then
                    ExcelWorksheet.Visible = XlSheetVisibility.xlSheetVisible

                    Dim headers As New List(Of String)
                    For Each pair In columnData
                        headers.Add(pair.Item1)
                    Next
                    Dim xlTable = ExcelApp.GetOrCreateTable(ExcelWorksheet, headers, tableName)
                    TableReferences(tableName) = New Tuple(Of ExcelAppTR, ListObject, DataGridView)(ExcelApp, xlTable, dgv)
                Else
                    MessageBox.Show("No se pudo obtener la hoja de Excel.")
                End If
            End If
        Next
    End Sub

    ' Method to export data to Excel
    Public Sub ExportDataToExcel(columnData As List(Of (String, String, String)), dataGridView As DataGridView, ByRef columnFormats As List(Of String)) 'columnFormats As List(Of String), 

        Dim op As New ExcelOptimizer(ExcelApp.ExcelApp)
        op.TurnEverythingOff()
        Try

            ' Buscar la tabla asociada con el DataGridView proporcionado
            Dim tableReference = TableReferences.Values.FirstOrDefault(Function(tuple) tuple.Item3 Is dataGridView)

            If tableReference IsNot Nothing Then
                Dim xl As ExcelAppTR = tableReference.Item1
                Dim xlTable As ListObject = tableReference.Item2
                Dim dgv As DataGridView = tableReference.Item3
                xl.xlTable = xlTable
                ' Añadir mensajes de depuración para verificar la tabla y el DataGridView
                Debug.WriteLine($"Exportando datos a la tabla: {xlTable.Name}, DataGridView: {dgv.Name}")

                AddDataToExcel(xl, dgv, columnData, columnFormats)
            Else
                MessageBox.Show("No se encontró ninguna tabla asociada con el DataGridView proporcionado.")
            End If

        Catch ex As Exception
            Console.Write(ex.Message)
        Finally
            op.TurnEverythingOn()
            op.Dispose()
        End Try
    End Sub

    ' Add data from DataGridView to Excel
    Private Sub AddDataToExcel(ByVal xl As ExcelAppTR, ByVal dataGridView As DataGridView, columnData As List(Of (String, String, String)), columnFormats As List(Of String))
        Dim headers = columnData.Select(Function(pair) pair.Item1).ToList()
        Try
            For Each dgvRow As DataGridViewRow In dataGridView.Rows
                If Not dgvRow.IsNewRow Then
                    Dim rowData As New List(Of Object)
                    For colIndex As Integer = 0 To headers.Count - 1
                        Dim columnName As String = headers(colIndex)
                        If dataGridView.Columns.Contains(columnName) Then
                            Dim cellValueInner As Object = dgvRow.Cells(columnName).Value
                            If columnName.ToLower() = "handle" Then 'TypeOf cellValueInner Is TabName OrElse
                                cellValueInner = cellValueInner.ToString() ' Ensure handle is treated as string
                            End If
                            'If xlTable.range.contains(cellValueInner) Then
                            rowData.Add(cellValueInner)
                            'End If
                        Else
                            MessageBox.Show("Column named '" & columnName & "' cannot be found in DataGridView.")
                        End If
                    Next

                    ' Check if the cell value is not Nothing before converting to string
                    Dim cellValue As Object = dgvRow.Cells("Comentarios").Value
                    Dim comments As String = If(cellValue IsNot Nothing, cellValue.ToString(), String.Empty)
                    xl.SetTransferData(rowData.ToArray(), headers.ToArray(), columnFormats.ToArray(), xl.xlTable, comments)
                End If
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    ' Get comment from Excel
    Public Function GetCommentFromExcel(handle As String) As String
        Try
            For Each keyValuePair In TableReferences
                Dim tableName As String = keyValuePair.Key
                Dim tuple As Tuple(Of ExcelAppTR, ListObject, DataGridView) = keyValuePair.Value
                Dim xl As ExcelAppTR = tuple.Item1
                Dim xlTable As ListObject = tuple.Item2

                ' Get headers as a string array
                Dim headers As String() = CType(xlTable.HeaderRowRange.Value, Object(,)).Cast(Of Object)().Select(Function(o) o.ToString()).ToArray()

                ' Get data from the table
                Dim data As Object(,) = CType(xlTable.DataBodyRange.Value, Object(,))
                Dim comments As String = xl.GetTransferData(data, headers, handle, "", "")("Comentarios")
                If Not String.IsNullOrEmpty(comments) Then
                    Return comments
                End If
            Next
            Return String.Empty
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Stop
            Return String.Empty ' Ensure a return value is always provided
        End Try
    End Function

End Class
