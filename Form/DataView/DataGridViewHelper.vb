Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
Public Class DataGridViewHelper
    Private WithEvents DdataGridView As DataGridView

    Public Sub New(dataGridView As DataGridView)
        DdataGridView = dataGridView
        AddHandler DdataGridView.KeyDown, AddressOf DataGridView_KeyDown
    End Sub

    Private Sub DataGridView_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.C Then
            CopyToClipboard()
        ElseIf e.Control AndAlso e.KeyCode = Keys.V Then
            PasteFromClipboard()
        ElseIf e.KeyCode = Keys.Delete Then
            DeleteSelectedCells()
        End If
    End Sub
    ' Add columns to DGView
    Public Shared Sub Addcolumns(DGView As DataGridView, columnData As List(Of (String, String, String))) ', columnFormats As List(Of String))
        'Dim Columns As DataGridViewColumnCollection = DGView.Columns
        Dim columnaformatos As List(Of String) = HandleDataProcessor.Headers.Select(Function(t) t.Item3).ToList()
        'Columns.Clear()
        If DGView.Columns.Count = 0 Then
            For Each pair In columnData
                DGView.Columns.Add(pair.Item1, pair.Item2)
            Next
            ' Aplicamos los formatos correspondientes
            For i As Integer = 0 To DGView.Columns.Count - 1
                ' Asignamos el formato a cada columna basada en el índice correspondiente
                Dim formato As String = columnData(i).Item3
                If i < DGView.Columns.Count Then
                    DGView.Columns(i).DefaultCellStyle.Format = formato
                End If
            Next
        End If
    End Sub
    Private Sub CopyToClipboard()
        If DdataGridView.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            Try
                Dim dataObj As DataObject = DdataGridView.GetClipboardContent()
                If dataObj IsNot Nothing Then
                    Clipboard.SetDataObject(dataObj)
                    'MessageBox.Show("Contenido copiado al portapapeles")
                End If
            Catch ex As ExternalException
                MessageBox.Show("No se pudo copiar al portapapeles")
            End Try
        End If
    End Sub

    Private Sub PasteFromClipboard()
        Try
            Dim clipboardText As String = Clipboard.GetText()
            Dim lines As String() = clipboardText.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

            If lines.Length = 1 AndAlso lines(0).Split(ControlChars.Tab).Length = 1 Then
                ' Caso especial: copiar una celda a múltiples celdas seleccionadas
                Dim singleValue As String = lines(0)
                For Each cell As DataGridViewCell In DdataGridView.SelectedCells
                    cell.Value = singleValue
                Next
            Else
                ' Pegado normal desde el portapapeles
                Dim currentRowIndex As Integer = DdataGridView.CurrentCell.RowIndex
                Dim currentColIndex As Integer = DdataGridView.CurrentCell.ColumnIndex

                For i As Integer = 0 To lines.Length - 1
                    If lines(i) <> String.Empty Then
                        Dim values As String() = lines(i).Split(ControlChars.Tab)
                        For j As Integer = 0 To values.Length - 1
                            If (currentColIndex + j) < DdataGridView.ColumnCount AndAlso (currentRowIndex + i) < DdataGridView.RowCount Then
                                DdataGridView(currentColIndex + j, currentRowIndex + i).Value = values(j)
                            End If
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("No se pudo pegar el contenido del portapapeles")
        End Try
    End Sub
    ' Add entity info to DGView
    Public Shared Function StrAddEntityInfo(ThisDrawing As Document, Alignment As Alignment, ByVal acEntHandle As String, ByRef DataGridView As DataGridView) As Integer
        Return AddEntityInfo(ThisDrawing, Alignment, CLHandle.GetEntityByStrHandle(acEntHandle), DataGridView)
    End Function
    Public Shared Function AddEntityInfo(ThisDrawing As Document, Alignment As Alignment, ByVal acEnt As Entity, ByRef DataGridView As DataGridView) As Integer
        If Alignment IsNot Nothing Then
            Dim rowIndex As Integer = DataGridView.Rows.Add()
            Dim selectedRow As DataGridViewRow = DataGridView.Rows(rowIndex)
            Return UpdateDataDGV.UpdateSelectedEntityInfo(ThisDrawing, acEnt, selectedRow, Alignment)
        End If
        Return -1
    End Function

    Private Sub DeleteSelectedCells()
        For Each cell As DataGridViewCell In DdataGridView.SelectedCells
            cell.Value = String.Empty
        Next
    End Sub
End Class
