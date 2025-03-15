Imports Microsoft.Office.Interop
Imports System.Windows.Forms

Public Class TranferSlapingInf
    'manejo de las actividades en excel
    Public Shared Sub ExceTR(DGView1 As DataGridView, columnData As List(Of (String, String, String)), columnFormats As List(Of String))
        Dim xl As New ExcelAppTR
        Dim ShName As String = "Cunetas - General"
        Dim tableName As String = "CunetasGeneral"
        Dim ExcelWorksheet As Excel.Worksheet = xl.GetSH(ShName)

        If ExcelWorksheet IsNot Nothing Then
            ExcelWorksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible

            Dim headers As New List(Of String)
            For Each pair In columnData
                headers.Add(pair.Item1)
            Next

            xl.CreateTable(ExcelWorksheet, headers, tableName)
            For Each row As DataGridViewRow In DGView1.Rows
                If Not row.IsNewRow Then
                    Dim data(headers.Count - 1) As Object
                    Dim formats(headers.Count - 1) As String
                    For i As Integer = 0 To headers.Count - 1
                        data(i) = row.Cells(headers(i)).Value
                        formats(i) = columnFormats(i)
                    Next
                    xl.SetTransferData(data, headers.ToArray(), formats)
                End If
            Next
        Else
            MessageBox.Show("No se pudo obtener la hoja de Excel.")
        End If
    End Sub
End Class
