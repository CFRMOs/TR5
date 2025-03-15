Imports System.Windows.Forms

Public Class Operaciones
    ' Función Shared para calcular la sumatoria de una columna en un DataGridView
    Public Shared Function CalcularSumaColumna(DGView As DataGridView, CName As String) As Long
        Dim suma As Long = 0

        ' Iterar a través de las filas del DataGridView
        For Each row As DataGridViewRow In DGView.Rows
            ' Asegurarse de que la fila no sea una fila nueva (NewRow)
            If Not row.IsNewRow Then
                ' Sumar el valor de la columna especificada
                suma += Convert.ToInt64(row.Cells(CName).Value)
            End If
        Next
        Return suma
    End Function
End Class

