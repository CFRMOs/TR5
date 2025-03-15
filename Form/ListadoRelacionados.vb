Public Class ListadoRelacionados
    Public WithEvents ThisDrawing As Document
    Private Sub DataGridView1_Click(sender As Object, e As EventArgs) Handles DataGridView1.Click
        Dim hdString As String = DataGridView1.CurrentRow.Cells("TabName").Value.ToString()
        AcadZoomManager.SelectedZoom(hdString, ThisDrawing)
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

    End Sub

    Private Sub ListadoRelacionados_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class