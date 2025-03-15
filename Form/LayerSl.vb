Imports System.Linq

Public Class LayerSl
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim nombresCapas As String() = {"CU-BO1", "CU-T1", "CU-T2", "BADEN", "FUERA DE PROYECTO(CU-T2)"}
        Dim capas As List(Of String) = CrearGrupoYCapas("Drenajes Longitudinales", nombresCapas)
        CargarCapas(capas)
    End Sub
    Private Sub CargarCapas(capas As List(Of String))
        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        ' Vincular la lista de capas al control DataGridView
        DataGridView1.DataSource = capas.Select(Function(x) New With {Key .LayerName = x}).ToList()
    End Sub
End Class