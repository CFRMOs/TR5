Imports System.Runtime.InteropServices
Imports System.Windows.Forms

'Public Class UserControl1

'    Private acadApp As AcadApplication = Nothing

'    Private Sub UserControl1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
'        ' Verificar si AutoCAD está abierto al cargar el control
'        Try
'            ' Intentar obtener la instancia activa de AutoCAD
'            acadApp = CType(Marshal.GetActiveObject("AutoCAD.Application"), AcadApplication)
'            MessageBox.Show("AutoCAD está abierto.")
'        Catch ex As Exception
'            ' AutoCAD no está abierto, así que lo abrimos
'            Try
'                acadApp = New AcadApplication With {
'                    .Visible = True
'                }
'                MessageBox.Show("AutoCAD se abrió correctamente.")
'            Catch ex2 As Exception
'                MessageBox.Show("No se pudo iniciar AutoCAD. Error: " & ex2.Message)
'            End Try
'        End Try
'    End Sub

'    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles CunetasExistentesDGView.CellContentClick
'        ' Aquí puedes manejar el evento CellContentClick del DataGridView
'        ' Por ejemplo:
'        ' MessageBox.Show("Se hizo clic en una celda.")
'    End Sub

'    ' Puedes agregar más eventos y métodos según tus necesidades

'End Class
