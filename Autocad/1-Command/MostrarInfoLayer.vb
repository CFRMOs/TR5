Imports System.Windows.Forms
Imports Autodesk.AutoCAD.Runtime

Public Class MostrarInfoLayer
    <CommandMethod("CMDMostrarInfoLayer")>
    Public Sub ShowLayerInfo()
        ' Inicializar el formulario y mostrarlo
        Dim form As New PolylineMG With {
            .Text = "Información de entidades",
            .StartPosition = FormStartPosition.CenterScreen
        }
        form.SetDatosIni()

        ' Mostrar el formulario
        form.Show() ' Show()
    End Sub
End Class
