Imports System.Windows.Forms

Public Class EventHandlerClass
    Private eventHandler As Action

    ' Constructor para un solo control (como lo usas ahora)
    Public Sub New(ctrl As Control, evento As String, rutina As Action)
        AddHandlerForControl(ctrl, evento, rutina)
    End Sub

    ' Sobrecarga: Constructor para múltiples controles
    Public Sub New(ctrls As Control(), evento As String, rutina As Action)
        For Each ctrl As Control In ctrls
            AddHandlerForControl(ctrl, evento, rutina)
        Next
    End Sub

    ' Método privado para agregar manejadores de eventos
    Private Sub AddHandlerForControl(ctrl As Control, evento As String, rutina As Action)
        eventHandler = rutina
        Select Case evento.ToLower()
            Case "click"
                AddHandler ctrl.Click, AddressOf EjecutarRutina
            Case Else
                Throw New ArgumentException("Evento no soportado.")
        End Select
    End Sub

    ' Ejecutar la rutina cuando se dispare el evento
    Private Sub EjecutarRutina(sender As Object, e As EventArgs)
        eventHandler?.Invoke()
    End Sub
End Class
