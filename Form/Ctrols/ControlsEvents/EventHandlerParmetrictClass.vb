Imports System.Windows.Forms

''' <summary>
''' Class to dynamically attach event handlers to controls.
''' </summary>
Public Class EventHandlerParametricClass
    Private WithEvents Control As Control
    Private eventHandlerKeyEvent As Action(Of Object, KeyEventArgs)
    Private eventHandlerEvent As Action(Of Object, EventArgs)
    Private eventHandlerMouseEvent As Action(Of Object, MouseEventArgs)

    ''' <summary>
    ''' Constructor for KeyDown event handler (Single Control).
    ''' </summary>
    ''' <param name="ctrl">The control to attach the event handler to.</param>
    ''' <param name="evento">The event name, should be "keydown".</param>
    ''' <param name="rutina">The action to execute when the event occurs.</param>
    ''' <exception cref="ArgumentException">Thrown if the event is not supported for KeyEventArgs.</exception>
    Public Sub New(ctrl As Control, evento As String, rutina As Action(Of Object, KeyEventArgs))
        If evento.ToLower() <> "keydown" Then
            Throw New ArgumentException("Evento no soportado para KeyEventArgs")
        End If
        eventHandlerKeyEvent = rutina
        AddHandlerForControl(ctrl, evento)
    End Sub

    ''' <summary>
    ''' Constructor for KeyDown event handler (Multiple Controls).
    ''' </summary>
    ''' <param name="ctrls">The controls to attach the event handler to.</param>
    ''' <param name="evento">The event name, should be "keydown".</param>
    ''' <param name="rutina">The action to execute when the event occurs.</param>
    ''' <exception cref="ArgumentException">Thrown if the event is not supported for KeyEventArgs.</exception>
    Public Sub New(ctrls As Control(), evento As String, rutina As Action(Of Object, KeyEventArgs))
        If evento.ToLower() <> "keydown" Then
            Throw New ArgumentException("Evento no soportado para KeyEventArgs")
        End If
        eventHandlerKeyEvent = rutina
        For Each ctrl As Control In ctrls
            AddHandlerForControl(ctrl, evento)
        Next
    End Sub

    ''' <summary>
    ''' Constructor for Click event handler (Single Control).
    ''' </summary>
    ''' <param name="ctrl">The control to attach the event handler to.</param>
    ''' <param name="evento">The event name, should be "click".</param>
    ''' <param name="rutina">The action to execute when the event occurs.</param>
    ''' <exception cref="ArgumentException">Thrown if the event is not supported for EventArgs.</exception>
    Public Sub New(ctrl As Control, evento As String, rutina As Action(Of Object, EventArgs))
        If evento.ToLower() <> "click" Then
            Throw New ArgumentException("Evento no soportado para EventArgs")
        End If
        eventHandlerEvent = rutina
        AddHandlerForControl(ctrl, evento)
    End Sub

    ''' <summary>
    ''' Constructor for Click event handler (Multiple Controls).
    ''' </summary>
    ''' <param name="ctrls">The controls to attach the event handler to.</param>
    ''' <param name="evento">The event name, should be "click".</param>
    ''' <param name="rutina">The action to execute when the event occurs.</param>
    ''' <exception cref="ArgumentException">Thrown if the event is not supported for EventArgs.</exception>
    Public Sub New(ctrls As Control(), evento As String, rutina As Action(Of Object, EventArgs))
        If evento.ToLower() <> "click" Then
            Throw New ArgumentException("Evento no soportado para EventArgs")
        End If
        eventHandlerEvent = rutina
        For Each ctrl As Control In ctrls
            AddHandlerForControl(ctrl, evento)
        Next
    End Sub

    ''' <summary>
    ''' Constructor for MouseWheel event handler (Single Control).
    ''' </summary>
    ''' <param name="ctrl">The control to attach the event handler to.</param>
    ''' <param name="evento">The event name, should be "mousewheel".</param>
    ''' <param name="rutina">The action to execute when the event occurs.</param>
    ''' <exception cref="ArgumentException">Thrown if the event is not supported for MouseEventArgs.</exception>
    Public Sub New(ctrl As Control, evento As String, rutina As Action(Of Object, MouseEventArgs))
        If evento.ToLower() <> "mousewheel" Then
            Throw New ArgumentException("Evento no soportado para MouseEventArgs")
        End If
        eventHandlerMouseEvent = rutina
        AddHandlerForControl(ctrl, evento)
    End Sub

    ''' <summary>
    ''' Constructor for MouseWheel event handler (Multiple Controls).
    ''' </summary>
    ''' <param name="ctrls">The controls to attach the event handler to.</param>
    ''' <param name="evento">The event name, should be "mousewheel".</param>
    ''' <param name="rutina">The action to execute when the event occurs.</param>
    ''' <exception cref="ArgumentException">Thrown if the event is not supported for MouseEventArgs.</exception>
    Public Sub New(ctrls As Control(), evento As String, rutina As Action(Of Object, MouseEventArgs))
        If evento.ToLower() <> "mousewheel" Then
            Throw New ArgumentException("Evento no soportado para MouseEventArgs")
        End If
        eventHandlerMouseEvent = rutina
        For Each ctrl As Control In ctrls
            AddHandlerForControl(ctrl, evento)
        Next
    End Sub

    ''' <summary>
    ''' Adds the appropriate event handler to a control based on the specified event.
    ''' </summary>
    ''' <param name="ctrl">The control to which the event handler will be added.</param>
    ''' <param name="evento">The event name.</param>
    ''' <exception cref="ArgumentException">Thrown if the event is not supported.</exception>
    Private Sub AddHandlerForControl(ctrl As Control, evento As String)
        Select Case evento.ToLower()
            Case "click"
                AddHandler ctrl.Click, AddressOf EjecutarRutinaEvent
            Case "keydown"
                AddHandler ctrl.KeyDown, AddressOf EjecutarRutinaKeyEvent
            Case "mousewheel"
                AddHandler ctrl.MouseWheel, AddressOf EjecutarRutinaMouseEvent
            Case Else
                Throw New ArgumentException("Evento no soportado.")
        End Select
    End Sub

    ''' <summary>
    ''' Executes the routine for the Click event.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">Event data.</param>
    Private Sub EjecutarRutinaEvent(sender As Object, e As EventArgs)
        eventHandlerEvent?.Invoke(sender, e)
    End Sub

    ''' <summary>
    ''' Executes the routine for the KeyDown event.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">Key event data.</param>
    Private Sub EjecutarRutinaKeyEvent(sender As Object, e As KeyEventArgs)
        eventHandlerKeyEvent?.Invoke(sender, e)
    End Sub

    ''' <summary>
    ''' Executes the routine for the MouseWheel event.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">Mouse event data.</param>
    Private Sub EjecutarRutinaMouseEvent(sender As Object, e As MouseEventArgs)
        eventHandlerMouseEvent?.Invoke(sender, e)
    End Sub
End Class
