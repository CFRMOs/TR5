Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq
Imports System.Runtime.InteropServices.WindowsRuntime
Imports System.Windows.Controls.Primitives
Imports System.Windows.Forms
Imports TR5.Civil3D

''' <summary>
''' Clase para gestionar y organizar TabControls en un formulario de Windows Forms.
''' </summary>
''' <remarks>
''' Proporciona métodos para agregar y posicionar dinámicamente TabControls en un formulario.
''' </remarks>
Public Class ControlManager
    ''' <summary>
    ''' TabControl principal para DataGridViews.
    ''' </summary>
    Public _tabControlDGView As TabControl

    ''' <summary>
    ''' TabControl principal para menús.
    ''' </summary>
    Public _tabControlMenu As TabControl

    ''' <summary>
    ''' Gestor de GroupBoxes asociado con la clase.
    ''' </summary>
    Public Property GroupBoxManager As GroupBoxManager

    ''' <summary>
    ''' Helper para operaciones específicas de AutoCAD.
    ''' </summary>
    Public Acahelp As ACAdHelpers

    ''' <summary>
    ''' Agrega un nuevo TabControl al formulario especificado.
    ''' </summary>
    ''' <param name="formulario">El formulario donde se agregará el TabControl.</param>
    ''' <param name="tabControlPrevio">Opcional. TabControl previo para posicionamiento relativo.</param>
    ''' <param name="Name">Opcional. Nombre del nuevo TabControl.</param>
    ''' <returns>El TabControl recién creado.</returns>
    Public Function AgregarTabControl(formulario As Form, Optional tabControlPrevio As TabControl = Nothing, Optional Name As String = "") As TabControl
        ' Crear el nuevo TabControl
        Dim miTabControl As New TabControl()
        Dim padding As Double = 5
        ' Definir las dimensiones del TabControl
        MoveTabCtrolMittleLeftDown(formulario, miTabControl, padding, tabControlPrevio)

        ' Agregar el TabControl al formulario
        formulario.Controls.Add(miTabControl)

        ' Retornar el TabControl creado
        Return miTabControl
    End Function

    ''' <summary>
    ''' Constructor de la clase ControlManager.
    ''' </summary>
    ''' <param name="formulario">Formulario al cual se agregarán los TabControls iniciales.</param>
    Public Sub New(formulario As Form)
        _tabControlDGView = AgregarTabControl(formulario)
        _tabControlMenu = AgregarTabControl(formulario, _tabControlDGView)
        Dim columnaformatos As List(Of String) = HandleDataProcessor.Headers.Select(Function(t) t.Item3).ToList()
    End Sub

    ''' <summary>
    ''' Posiciona un TabControl en la parte inferior izquierda del formulario.
    ''' </summary>
    ''' <param name="formulario">El formulario donde se posicionará el TabControl.</param>
    ''' <param name="miTabControl">El TabControl que será posicionado.</param>
    ''' <param name="Padding">El margen entre el TabControl y el borde del formulario.</param>
    ''' <param name="tabControlPrevio">Opcional. TabControl previo para posicionamiento relativo.</param>
    Public Sub MoveTabCtrolMittleLeftDown(formulario As Form, miTabControl As TabControl, Padding As Double, Optional tabControlPrevio As TabControl = Nothing)
        If tabControlPrevio Is Nothing Then
            miTabControl.Width = formulario.ClientSize.Width - Padding
            miTabControl.Height = formulario.ClientSize.Height \ 2
            miTabControl.Location = New Point(Padding, 0)
        Else
            miTabControl.Width = formulario.ClientSize.Width * 0.6
            miTabControl.Height = (formulario.ClientSize.Height - Padding * 6) \ 2
            Dim margen As Integer = Padding
            miTabControl.Location = New Point(Padding * 4, tabControlPrevio.Bottom + margen)
        End If

        AddHandler formulario.Resize, Sub(sender As Object, e As EventArgs)
                                          If tabControlPrevio Is Nothing Then
                                              miTabControl.Width = formulario.ClientSize.Width - Padding
                                              miTabControl.Height = formulario.ClientSize.Height \ 2
                                          Else
                                              miTabControl.Width = formulario.ClientSize.Width * 0.5 - Padding
                                              miTabControl.Height = (formulario.ClientSize.Height - Padding * 6) \ 2
                                              miTabControl.Location = New Point(Padding, tabControlPrevio.Bottom + 2 * Padding)
                                          End If
                                      End Sub
    End Sub

    ''' <summary>
    ''' Posiciona un TabControl en la parte inferior derecha del formulario.
    ''' </summary>
    ''' <param name="formulario">El formulario donde se posicionará el TabControl.</param>
    ''' <param name="miTabControl">El TabControl que será posicionado.</param>
    ''' <param name="Padding">El margen entre el TabControl y el borde del formulario.</param>
    ''' <param name="tabControlPrevio">Opcional. TabControl previo para posicionamiento relativo.</param>
    ''' <remarks>
    ''' Este método calcula las dimensiones y posición del TabControl en función de los controles existentes 
    ''' y ajusta su posición cada vez que se redimensiona el formulario.
    ''' </remarks>
    Public Sub MoveTabCtrolMittleRighttDown(formulario As Form, miTabControl As TabControl, Padding As Double, Optional tabControlPrevio As TabControl = Nothing)
        ' Definir las dimensiones del TabControl
        If tabControlPrevio Is Nothing Then
            ' Si no hay TabControl previo, usar todo el ancho y la mitad de la altura del formulario
            miTabControl.Width = formulario.ClientSize.Width - Padding
            miTabControl.Height = formulario.ClientSize.Height \ 2

            ' Posicionar en la parte superior
            miTabControl.Location = New Point(Padding, 0)
        Else
            ' Si hay un TabControl previo, ocupar la mitad del ancho y la mitad del alto del formulario
            miTabControl.Width = formulario.ClientSize.Width * 0.5 - Padding
            miTabControl.Height = (formulario.ClientSize.Height - Padding * 6) \ 2

            ' Posicionar debajo del TabControl previo, dejando margen
            Dim margen As Integer = Padding ' Espacio entre los controles
            miTabControl.Location = New Point(Padding + miTabControl.Width, tabControlPrevio.Bottom + 2 * margen)
        End If

        ' Hacer que el TabControl se redimensione junto con el formulario
        AddHandler formulario.Resize, Sub(sender As Object, e As EventArgs)
                                          If tabControlPrevio Is Nothing Then
                                              ' Si no hay TabControl previo, redimensionar ocupando todo el ancho y la mitad del alto
                                              miTabControl.Width = formulario.ClientSize.Width - Padding
                                              miTabControl.Height = formulario.ClientSize.Height \ 2
                                          Else
                                              ' Si hay un TabControl previo, redimensionar ocupando el 70% del ancho y 1/3 del alto
                                              miTabControl.Width = formulario.ClientSize.Width * 0.5 - Padding
                                              miTabControl.Height = (formulario.ClientSize.Height - Padding * 6) \ 2

                                              ' Actualizar la posición
                                              miTabControl.Location = New Point(Padding + miTabControl.Width, tabControlPrevio.Bottom + 2 * Padding)
                                          End If
                                      End Sub
    End Sub

End Class


''' <summary>
''' Clase que gestiona un GroupBox con una TextBox en un TabPage.
''' Proporciona métodos para controlar la ubicación del GroupBox en el TabPage.
''' </summary>
Public Class GroupBoxManager

    ' Declara el evento que se dispara cuando cambia la lista de GroupBox
    Public Event GroupBoxListChanged As EventHandler

    ' Propiedad privada para almacenar los GroupBox y sus TextBox asociados
    Private _listGroupBox As New List(Of GroupBox)()
    Private _textBoxes As New List(Of TextBox)()

    ''' <summary>
    ''' Propiedad pública para acceder a la lista de GroupBox con notificación de cambio.
    ''' </summary>
    Public Property ListGroupBox As List(Of GroupBox)
        Get
            Return _listGroupBox
        End Get
        Private Set(value As List(Of GroupBox))
            _listGroupBox = value
            RaiseEvent GroupBoxListChanged(Me, EventArgs.Empty)
        End Set
    End Property

    ''' <summary>
    ''' Lista pública de TextBox para acceder a cada TextBox en los GroupBox.
    ''' </summary>
    Public ReadOnly Property TextBoxes As List(Of TextBox)
        Get
            Return _textBoxes
        End Get
    End Property

    Private _groupBox As GroupBox

    ''' <summary>
    ''' Obtiene o establece el TabPage que contiene el GroupBox.
    ''' </summary>
    Public Property TabPage As TabPage

    ''' <summary>
    ''' Inicializa una nueva instancia de la clase <see cref="GroupBoxManager"/> y agrega un GroupBox con TextBox.
    ''' </summary>
    ''' <param name="titulo">Título del GroupBox.</param>
    ''' <param name="tamaño">Tamaño del GroupBox.</param>
    ''' <param name="ubicación">Ubicación del GroupBox en el TabPage.</param>
    ''' <param name="ctabPage">TabPage que contendrá el GroupBox.</param>
    Public Sub New(titulo As String, tamaño As Size, ubicación As Point, ctabPage As TabPage)
        TabPage = ctabPage
        AddNewGBoxTextBox(titulo, tamaño, ubicación)


    End Sub

    ''' <summary>
    ''' Método para agregar un nuevo GroupBox con una TextBox al TabPage dinámicamente.
    ''' </summary>
    ''' <param name="titulo">Título del nuevo GroupBox.</param>
    ''' <param name="tamaño">Tamaño del nuevo GroupBox.</param>
    ''' <param name="ubicación">Ubicación inicial del nuevo GroupBox.</param>
    Public Sub AddNewGBoxTextBox(titulo As String, tamaño As Size, ubicación As Point)
        Dim padding As New Size(10, 10)

        Dim newGroupBox As New GroupBox() With {
            .Text = titulo,
            .Size = tamaño + New Size(padding.Width * 2, padding.Height * 2),
            .Location = ubicación
        }

        Dim newTextBox As New TextBox() With {
            .Size = New Size(tamaño.Width, tamaño.Height),
            .Location = New Point(padding.Width, padding.Height * 1.5)
        }

        newGroupBox.Controls.Add(newTextBox)
        TabPage.Controls.Add(newGroupBox)

        ' Agregar el nuevo GroupBox y TextBox a sus respectivas listas
        _listGroupBox.Add(newGroupBox)
        _textBoxes.Add(newTextBox)

        ' Disparar el evento de cambio
        RaiseEvent GroupBoxListChanged(Me, EventArgs.Empty)

        ' Añadir el evento de redimensionamiento para ajustar la posición de los GroupBox
        AddHandler TabPage.FindForm().Resize, AddressOf AdjustGroupBoxPositionAndSize
    End Sub

    ''' <summary>
    ''' Ajusta la posición y el tamaño del GroupBox al redimensionar el TabPage.
    ''' </summary>
    ''' <param name="sender">Objeto que desencadena el evento (opcional).</param>
    ''' <param name="e">Datos del evento (opcional).</param>
    Private Sub AdjustGroupBoxPositionAndSize(Optional sender As Object = Nothing, Optional e As EventArgs = Nothing)
        If _listGroupBox.Count = 0 Then Exit Sub

        Dim buttons = TabPage.Controls.OfType(Of Button)().ToList()
        Dim lastButton = buttons.LastOrDefault()

        Dim initialLocation = If(lastButton IsNot Nothing,
                                 New Point(lastButton.Left, lastButton.Bottom + 10),
                                 New Point(10, 10))




        _listGroupBox(0).Location = initialLocation

        ' Ajustar la ubicación de cada GroupBox, teniendo en cuenta los bordes del TabPage
        For i As Integer = 1 To _listGroupBox.Count - 1
            Dim previousGroupBox = _listGroupBox(i - 1)
            Dim currentGroupBox = _listGroupBox(i)

            ' Calcular la nueva ubicación debajo del GroupBox anterior
            Dim newLocation = New Point(previousGroupBox.Left, previousGroupBox.Bottom + 5)

            ' Verificar si se sale de la altura visible del TabPage
            If newLocation.Y + currentGroupBox.Height > TabPage.ClientSize.Height Then
                ' Mover a la siguiente columna
                newLocation = New Point(previousGroupBox.Right + 5, 5) ' Nueva columna en la parte superior
            End If

            ' Verificar si se excede el ancho del TabPage
            If newLocation.X + currentGroupBox.Width > TabPage.ClientSize.Width Then
                MessageBox.Show("No hay suficiente espacio para ubicar más GroupBox.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit For
            End If

            ' Establecer la nueva ubicación
            currentGroupBox.Location = newLocation
        Next
    End Sub
End Class

'' Método para organizar los botones en el TabControl
'Private Sub OrganizarBotones(miTabControl As TabControl)
'    'For Each tabPage As TabPage In miTabControl.TabPages
'    '    ' Aquí puedes llamar al método de organización de botones
'    '    TabControlManager.OrganizeButtonsInGrid(tabPage, New Size(190, 55), New Size(5, 5)) ' Ajusta el tamaño y el padding según sea necesario
'    'Next
'    'If GroupBoxManager IsNot Nothing Then
'    '    'GroupBoxManager
'    'End If
'End Sub
'' Método para agregar un nuevo TabPage con un título específico.
''Public Sub AddTab(tabTitle As String)
''    Dim newTab As New TabPage(tabTitle)
''    _tabControlMenu.TabPages.Add(newTab)
''End Sub
'' Método para agregar un control a una TabPage específica por índice.
''Public Sub AddControlToTab(tabIndex As Integer, control As Control)
''    If tabIndex >= 0 AndAlso tabIndex < _tabControlMenu.TabPages.Count Then
''        _tabControlMenu.TabPages(tabIndex).Controls.Add(control)
''    Else
''        Throw New ArgumentOutOfRangeException("El índice del tab es inválido.")
''    End If
''End Sub
''Metodo para chekear si existe un pestaña
''Public Function PageExist(Tabctrl As TabControl, TabName As String) As Boolean
''    ''For Each Ctrol As Tabctrl.TabPages
''    Return True
''End Function

'' Sobrecarga: Método para agregar un control a una TabPage específica por título.
''Public Sub AddControlToTab(tabTitle As String, control As Control)
''    For Each tabPage As TabPage In _tabControlMenu.TabPages
''        If tabPage.Text = tabTitle Then
''            tabPage.Controls.Add(control)
''            Exit Sub
''        End If
''    Next
''    Throw New ArgumentException("No se encontró un tab con el título especificado.")
''End Sub
'' Método para organizar botones en un patrón de cuadrícula dentro de una TabPage.
''Public Shared Sub OrganizeButtonsInGrid(ByVal tabPage As TabPage, ByVal buttonSize As Size, ByVal padding As Size)
''    ' Obtener el TabControl que contiene el TabPage
''    Dim tabControl As TabControl = CType(tabPage.Parent, TabControl)

''    ' Obtener el tamaño del área utilizable dentro del TabControl (sin contar las pestañas)
''    Dim xPos As Integer = padding.Width
''    Dim yPos As Integer = padding.Height

''    ' El tamaño real es el del TabControl menos el tamaño de las pestañas
''    Dim tabPageWidth As Integer = tabControl.ClientSize.Width
''    Dim tabPageHeight As Integer = tabControl.ClientSize.Height - tabControl.ItemSize.Height

''    ' Filtrar solo los botones del TabPage
''    Dim buttons As IEnumerable(Of Button) = tabPage.Controls.OfType(Of Button)()

''    ' Iterar a través de los botones
''    For Each btn As Button In buttons

''        ' Asignar la ubicación del botón
''        btn.Location = New Point(xPos, yPos)
''        btn.Size = buttonSize

''        ' Actualizar la posición en Y para el siguiente botón
''        yPos += buttonSize.Height + padding.Height

''        ' Si el siguiente botón se sale del alto visible del TabPage, mover a la siguiente columna
''        If yPos + buttonSize.Height > tabPageHeight Then
''            yPos = padding.Height ' Reiniciar la posición Y para la nueva columna
''            xPos += buttonSize.Width + padding.Width ' Mover la posición en X hacia la derecha
''        End If
''    Next
''End Sub

'''' <summary>
'''' Agrega un <see cref="GroupBoxManager"/> debajo del último botón en un <see cref="TabPage"/>,
'''' manteniendo el orden de los botones organizados en la cuadrícula.
'''' </summary>
'''' <param name="tabPage">El <see cref="TabPage"/> en el que se añadirá el <see cref="GroupBox"/>.</param>
'''' <param name="titulo">El título que se mostrará en el encabezado del <see cref="GroupBox"/>.</param>
'''' <remarks>
'''' Este método determina la posición del último botón en el <see cref="TabPage"/> y coloca el <see cref="GroupBox"/>
'''' directamente debajo, con un margen adicional.
'''' </remarks>
'Public Sub AddGroupBoxAfterLastButton(tabPage As TabPage, titulo As String)
'    ' Obtener todos los botones en el TabPage para determinar la última posición
'    Dim buttons As IEnumerable(Of Button) = tabPage.Controls.OfType(Of Button)()
'    Dim lastButton As Button = buttons.OrderByDescending(Function(b) b.Bottom).FirstOrDefault()

'    ' Calcular la posición para el nuevo GroupBox, debajo del último botón
'    Dim groupBoxLocation As Point
'    If lastButton IsNot Nothing Then
'        ' Si hay un botón, colocamos el GroupBox debajo del último botón
'        groupBoxLocation = New Point(lastButton.Left, lastButton.Bottom + 10) ' 10 píxeles de espacio
'    Else
'        ' Si no hay botones, colocamos el GroupBox en la posición inicial
'        groupBoxLocation = New Point(0, 10)
'    End If

'    ' Crear y configurar el GroupBoxManager con el título y la ubicación calculada
'    Dim groupBoxManager As New GroupBoxManager(titulo, New Size(170, 25), groupBoxLocation, tabPage)
'    'groupBoxManager.AddGroupBoxToContainer(tabPage) ' Agregar el GroupBox al TabPage
'End Sub
'' Método para verificar si un control ya existe en el contenedor
'Public Shared Function ControlExists(ByVal container As Control.ControlCollection, ByVal controlName As String, ByVal controlType As Type) As Control
'    For Each ctrl As Control In container
'        If ctrl.Name = controlName AndAlso ctrl.GetType() = controlType Then
'            Return ctrl
'        End If
'    Next
'    Return Nothing
'End Function