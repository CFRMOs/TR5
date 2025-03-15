Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq
Imports System.Runtime.InteropServices.WindowsRuntime
Imports System.Windows.Forms

Public Class TabControlManager
    Private ReadOnly _tabControl As TabControl

    ' Constructor que recibe el TabControl en el cual se van a agregar los tabs y controles.
    Public Sub New(tabControl As TabControl)
        _tabControl = tabControl
    End Sub

    ' Método para agregar un nuevo TabPage con un título específico.
    Public Sub AddTab(tabTitle As String)
        Dim newTab As New TabPage(tabTitle)
        _tabControl.TabPages.Add(newTab)
    End Sub

    ' Método para agregar un control a una TabPage específica por índice.
    Public Sub AddControlToTab(tabIndex As Integer, control As Control)
        If tabIndex >= 0 AndAlso tabIndex < _tabControl.TabPages.Count Then
            _tabControl.TabPages(tabIndex).Controls.Add(control)
        Else
            Throw New ArgumentOutOfRangeException("El índice del tab es inválido.")
        End If
    End Sub
    'Metodo para chekear si existe un pestaña
    Public Function PageExist(Tabctrl As TabControl, TabName As String) As Boolean
        ''For Each Ctrol As Tabctrl.TabPages
        Return True
    End Function

    ' Sobrecarga: Método para agregar un control a una TabPage específica por título.
    Public Sub AddControlToTab(tabTitle As String, control As Control)
        For Each tabPage As TabPage In _tabControl.TabPages
            If tabPage.Text = tabTitle Then
                tabPage.Controls.Add(control)
                Exit Sub
            End If
        Next
        Throw New ArgumentException("No se encontró un tab con el título especificado.")
    End Sub
    ' Método para organizar botones en un patrón de cuadrícula dentro de una TabPage.
    Public Shared Sub OrganizeButtonsInGrid(ByVal tabPage As TabPage, ByVal buttonSize As Size, ByVal padding As Size)
        ' Obtener el TabControl que contiene el TabPage
        Dim tabControl As TabControl = CType(tabPage.Parent, TabControl)

        ' Obtener el tamaño del área utilizable dentro del TabControl (sin contar las pestañas)
        Dim xPos As Integer = padding.Width
        Dim yPos As Integer = padding.Height

        tabControl.Refresh()

        ' El tamaño real es el del TabControl menos el tamaño de las pestañas
        Dim tabPageWidth As Integer = tabControl.ClientSize.Width
        Dim tabPageHeight As Integer = tabControl.ClientSize.Height - tabControl.ItemSize.Height

        ' Filtrar solo los botones del TabPage
        Dim buttons As IEnumerable(Of Button) = tabPage.Controls.OfType(Of Button)()

        ' Iterar a través de los botones
        For Each btn As Button In buttons
            ' Asignar la ubicación del botón
            btn.Location = New Point(xPos, yPos)
            btn.Size = buttonSize

            ' Actualizar la posición en Y para el siguiente botón
            yPos += buttonSize.Height + padding.Height

            ' Si el siguiente botón se sale del alto visible del TabPage, mover a la siguiente columna
            If yPos + buttonSize.Height > tabPageHeight Then
                yPos = padding.Height ' Reiniciar la posición Y para la nueva columna
                xPos += buttonSize.Width + padding.Width ' + padding.Width ' Mover la posición en X hacia la derecha
            End If
        Next
    End Sub
    ' Método para verificar si un control ya existe en el contenedor
    Public Shared Function ControlExists(ByVal container As Control.ControlCollection, ByVal controlName As String, ByVal controlType As Type) As Control
        For Each ctrl As Control In container
            If ctrl.Name = controlName AndAlso ctrl.GetType() = controlType Then
                Return ctrl
            End If
        Next
        Return Nothing
    End Function
End Class
'Clase para representar una utilida de una function o otra clase en windows form en un panel de utilidades 
Public Class TabsMenu
    ' Public Property
    Public Property TabPage As TabPage = Nothing
    Public Property TabCtrol As TabControl = Nothing
    Public Property Button As Button = Nothing
    Public Property Buttons As New Dictionary(Of String, Button)

    Private EventsHandler As New List(Of EventHandlerClass)

    Public EventHandlerParmetrict As New List(Of EventHandlerParametricClass)
    ' Declarar la variable privada para almacenar las acciones de los botones
    Private _buttonActions As New Dictionary(Of String, Action)

    ' Definir un evento para detectar cuando cambien las acciones del botón
    Public Event ButtonActionsChanged As EventHandler

    ' Propiedad pública que permite el acceso a las acciones del botón
    Public Property ButtonActions As Dictionary(Of String, Action)
        Get
            Return _buttonActions

        End Get

        Set(value As Dictionary(Of String, Action))
            ' Comparar el contenido del diccionario para evitar asignaciones innecesarias
            If Not DictionariesAreEqual(_buttonActions, value) Then
                _buttonActions = value

                ' Disparar el evento cuando cambian las acciones de los botones
                RaiseEvent ButtonActionsChanged(Me, EventArgs.Empty)
                SuscripActions()
            End If
        End Set
    End Property
    ' Función para comparar dos diccionarios
    Private Function DictionariesAreEqual(dict1 As Dictionary(Of String, Action), dict2 As Dictionary(Of String, Action)) As Boolean
        ' Comparar primero el número de entradas en ambos diccionarios
        If dict1.Count <> dict2.Count Then
            Return False
        End If

        ' Comparar cada entrada por clave y valor
        For Each kvp In dict1
            If Not dict2.ContainsKey(kvp.Key) OrElse dict2(kvp.Key) IsNot kvp.Value Then
                Return False
            End If
        Next

        Return True
    End Function

    ' Declare the event
    Public Event TabNameChanged As EventHandler
    Public Event ButtonNameChanged As EventHandler

    ' Private backing fields
    Private _tabName As String = ""
    Private _buttonName As String = ""

    ' Public Property for TabName with change notification
    Public Property TabName As String
        Get
            Return _tabName
        End Get
        Set(value As String)
            ' Trigger the event only if the value changes
            If _tabName <> value Then
                _tabName = value

                ' Trigger the HandleChanged event
                RaiseEvent TabNameChanged(Me, EventArgs.Empty)

                ' Ensure TabCtrol is not Nothing before using it
                If TabCtrol IsNot Nothing Then
                    TabPage = AddDrenajeUTLPage(_tabName, TabCtrol)
                End If
            End If
        End Set
    End Property
    ' Optional: Property for ButtonName if needed
    Public Property ButtonName As String
        Get
            Return _buttonName
        End Get
        Set(value As String)
            If _buttonName <> value Then
                _buttonName = value
                ' Trigger the HandleChanged event
                RaiseEvent ButtonNameChanged(Me, EventArgs.Empty)
                Button = AddDrenajeUTLButton(TabPage, _buttonName)
                Dim BSize As New Size() With {.Height = 48, .Width = 180
                }
                Dim PaddingSize As New Size() With {.Height = 3, .Width = 5
                }
                TabControlManager.OrganizeButtonsInGrid(TabPage, BSize, PaddingSize)
                'Buttons.Add(_buttonName, Button)
            End If
        End Set
    End Property

    'constructor 
    ' Constructor that initializes the tab name and button name
    Public Sub New(cControl As TabControl) 'cTabName As String, cButtonName As String, 
        TabCtrol = cControl
    End Sub
    ' Asume que tienes un diccionario o lista para los botones
    Public Function CreateButtonAndGet(ByVal name As String) As Button
        Me.ButtonName = name
        If Not Me.Buttons.ContainsKey(name) Then
            Me.Buttons.Add(name, Button)
        End If
        Return Me.Buttons(name)
    End Function
    'rutina encargada de asignar las acciones imediatamente estas son asignadas 
    Public Sub SuscripActions()
        'simplificar con el uso de array y for y una funcion que genere las acciones(Rutina) desde un texto 
        ' Iterar sobre el diccionario para crear los botones y asignarles las acciones correspondientes
        For Each buttonEntry In ButtonActions
            EventsHandler.Add(New EventHandlerClass(CreateButtonAndGet(buttonEntry.Key), "click", buttonEntry.Value))
        Next
    End Sub
    'crear pagina con boton
    Public Shared Function AddDrenajeUTLPage(ByVal TabName As String, Optional ByRef TabCtrol As TabControl = Nothing) As TabPage
        ' Verifica si el TabControl tiene una página seleccionada
        Try
            If TabCtrol IsNot Nothing Then
                ' Accede a la página activa
                For Each Tab As TabPage In TabCtrol.TabPages
                    If Tab.Text = TabName Then
                        Return Tab
                    End If
                Next
                ' Crear pestaña (TabPage)
                Dim tabPage As New TabPage(TabName)
                ' Agregar las pestañas al TabControl
                TabCtrol.TabPages.Add(tabPage)
                Return tabPage

            Else
                MessageBox.Show("el TabControl no existe!!")
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("Error:" & ex.Message)
            Return Nothing
        End Try

    End Function
    Public Shared Function AddDrenajeUTLButton(ByVal TabPage As TabPage, ByVal ButtonName As String) As Button
        ' Verifica si el TabControl tiene una página seleccionada
        Try
            If TabPage IsNot Nothing Then
                ' Accede a la página activa
                For Each Ctrol As Control In TabPage.Controls
                    If Ctrol.Text = ButtonName Then
                        Return Ctrol
                    End If
                Next
                ' Crear pestaña (TabPage)
                Dim Button As New Button With {
                    .Text = ButtonName
                }
                ' Agregar las pestañas al TabControl
                TabPage.Controls.Add(Button)
                Return Button

            Else
                MessageBox.Show("el TabControl no existe!!")
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("Error:" & ex.Message)
            Return Nothing
        End Try
    End Function
End Class

Public Class TabDGViews
    ' Public Property
    Public Property TabPage As TabPage = Nothing
    Public Property TabCtrol As TabControl = Nothing
    Public Property DGView As DataGridView = Nothing
    Public Property HandleProcessor As HandleDataProcessor
    Public Property ParentForm As Form = Nothing
    Public Property AcadHelp As ACAdHelpers = Nothing
    Public Property EventHandler As DataGridViewEventHandler
    Public Property ColumnFormats As List(Of String)

    ' Declare the event
    Public Event TabNameChanged As EventHandler

    Public Event DGViewNameChanged As EventHandler

    ' Private backing fields
    Private _tabName As String = ""

    Private _DGViewName As String = ""

    ' Public Property with change notification
    Public Property TabName As String
        Get
            Return _tabName
        End Get
        Set(value As String)
            ' Trigger the event only if the value changes
            If _tabName <> value Then
                _tabName = value

                ' Trigger the HandleChanged event
                RaiseEvent TabNameChanged(Me, EventArgs.Empty)

                ' Ensure TabCtrol is not Nothing before using it
                If TabCtrol IsNot Nothing Then
                    TabPage = AddDGViewPage(_tabName, TabCtrol)
                End If
            End If
        End Set
    End Property

    Public Property DGViewName As String
        Get
            Return _DGViewName
        End Get
        Set(value As String)
            ' Trigger the event only if the value changes
            If _DGViewName <> value Then
                _DGViewName = value

                ' Trigger the HandleChanged event
                RaiseEvent TabNameChanged(Me, EventArgs.Empty)

                ' Ensure TabCtrol is not Nothing before using it
                If TabCtrol IsNot Nothing Then
                    DGView = AddDGView(_DGViewName, TabPage)
                End If
            End If
        End Set
    End Property


    ' Constructor that initializes the tab name and button name
    Public Sub New(cControl As TabControl, cColumnFormats As List(Of String), ByRef CAcadHelp As ACAdHelpers,
                   Optional ByRef parentForm As Form = Nothing, Optional ByRef HandleDataProcessor As HandleDataProcessor = Nothing) 'cTabName As String, cButtonName As String, 
        HandleProcessor = HandleDataProcessor
        parentForm = parentForm
        AcadHelp = CAcadHelp
        TabCtrol = cControl
        ColumnFormats = cColumnFormats

    End Sub
    'constructor 

    'crear pagina con boton
    Public Shared Function AddDGViewPage(ByVal TabName As String, Optional ByRef TabCtrol As TabControl = Nothing) As TabPage
        ' Verifica si el TabControl tiene una página seleccionada
        Try
            If TabCtrol IsNot Nothing Then
                ' Accede a la página activa

                For Each Tab As TabPage In TabCtrol.TabPages
                    If Tab.Text = TabName Then
                        Return Tab
                    End If
                Next
                ' Crear pestaña (TabPage)
                Dim tabPage As New TabPage(TabName)
                ' Agregar las pestañas al TabControl
                TabCtrol.TabPages.Add(tabPage)
                Return tabPage

            Else
                MessageBox.Show("el TabControl no existe!!")
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("Error:" & ex.Message)
            Return Nothing
        End Try
    End Function

    'crear pagina con boton
    Public Function AddDGView(ByVal DGViewName As String, TabPage As TabPage) As DataGridView
        ' Verifica si el TabControl tiene una página seleccionada
        Try
            If TabPage IsNot Nothing Then
                ' Accede a la página activa
                For Each Ctrl As Control In TabPage.Controls
                    If Ctrl.Text = DGViewName AndAlso TypeOf Ctrl Is DataGridView Then
                        Return Ctrl
                    End If
                Next
                '' Crear pestaña (TabPage)
                Dim DGView As New DataGridView With {
                    .Name = DGViewName,
                    .Dock = DockStyle.Fill ' Hace que el DataGridView se ajuste al tamaño del TabPage
                    } 'From { .name = DGViewName}

                DataGridViewHelper.Addcolumns(DGView, HandleDataProcessor.Headers)

                DateViewSet.ConfigureDataGridView(DGView, ColumnFormats)

                TabPage.Controls.Add(DGView)

                Dim Form As Form = DGView.FindForm()

                EventHandler = New DataGridViewEventHandler(DGView, AcadHelp, Form, HandleProcessor)
                Return DGView
            Else
                MessageBox.Show("el TabControl no existe!!")
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("Error:" & ex.Message)
            Return Nothing
        End Try
    End Function
End Class

Public Class TxtBoxWithLabelMenu
    Public Property TabPage As TabPage = Nothing
    Public Property TabCtrol As TabControl = Nothing
    Public Property TextBox As TextBox = Nothing
    Public Property Label As Label = Nothing
    Private ReadOnly EventsHandler As New List(Of EventHandlerClass)
    Public Event PositionEventHandler As EventHandler
    Public _position As Point
    Public Property Position As Point
        Get
            Position = _position
        End Get
        Set(value As Point)
            If Not _position.Equals(value) Then

            End If
        End Set
    End Property

    Public Sub New()

    End Sub
    'crear grupo que contenga un label descriptivo sobre un textbox.
    'el textbox y el lebel tendran una pposicion del grupo y la posicion de la clase sera la del grupo

    Public Function AddTextbox(ByVal TextboxName As String, TBoxSize As Size) As TextBox
        ' Verifica si el TabControl tiene una página seleccionada
        Try
            If TabPage IsNot Nothing Then
                ' Accede a la página activa
                For Each Ctrl As Control In TabPage.Controls
                    If Ctrl.Name = TextboxName AndAlso TypeOf Ctrl Is TextBox Then
                        Return Ctrl
                    End If
                Next
                '' Crear pestaña (TabPage)
                'Dim TBoxSize As New Size() With {.Width = 45, .Height = 150}
                Dim TextBox As New TextBox() With {.Name = TextboxName, .Size = TBoxSize}
                Return TextBox
            Else
                MessageBox.Show("el TabControl no existe!!")
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("Error:" & ex.Message)
            Return Nothing
        End Try
    End Function
    Public Function AddLabel(ByVal LabelName As String, TBoxSize As Size) As Label
        ' Verifica si el TabControl tiene una página seleccionada
        Try
            If TabPage IsNot Nothing Then
                ' Accede a la página activa
                For Each Ctrl As Control In TabPage.Controls
                    If Ctrl.Name = LabelName AndAlso TypeOf Ctrl Is TextBox Then
                        Return Ctrl
                    End If
                Next
                '' Crear pestaña (TabPage)
                'Dim TBoxSize As New Size() With {.Width = 45, .Height = 150}
                Dim label As New Label() With {.Name = LabelName, .Size = TBoxSize}
                Return label
            Else
                MessageBox.Show("el TabControl no existe!!")
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("Error:" & ex.Message)
            Return Nothing
        End Try
    End Function
    Public Function CrearGroup(ByVal GroupName As String, TBoxSize As Size) As Label
        ' Verifica si el TabControl tiene una página seleccionada
        Try
            If TabPage IsNot Nothing Then
                ' Accede a la página activa
                For Each Ctrl As Control In TabPage.Controls
                    If Ctrl.Name = GroupName AndAlso TypeOf Ctrl Is TextBox Then
                        Return Ctrl
                    End If
                Next
                '' Crear pestaña (TabPage)
                'Dim TBoxSize As New Size() With {.Width = 45, .Height = 150}
                Dim label As New Label() With {.Name = GroupName, .Size = TBoxSize}
                Return label
            Else
                MessageBox.Show("el TabControl no existe!!")
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("Error:" & ex.Message)
            Return Nothing
        End Try
    End Function
End Class

' Ejemplo en el Formulario:
Public Class Form1
    Inherits Form
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Crear una instancia de TabControl
        ' Configurar propiedades del TabControl
        Dim myTabControl As New TabControl With {
            .Size = New Size(400, 200),  ' Tamaño del TabControl
            .Location = New Point(20, 20)  ' Ubicación del TabControl en el formulario
            }

        ' Agregar el TabControl al formulario
        Me.Controls.Add(myTabControl)

        ' Crear algunas pestañas (TabPages)
        Dim tabPage1 As New TabPage("Pestaña 1")
        Dim tabPage2 As New TabPage("Pestaña 2")

        ' Agregar las pestañas al TabControl
        myTabControl.TabPages.Add(tabPage1)
        myTabControl.TabPages.Add(tabPage2)

        ' Agregar un botón a la primera pestaña
        Dim button1 As New Button With {
            .Text = "Botón en Pestaña 1",
            .Size = New Size(120, 30),
            .Location = New Point(30, 30)
        }
        tabPage1.Controls.Add(button1)

        ' Agregar un TextBox a la segunda pestaña
        Dim textBox1 As New TextBox With {
            .Size = New Size(200, 30),
            .Location = New Point(30, 30)
        }
        tabPage2.Controls.Add(textBox1)
    End Sub
End Class

