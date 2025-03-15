'Imports System.Runtime.InteropServices
'Imports Autodesk.AutoCAD.Interop
'Imports System.Windows.Forms
'Imports System.Drawing

'Public Class AutoCADEmbedder
'    ' Declaración de las funciones de la API de Windows para manipular ventanas
'    <DllImport("user32.dll", SetLastError:=True)>
'    Private Shared Function SetParent(hWndChild As IntPtr, hWndNewParent As IntPtr) As IntPtr
'    End Function

'    <DllImport("user32.dll", SetLastError:=True)>
'    Private Shared Function MoveWindow(hWnd As IntPtr, x As Integer, y As Integer, nWidth As Integer, nHeight As Boolean, bRepaint As Boolean) As Boolean
'    End Function

'    ' Variables de AutoCAD
'    Private acadApp As AcadApplication
'    Private acadDocument As AcadDocument
'    Private ReadOnly parentForm As Form
'    Private panelAutoCAD As PanelMediciones1

'    ' Constructor que recibe el formulario donde se incrustará AutoCAD
'    Public Sub New(form As Form)
'        ' Guarda el formulario
'        Me.parentForm = form

'        ' Crea y añade el panel al formulario
'        CreatePanel()

'        ' Inicializa AutoCAD
'        InitializeAutoCAD()

'        ' Habilita el panel de "Quick Properties"
'        ShowQuickProperties()

'        ' Incrusta AutoCAD en el panel creado
'        'EmbedAutoCADIntoPanel(panelAutoCAD)
'    End Sub

'    ' Método para crear el panel y agregarlo al formulario
'    Private Sub CreatePanel()
'        panelAutoCAD = New PanelMediciones1 With {
'            .Name = "PanelAutoCAD",
'            .Size = New Size(800, 600),
'            .Location = New Point(10, 10)
'        }

'        ' Añadir el panel al formulario
'        parentForm.Controls.Add(panelAutoCAD)
'    End Sub

'    ' Inicializa la aplicación de AutoCAD si no está abierta
'    Private Sub InitializeAutoCAD()
'        Try
'            ' Intenta obtener una instancia de AutoCAD en ejecución
'            acadApp = Marshal.GetActiveObject("AutoCAD.Application")
'        Catch ex As COMException
'            ' Si no hay AutoCAD en ejecución, crea una nueva instancia
'            acadApp = New AcadApplication With {
'                .Visible = True
'            }
'        End Try

'        ' Obtén el documento activo
'        acadDocument = acadApp.ActiveDocument
'    End Sub

'    '' Método para insertar el control de AutoCAD en el PanelMediciones1
'    'Private Sub EmbedAutoCADIntoPanel(panel As PanelMediciones1)
'    '    ' Obtener el handle (HWND) de la ventana de AutoCAD
'    '    Dim acadHWND As IntPtr = CType(acadApp.HWND, IntPtr)

'    '    ' Establecer el panel como el padre de la ventana de AutoCAD
'    '    SetParent(acadHWND, panel.TabName)

'    '    ' Ajustar el tamaño y posición de la ventana de AutoCAD dentro del panel
'    '    MoveWindow(acadHWND, 0, 0, panel.Width, panel.Height, True)

'    '    ' Añadir un manejador para el evento de redimensionamiento
'    '    AddHandler panel.Resize, AddressOf OnPanelResize
'    'End Sub

'    ' Evento para redimensionar la ventana de AutoCAD cuando el panel cambie de tamaño
'    Private Sub OnPanelResize(sender As Object, e As EventArgs)
'        Dim panel As PanelMediciones1 = CType(sender, PanelMediciones1)
'        Dim acadHWND As IntPtr = CType(acadApp.HWND, IntPtr)
'        MoveWindow(acadHWND, 0, 0, panel.Width, panel.Height, True)
'    End Sub

'    ' Habilita el panel de "Quick Properties"
'    Private Sub ShowQuickProperties()
'        ' Activa las propiedades rápidas
'        acadApp.ActiveDocument.SetVariable("QUICKPROPERTIES", 1)
'    End Sub

'End Class
