Imports System.Diagnostics
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class ACAdHelpers
    ' Class variables
    Public ReadOnly List_ALing As List(Of Alignment)
    Public ReadOnly List_ALingName As List(Of String)
    Public ReadOnly List_alingCount As Integer
    Public Alignment As Alignment
    Public selectedLayers As New List(Of String)
    Public nombresCapas As String()
    Public DCapas As New Dictionary(Of String, String())

    'Private ReadOnly filePath As String = ""
    Public WithEvents ThisDrawing As Document
    'Private WithEvents AcadApp As Application 'acad 
    Public WithEvents Form As PolylineMG
    Public WBFilePath As String
    ' Constructor
    Public Sub New()
        Try
            List_ALing = GetAllAlignments(List_ALingName)
            List_alingCount = List_ALing.Count

            ThisDrawing = Application.DocumentManager.MdiActiveDocument
            'filePath = ThisDrawing.Name
            ' Agregar manejadores para eventos de activación de documentos
            AddHandler Application.DocumentManager.DocumentActivated, AddressOf OnDocumentActivated
            AddHandler Application.DocumentManager.DocumentToBeDestroyed, AddressOf OnDocumentClosed
            'Ensure InitializeComponent() Is called first
            'Cargar los datos utilizando la función ProcesarHandlesDesdeDoc desde AutoCADHelper2
            If FUrlCarpetaProcessor.ArchivoExiste("C:\Users\typsa\Desktop\1.0-Mediciones\Resumen Hormigones Typsa Med-14-R1 06-11-2024 1116 CF.xlsm") Then
                WBFilePath = "C:\Users\typsa\Desktop\1.0-Mediciones\Resumen Hormigones Typsa Med-14-R1 06-11-2024 1116 CF.xlsm"
            ElseIf FUrlCarpetaProcessor.ArchivoExiste("D:\Desktop\Typsa - Las Placetas\1.0-Mediciones\Resumen Hormigones Typsa Med-14-R1 22-10-2024 1139 CF.xlsm") Then
                WBFilePath = "D:\Desktop\Typsa - Las Placetas\1.0-Mediciones\Resumen Hormigones Typsa Med-14-R1 22-10-2024 1139 CF.xlsm"
            End If
        Catch ex As Exception
            Debug.WriteLine("Error al iniciar AutoCAD: " & ex.Message)
            Throw New Exception("No se pudo iniciar AutoCAD: " & ex.Message)
        End Try
    End Sub
    Public Sub CheckThisDrowing()
        If ThisDrawing IsNot Nothing Then
            Application.DocumentManager.MdiActiveDocument = ThisDrawing
        Else
            'MessageBox.Show("El documento original no está disponible.")
            Return
        End If
    End Sub
    ' Evento para activar el formulario cuando se activa un documento
    Private Sub OnDocumentActivated(sender As Object, e As DocumentCollectionEventArgs)
        If Form IsNot Nothing And Not Form?.IsDisposed AndAlso e.Document Is ThisDrawing Then
            Form?.Activate()
            Form?.Show()
        End If
    End Sub

    ' Evento para cerrar el formulario cuando se cierra un documento
    Private Sub OnDocumentClosed(sender As Object, e As DocumentCollectionEventArgs)
        If Form IsNot Nothing AndAlso e.Document Is ThisDrawing Then
            Form?.Close()
        End If
    End Sub

    ' Manejador de activación de AutoCAD (opcional si quieres manejar también este evento)
    'Private Sub AcadApp_AppActivate() Handles AcadApp
    '    If Form IsNot Nothing Then
    '        Form.Activate()
    '        Form.Show()
    '    End If
    'End Sub
    ' TabName document close event
    Private Sub ThisDrawing_BeginDocumentClose(sender As Object, e As DocumentBeginCloseEventArgs) Handles ThisDrawing.BeginDocumentClose
        If Form IsNot Nothing Then
            Form?.Close()
            Form = Nothing
        End If
    End Sub

    'Private Sub AcadApp_AppActivate() Handles AcadApp.AppActivate
    '    If AcadApp.DocumentManager.MdiActiveDocument.Name = ThisDrawing.Name Then
    '        Form?.Activate()
    '        Form?.BringToFront()
    '    End If
    'End Sub

    'Private Sub Form_Activated(sender As Object, e As EventArgs) Handles Form.Activated
    '    'ThisDrawingAcadApp.DocumentManager.MdiActiveDocument = ThisDrawing

    'End Sub
End Class
