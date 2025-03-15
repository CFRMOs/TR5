Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class ExcelOptimizer
    Implements IDisposable

    ' Instancia compartida de Excel a nivel de clase
    Private Shared ExcelAppInstance As Application
    Private ReadOnly ExcelApp As Application
    Private _disposed As Boolean = False
    Private Shared isOwner As Boolean = False ' Para verificar si este objeto es el dueño de la instancia

    ' Constructor que recibe opcionalmente una instancia de Excel.Application
    Public Sub New(Optional excelApp As Application = Nothing)
        If excelApp Is Nothing Then
            If ExcelAppInstance Is Nothing Then
                ' Si no hay una instancia compartida, se crea una nueva
                Dim ExcelInstances As Process() = Process.GetProcessesByName("EXCEL")
                If ExcelInstances.Length = 0 Then
                    ExcelAppInstance = New Application With {.Visible = True}
                    isOwner = True ' Este objeto es el dueño de la instancia
                Else
                    Try
                        ExcelAppInstance = TryCast(Marshal.GetActiveObject("Excel.Application"), Application)
                    Catch ex As COMException
                        ' Manejar la excepción si GetActiveObject falla
                        ExcelAppInstance = New Application With {.Visible = True}
                        isOwner = True ' Este objeto es el dueño de la instancia
                    End Try
                End If
            End If
            Me.ExcelApp = ExcelAppInstance ' Usa la instancia compartida
        Else
            Me.ExcelApp = excelApp ' Usa la instancia proporcionada
        End If
    End Sub

    ' Turn off automatic calculations, events, and screen updates for performance
    Public Sub TurnEverythingOff()
        Try
            With ExcelApp
                .Calculation = XlCalculation.xlCalculationManual
                .EnableEvents = False
                .DisplayAlerts = False
                .ScreenUpdating = False
            End With
        Catch ex As Exception
            Debug.WriteLine($"Error turning Excel settings off: {ex.Message}")
        End Try
    End Sub

    ' Restore Excel to normal operation
    Public Sub TurnEverythingOn()
        Try
            With ExcelApp
                .EnableEvents = True
                .DisplayAlerts = True
                .ScreenUpdating = True
                .Calculation = XlCalculation.xlCalculationAutomatic
            End With
        Catch ex As Exception
            Debug.WriteLine($"Error restoring Excel settings: {ex.Message}")
        End Try
    End Sub

    ' Patrón Dispose para liberar objetos COM de Excel
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    ' Núcleo del patrón Dispose
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not _disposed Then
            If disposing Then
                ' Solo liberar ExcelApp si este objeto es el dueño de la instancia
                If ExcelApp IsNot Nothing AndAlso isOwner Then
                    Try
                        ExcelApp.Quit() ' Cerrar Excel si esta instancia es la responsable
                        Marshal.ReleaseComObject(ExcelApp)
                        ExcelAppInstance = Nothing ' Borrar la referencia compartida
                    Catch ex As Exception
                        Debug.WriteLine($"Error releasing Excel COM object: {ex.Message}")
                    End Try
                End If
            End If
            _disposed = True
        End If
    End Sub

    ' Destructor/Finalizer
    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub
End Class
