Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Module ExcelFunctions
    ' Declaraciones para las funciones de la API de Windows
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function LockWindowUpdate(ByVal hWndLock As IntPtr) As Boolean
    End Function

    Private ReadOnly VBAVISIBLE As Boolean

    ' Método para apagar las funcionalidades
    Public Sub TurnEverythingOff()
        Dim VBEHwnd As IntPtr
        Try
            Dim excelApp As Application = CType(GetObject(, "Excel.Application"), Application)

            With excelApp
                .Calculation = XlCalculation.xlCalculationManual
                .EnableEvents = False
                .DisplayAlerts = False
                .ScreenUpdating = False

                VBEHwnd = FindWindow("wndclass_desked_gsk", .Caption)
                If VBEHwnd <> IntPtr.Zero Then
                    LockWindowUpdate(VBEHwnd)
                End If
            End With
        Catch ex As Exception
            Console.WriteLine("Error en TurnEverythingOff: " & ex.Message)
        End Try
    End Sub

    ' Método para encender las funcionalidades
    Public Sub TurnEverythingOn()
        Try
            Dim excelApp As Application = CType(GetObject(, "Excel.Application"), Application)

            With excelApp
                .EnableEvents = True
                .DisplayAlerts = True
                .ScreenUpdating = True
                .Calculation = XlCalculation.xlCalculationAutomatic
            End With

            LockWindowUpdate(IntPtr.Zero)
        Catch ex As Exception
            Console.WriteLine("Error en TurnEverythingOn: " & ex.Message)
        End Try
    End Sub
End Module
