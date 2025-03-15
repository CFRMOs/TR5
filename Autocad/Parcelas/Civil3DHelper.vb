'Imports System.Diagnostics

'Public Class Civil3DHelper
'    Public Shared Function GetCivilApp() As Object
'        Dim acadApp As AcadApplication

'        Try
'            ' Intentar obtener una instancia en ejecución de AutoCAD
'            acadApp = CType(GetObject(, "AutoCAD.Application"), AcadApplication)
'        Catch ex As Exception
'            ' Si falla, intentar crear una nueva instancia de AutoCAD
'            Try
'                acadApp = CType(CreateObject("AutoCAD.Application"), AcadApplication)
'                Debug.Print("Nueva instancia de AutoCAD creada.")
'            Catch createEx As Exception
'                Debug.Print("Error al crear una nueva instancia de AutoCAD: " & createEx.Message)
'                Return Nothing
'            End Try
'        End Try

'        If acadApp Is Nothing Then
'            Debug.Print("Error: No se pudo obtener ni crear la aplicación de AutoCAD.")
'            Return Nothing
'        End If

'        ' Intentar obtener la aplicación de Civil 3D
'        Dim versions As String() = {
'            "AeccXUiLand.AeccApplication.13.6", ' Civil 3D 2024
'            "AeccXUiLand.AeccApplication.13.0", ' Civil 3D 2023
'            "AeccXUiLand.AeccApplication.12.0", ' Civil 3D 2022
'            "AeccXUiLand.AeccApplication.11.0", ' Civil 3D 2021
'            "AeccXUiLand.AeccApplication.10.0", ' Civil 3D 2020
'            "AeccXUiLand.AeccApplication.9.0",  ' Civil 3D 2019
'            "AeccXUiLand.AeccApplication.8.0"   ' Civil 3D 2018
'        }

'        For Each version As String In versions
'            Try
'                Dim civilApp As Object = acadApp.GetInterfaceObject(version)
'                If civilApp IsNot Nothing Then
'                    Debug.Print("Loaded Civil 3D version: " & version)
'                    Return civilApp
'                End If
'            Catch comEx As System.Runtime.InteropServices.COMException
'                Debug.Print("COMException for version: " & version & " - " & comEx.Message)
'            Catch ex As Exception
'                Debug.Print("Exception for version: " & version & " - " & ex.Message)
'            End Try
'        Next

'        Debug.Print("No se pudo cargar ninguna versión de Civil 3D.")
'        Return Nothing
'    End Function
'End Class
