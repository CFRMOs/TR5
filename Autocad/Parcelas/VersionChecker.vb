'Imports Autodesk.AutoCAD.Runtime
'Imports WinRegistry = Microsoft.Win32.Registry
'Imports WinRegistryKey = Microsoft.Win32.RegistryKey

'Public Class VersionChecker

'    <CommandMethod("CheckC3DVersion")>
'    Public Sub CheckC3DVersion()
'        Try
'            ' Inicializa el objeto COM de AutoCAD
'            Dim acadApp As AcadApplication = CType(Application.AcadApplication, AcadApplication)

'            ' Obtiene la clave del registro para la versión de Civil 3D
'            Dim regKeyPath As String = "HKEY_LOCAL_MACHINE\"
'            Dim productKey As String = GetProductKey()
'            If String.IsNullOrEmpty(productKey) Then
'                Throw New Exception("Product key not found.")
'            End If

'            regKeyPath &= productKey
'            Dim releaseVersion As String = GetRegistryValue(regKeyPath, "Release")

'            If String.IsNullOrEmpty(releaseVersion) Then
'                Throw New Exception("Release version not found.")
'            End If

'            ' Obtiene la versión abreviada
'            Dim versionShort As String = releaseVersion.Substring(0, releaseVersion.IndexOf(".", releaseVersion.IndexOf(".") + 1))
'            Dim progId As String = "AeccXUiLand.AeccApplication." & versionShort

'            ' Obtiene el objeto de la aplicación de Civil 3D
'            Dim civ3dApp As Object = acadApp.GetInterfaceObject(progId)
'            Dim aeccDoc As Object = civ3dApp.ActiveDocument

'            ' Muestra la información de la versión
'            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
'                "Civil 3D Version: " & versionShort & vbCrLf &
'                "Active Document Name: " & aeccDoc.Name & vbCrLf)

'        Catch ex As Exception
'            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
'                "Error: " & ex.Message & vbCrLf)
'        End Try
'    End Sub

'    Private Function GetProductKey() As String
'        Try
'            Dim acadVer As String = Application.GetSystemVariable("ACADVER").ToString()
'            Return "SOFTWARE\Autodesk\AutoCAD\" & acadVer
'        Catch ex As Exception
'            Return String.Empty
'        End Try
'    End Function

'    Private Function GetRegistryValue(keyPath As String, valueName As String) As String
'        Try
'            Using key As WinRegistryKey = WinRegistry.LocalMachine.OpenSubKey(keyPath)
'                If key IsNot Nothing Then
'                    Return key.GetValue(valueName).ToString()
'                End If
'            End Using
'        Catch ex As Exception
'            Return String.Empty
'        End Try

'        Return String.Empty
'    End Function

'End Class
