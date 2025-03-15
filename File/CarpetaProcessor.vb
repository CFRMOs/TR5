Imports System.IO
Imports System.Text.RegularExpressions

Public Class CarpetaProcessor

    ' Función principal para convertir la colección de archivos en un array bidimensional
    Public Function TCARPETA(ByVal nCarpeta As String, ByVal ParamArray EXT() As String) As Object(,)
        Dim LFiles As New List(Of Object())
        Dim resultado(,) As Object

        ' Llamamos a la función recursiva para llenar la colección
        For Each extension In EXT
            RecursivelyCollectFiles(nCarpeta, extension, LFiles)
        Next

        ' Convertimos la lista en un array bidimensional
        resultado = CollectionToArray(LFiles)

        ' Devolvemos el array con los resultados
        Return resultado
    End Function

    ' Función para obtener archivos de una carpeta específica
    Public Function GetFilesCollection(ByVal vBasePath As String, ByVal EXT As String) As List(Of Object())
        Dim archivos As New List(Of Object())
        ' Obtener todos los archivos de la carpeta base
        Dim directoryInfo As New DirectoryInfo(vBasePath)
        Dim files As FileInfo() = directoryInfo.GetFiles()

        ' Devolvemos la colección de archivos filtrados
        Return ConstructCollection(files, EXT)
    End Function

    ' Función recursiva para recorrer todas las carpetas y subcarpetas y acumular archivos en una lista
    Private Sub RecursivelyCollectFiles(ByVal nCarpeta As String, ByVal EXT As String, ByRef LFiles As List(Of Object()))
        Dim directoryInfo As New DirectoryInfo(nCarpeta)

        ' Agregar archivos de la carpeta actual a la lista
        AddFilesToCollection(LFiles, directoryInfo.GetFiles(), EXT)

        ' Bucle para recorrer todas las subcarpetas
        For Each subDir In directoryInfo.GetDirectories()
            RecursivelyCollectFiles(subDir.FullName, EXT, LFiles)
        Next
    End Sub

    ' Función para agregar archivos a la lista desde una carpeta específica
    Private Sub AddFilesToCollection(ByRef LFiles As List(Of Object()), ByVal archivos As FileInfo(), ByVal EXT As String)
        EXT = AddPExt(EXT) ' Aseguramos que EXT empieza con un punto

        ' Filtrar solo archivos con la extensión especificada
        For Each archivo As FileInfo In archivos
            If archivo.Extension.ToLower() = EXT.ToLower() Then
                ' Añadimos los detalles del archivo a la lista
                Dim archivoInfo As Object() = {
                    archivo.Name,
                    archivo.FullName,
                    EXT,
                    archivo.LastWriteTime,
                    ExtraerCodigo(archivo.FullName, "Cubicación \d+")
                }
                LFiles.Add(archivoInfo)
            End If
        Next
    End Sub

    ' Función privada para construir la lista de resultados a partir de los archivos filtrados
    Private Function ConstructCollection(ByVal archivos As FileInfo(), ByVal EXT As String) As List(Of Object())
        Dim resultados As New List(Of Object())
        EXT = AddPExt(EXT) ' Aseguramos que EXT empieza con un punto

        For Each archivo As FileInfo In archivos
            If archivo.Extension.ToLower() = EXT.ToLower() Then
                ' Añadir los detalles del archivo a la lista
                resultados.Add({
                    archivo.Name,
                    archivo.FullName,
                    EXT,
                    archivo.LastWriteTime
                })
            End If
        Next

        Return resultados
    End Function

    ' Función para convertir una lista en un array bidimensional
    Private Function CollectionToArray(ByVal LFiles As List(Of Object())) As Object(,)
        Dim resultado(,) As Object

        If LFiles.Count > 0 Then
            ReDim resultado(LFiles.Count - 1, 4)
            For i As Integer = 0 To LFiles.Count - 1
                For j As Integer = 0 To 4
                    resultado(i, j) = LFiles(i)(j)
                Next
            Next
        Else
            resultado = New Object(,) {} ' Array vacío si no hay archivos
        End If

        Return resultado
    End Function

    ' Función para asegurar que la extensión comienza con un punto
    Private Function AddPExt(ByVal EXT As String) As String
        If Not EXT.StartsWith(".") Then
            EXT = "." & EXT
        End If
        Return EXT
    End Function

    ' Función para extraer el código de cubicación (debes adaptarla según tus necesidades)
    Private Function ExtraerCodigo(ByVal path As String, ByVal pattern As String) As String
        Dim match As Match = Regex.Match(path, pattern)
        If match.Success Then
            Return match.Value
        Else
            Return String.Empty
        End If
    End Function

End Class
Public Class FUrlCarpetaProcessor
    'FUrlCarpetaProcessor.ArchivoExiste
    Public Shared Function ArchivoExiste(ruta As String) As Boolean
        ' Verifica si el archivo existe usando la clase File
        If File.Exists(ruta) Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
