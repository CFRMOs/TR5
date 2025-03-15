Imports System.Text.RegularExpressions

Public Class CodigoExtractor

    ' Función para obtener el número de código AC
    Public Shared Function GetACNum(texto As String) As Double

        Return CDbl(ExtraerCodigo(Excodigo("AC-\d+", texto), "\d+"))
    End Function

    ' Función para obtener el número de código MED
    Public Shared Function GetMEDNum(texto As String) As Double
        Return CDbl(ExtraerCodigo(Excodigo("MED-\d+", texto), "\d+"))
    End Function

    ' Función para extraer un código basado en un patrón de búsqueda
    Public Shared Function ExtraerCodigo(ByVal texto As String, ByVal UsPattern As String) As String
        Dim regex As New Regex(UsPattern)
        Dim match As Match = regex.Match(texto)

        If match.Success Then
            Return match.Value
        Else
            Return String.Empty
        End If
    End Function

    ' Función para simular la obtención de un código en función de un patrón
    Public Shared Function Excodigo(ByVal Codigo As String, texto As String) As String
        'Dim texto As String = "Texto de ejemplo que puede cambiar"
        If String.IsNullOrEmpty(Codigo) Then
            Return "Codigo is Empty"
        Else
            Return ExtraerCodigo(texto, Codigo)
        End If
    End Function
    'en los nombres de los archivos se codifica y mediante a esta funccion obtenesmos el codigo representativo del acceso 
    Public Shared Function AccesoNum(Text As String) As String
        ' Buscar información del acceso en el nombre del archivo 
        Dim RefAcceso As String = String.Empty
        For Each Cod As String In {"ACC-\d+", "ACC\d+", "AC-\d+"}
            If String.IsNullOrEmpty(RefAcceso) Then
                RefAcceso = ExtraerCodigo(Text, Cod)
            Else
                Exit For
            End If
        Next
        Return RefAcceso
    End Function
End Class
