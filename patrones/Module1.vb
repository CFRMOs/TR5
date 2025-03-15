Imports System.Text.RegularExpressions

Public Module ModulePrueba
    Sub Main()
        Dim texto As String = "AGO-HLP-AC-04-TRT-012-R0"
        Dim patron As String = "AGO-HLP-AC-00-TRT-000-R0"
        Dim resultado As String = ModificarPatron(texto, patron)
        Console.WriteLine(resultado)
    End Sub

    Public Function ModificarPatron(ByVal texto As String, ByVal patron As String) As String
        ' Dividir ambos en partes usando el separador "-"
        Dim textoPartes As String() = texto.Split("-"c)
        Dim patronPartes As String() = patron.Split("-"c)

        ' Reemplazar las partes del patrón con las correspondientes del texto
        For i As Integer = 0 To patronPartes.Length - 1
            ' Si el componente del patrón es un número, reemplazarlo con el componente correspondiente del texto
            If IsNumeric(patronPartes(i)) Then
                patronPartes(i) = textoPartes(i)
            End If
        Next

        ' Reconstruir el patrón modificado
        Dim patronModificado As String = String.Join("-", patronPartes)

        Return patronModificado
    End Function

    ' Función auxiliar para verificar si una cadena es numérica
    Function IsNumeric(ByVal value As String) As Boolean
        Return Regex.IsMatch(value, "^\d+$")
    End Function
End Module
