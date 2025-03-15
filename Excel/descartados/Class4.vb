'Imports System.Collections.Generic
'Imports Microsoft.Office.Interop.Excel

'Public Class FiltroMatriz

'    ' Método para filtrar una matriz basada en varios criterios
'    Public Function FiltrarConjuntos(matriz As Range, ParamArray criterios() As Boolean(,)) As Object(,)
'        Dim i As Integer, j As Integer
'        Dim numFilas As Integer = matriz.Rows.Count
'        Dim numCols As Integer = matriz.Columns.Count
'        Dim coincide As Boolean
'        Dim matrizValores(,) As Object = matriz.Value2
'        Dim resultadosTemp As New List(Of Object())

'        ' Iterar sobre cada fila de la Matriz
'        For i = 1 To numFilas
'            coincide = True

'            ' Verificar cada criterio
'            For j = 0 To criterios.Length - 1
'                If Not criterios(j)(i, 1) Then
'                    coincide = False
'                    Exit For
'                End If
'            Next

'            ' Si la fila cumple con todos los criterios, agregarla a la lista
'            If coincide Then
'                Dim fila(numCols - 1) As Object
'                For k As Integer = 1 To numCols
'                    fila(k - 1) = matrizValores(i, k)
'                Next
'                resultadosTemp.Add(fila)
'            End If
'        Next

'        ' Si se encontraron resultados, transferirlos a un array
'        If resultadosTemp.Count > 0 Then
'            Dim resultado(resultadosTemp.Count - 1, numCols - 1) As Object

'            For i = 0 To resultadosTemp.Count - 1
'                For j = 0 To numCols - 1
'                    resultado(i, j) = resultadosTemp(i)(j)
'                Next
'            Next

'            Return resultado
'        Else
'            ' Si no hay filas que cumplan con el criterio, devolver un mensaje
'            Return New Object(,) {{"No se encontraron coincidencias"}}
'        End If
'    End Function

'    ' Método para apilar matrices verticalmente
'    Public Function APILARV_VBA(ParamArray matrices() As Object(,)) As Object(,)
'        Dim totalFilas As Integer = 0
'        Dim maxColumnas As Integer = 0

'        ' Calcular el número total de filas y el máximo de columnas
'        For Each matriz In matrices
'            Dim filas As Integer = matriz.GetLength(0)
'            Dim columnas As Integer = matriz.GetLength(1)
'            totalFilas += filas
'            If columnas > maxColumnas Then
'                maxColumnas = columnas
'            End If
'        Next

'        ' Redimensionar el array de resultado
'        Dim resultado(totalFilas - 1, maxColumnas - 1) As Object
'        Dim currentFila As Integer = 0

'        ' Rellenar el array de resultado con los valores de las matrices
'        For Each matriz In matrices
'            Dim filas As Integer = matriz.GetLength(0)
'            Dim columnas As Integer = matriz.GetLength(1)

'            For i As Integer = 0 To filas - 1
'                For j As Integer = 0 To columnas - 1
'                    resultado(currentFila, j) = matriz(i, j)
'                Next
'                ' Rellenar con #N/A si hay menos columnas en esta matriz
'                For j As Integer = columnas To maxColumnas - 1
'                    resultado(currentFila, j) = ExcelError.ExcelErrorNA
'                Next
'                currentFila += 1
'            Next
'        Next

'        ' Devolver el resultado
'        Return resultado
'    End Function

'End Class
