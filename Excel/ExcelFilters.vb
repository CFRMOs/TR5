Imports Microsoft.Office.Interop.Excel
Imports System.Collections.Generic

Class ExcelFilters
    Private FiltrosGuardados As New Dictionary(Of Integer, Tuple(Of String, String, XlAutoFilterOperator))
    ' Función para guardar los filtros aplicados en una tabla
    Public Sub GuardarFiltros(tbl As ListObject)

        ' Limpiar el diccionario antes de guardar
        FiltrosGuardados.Clear()

        ' Verificar si hay filtros aplicados
        If tbl.AutoFilter IsNot Nothing Then
            Dim filtros As AutoFilter = tbl.AutoFilter
            For i As Integer = 1 To tbl.ListColumns.Count
                Dim filtro As Filter = filtros.Filters(i)
                If filtro.On Then
                    Dim criterio1 As String = filtro.Criteria1
                    Dim criterio2 As String = If(filtro.Operator = XlAutoFilterOperator.xlAnd, filtro.Criteria2, "")
                    Dim operador As XlAutoFilterOperator = filtro.Operator

                    ' Guardar en el diccionario
                    FiltrosGuardados(i) = Tuple.Create(criterio1, criterio2, operador)
                End If
            Next
            If tbl.AutoFilter.FilterMode Then
                tbl.AutoFilter.ShowAllData()
            End If
        End If
    End Sub

    ' Función para aplicar los filtros guardados en una tabla
    Public Sub AplicarFiltros(tbl As ListObject)
        If FiltrosGuardados.Count = 0 Then Exit Sub ' No hay filtros guardados

        ' Restaurar los filtros guardados
        For Each key As Integer In FiltrosGuardados.Keys
            Dim filtro As Tuple(Of String, String, XlAutoFilterOperator) = FiltrosGuardados(key)
            If Not String.IsNullOrEmpty(filtro.Item2) Then
                tbl.Range.AutoFilter(Field:=key, Criteria1:=filtro.Item1, Operator:=filtro.Item3, Criteria2:=filtro.Item2)
            Else
                tbl.Range.AutoFilter(Field:=key, Criteria1:=filtro.Item1)
            End If
        Next
    End Sub
End Class
