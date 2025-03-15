Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class UpdateExcelTable
    Public Shared Sub CheckExcelRange(ByRef CAcadHelp As ACAdHelpers, Headers As List(Of String), DGView As DataGridView, Optional ByRef handleProcessor As HandleDataProcessor = Nothing)
        SelectRowOnTbl(CAcadHelp, "CunetasGeneral", Headers, DGView, handleProcessor)
    End Sub

    Public Shared Sub SelectRowOnTbl(ByRef CAcadHelp As ACAdHelpers, TblName As String, Headers As List(Of String), DGView As DataGridView, Optional ByRef handleProcessor As HandleDataProcessor = Nothing)
        Dim ExRNGSe As New ExcelRangeSelector()
        Dim optimizer As ExcelOptimizer = Nothing ' Instancia de ExcelOptimizer

        Try
            ' Inicializamos el optimizador con la aplicación Excel activa
            Dim workbook = ExRNGSe.ExcelApp.GetWorkB()
            If workbook Is Nothing Then
                Debug.WriteLine("No se pudo obtener el libro de trabajo.")
                Exit Sub
            End If

            optimizer = New ExcelOptimizer(workbook.Application)
            ' Desactivamos cálculos automáticos, eventos y actualizaciones
            optimizer.TurnEverythingOff()

            ' Obtener la Tabla de Excel proporcionada por el nombre
            Dim tbl As ListObject = ExRNGSe.GetTableOnWorkBkByName(TblName)

            ' Verificar si se encontró la tabla
            If tbl Is Nothing Then
                Debug.WriteLine("La tabla no fue encontrada.")
                Exit Sub
            End If
            ' Si no se encuentra el código, salir
            Dim RefAcceso As String = AutoCADHelper2.AccesoNum(CAcadHelp.ThisDrawing.Name)

            If String.IsNullOrEmpty(RefAcceso) Then
                Console.WriteLine("El archivo no presenta referencia de accesos.")
                Exit Sub
            End If

            ' Obtener el número de acceso a partir de la referencia extraída
            Dim ACCESO As Double = CDbl(CodigoExtractor.ExtraerCodigo(RefAcceso, "\d+"))

            ' Verificar si se ha encontrado un código de acceso válido
            If ACCESO = 0 Then
                Throw New Exception("No se pudo extraer el código de acceso del nombre del archivo.")
            End If

            ' Inicializar la lista de rangos donde se hará la búsqueda
            Dim Ranges As New List(Of Range)

            ' Almacenar el rango de los encabezados para evitar múltiples accesos a objetos 
            Dim headerRange As Range = tbl.HeaderRowRange

            ' Almacenar las columnas que necesitas y sus índices
            Dim columnas() As String = {"StartStation", "EndStation", "Longitud"}
            Dim indexcolumnas As Integer() = ConArray(columnas, headerRange)

            '' Aquí calculamos los índices de las columnas para `Headers` una sola vez
            Dim indexHeaders As Integer() = ConArray(Headers, headerRange)

            '' Aquí calculamos los índices de las columnas para `construir Layer Name` una sola vez
            Dim indexHeadersCLayer As Integer() = ConArray({"Tipo", "Medicion"}, headerRange)

            ' Obtener la fila seleccionada en el DataGridView

            ' Hacer una copia de las celdas seleccionadas para no modificar el listado original
            Dim selectedCellsList As New List(Of DataGridViewCell)()

            For Each cell As DataGridViewCell In DGView.SelectedCells
                selectedCellsList.Add(cell)
            Next

            For Each Cell As DataGridViewCell In selectedCellsList

                Dim Row As DataGridViewRow = DGView.Rows(Cell.RowIndex) 'DGView.Rows(DGView.SelectedCells(0).RowIndex)
                ' Iterar a través de las filas de la tabla
                For Each TblRow As ListRow In tbl.ListRows
                    ' Obtener el rango de la fila actual
                    Dim RNG As Range = TblRow.Range

                    ' Verificar si el valor de la primera celda de la fila coincide con el código de acceso
                    If RNG.Cells(1).Value = ACCESO Then
                        Dim CIFRAS As Integer = 2
                        Dim Tolerancias As Double = 2
                        Dim ACTUALIZAR As Boolean = True

                        ' Iterar sobre los encabezados que se deben buscar
                        For i As Integer = LBound(indexcolumnas) To UBound(indexcolumnas)

                            ' Obtener el índice de la columna correspondiente
                            Dim columnIndex As Integer = indexcolumnas(i)

                            ' Verificar si los valores de la fila y la tabla coinciden, redondeando los valores
                            ' Considerando una tolerancia de dos
                            If Not Math.Abs(Math.Round(Row.Cells(columnas(i)).Value, CIFRAS) - Math.Round(RNG.Cells(columnIndex).Value, CIFRAS)) <= 2 Then

                                ' Si los valores no coinciden, no se actualiza la tabla
                                ACTUALIZAR = False
                                Exit For
                            End If
                        Next

                        ' Si se debe actualizar, proceder con la actualización de valores desde el DataGridView
                        If ACTUALIZAR Then
                            'iteracion en los encabezados
                            Dim j As Integer = 1
                            For i As Integer = 0 To Headers.Count - 2
                                Dim columnIndex As Integer = indexHeaders(i) ' Usamos el índice prealmacenado para Headers
                                If Headers(i) <> "Comentarios" Then
                                    If Headers(i) <> "Layer" Then
                                        RNG.Cells(j, columnIndex).Value = Row.Cells(Headers(i)).Value
                                    Else
                                        'estoy contruyendo el Layer Name
                                        RNG.Cells(j, columnIndex).Value = "MED-" & Format(RNG.Cells(j, indexHeadersCLayer(1)).Value, "00") & "-" & RNG.Cells(j, indexHeadersCLayer(0)).Value

                                        If handleProcessor.DicHandlesByLength.ContainsKey(RNG.Cells(j, columnIndex - 1).Value) Then
                                            ' set layer to Entity
                                            CLayerHelpers.ChangeLayersAcEnt(RNG.Cells(j, columnIndex - 1).Value, RNG.Cells(j, columnIndex).Value)
                                            Dim Cuneta As CunetasHandleDataItem = handleProcessor.DicHandlesByLength(RNG.Cells(j, columnIndex - 1).Value)

                                            If handleProcessor.DicHandlesExistentes.ContainsKey(Cuneta.Handle) = False Then
                                                handleProcessor.DicHandlesExistentes.Add(Cuneta.Handle, Cuneta)
                                                'remover del listado de "Por Longitud"
                                            End If
                                            If handleProcessor.DicHandlesExistentes.ContainsKey(Cuneta.Handle) Then
                                                handleProcessor.DicHandlesByLength.Remove(Cuneta.Handle)
                                            End If
                                        Else
                                            Dim Cuneta As New CunetasHandleDataItem(listObject:=tbl)

                                            Cuneta.SetPropertiesFromDWG(Row.Cells("Handle").Value, Row.Cells("AlignmentHDI").Value)

                                            handleProcessor.DicHandlesExistentes.Add(Cuneta.Handle, Cuneta)

                                            Row.Cells("AlignmentHDI").Value = Cuneta.AlignmentHDI
                                        End If

                                    End If

                                End If
                            Next
                            Exit For
                        End If
                    End If
                Next

            Next


            ' Liberar el objeto COM del rango de encabezados para optimizar memoria
            Marshal.ReleaseComObject(headerRange)
            Marshal.ReleaseComObject(tbl.HeaderRowRange)
            Marshal.ReleaseComObject(tbl)

        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        Finally
            ' Restauramos cálculos automáticos, eventos y actualizaciones al final del proceso
            optimizer?.TurnEverythingOn()
            optimizer?.Dispose()
        End Try
    End Sub

    Public Shared Function ConArray(Arr As IEnumerable(Of String), headerRange As Excel.Range) As Integer()

        ' Crear un diccionario para almacenar las posiciones de los encabezados ya encontrados
        Dim headerDict As New Dictionary(Of String, Integer)
        Dim indexList As New List(Of Integer)

        For Each iHeader As String In Arr
            If headerDict.ContainsKey(iHeader) Then
                ' Si ya hemos encontrado este encabezado antes, usamos el índice almacenado
                indexList.Add(headerDict(iHeader))
            Else
                ' Si no, lo buscamos en el rango de encabezados
                Dim columnIndex As Excel.Range = Nothing
                For Each RNG As Range In headerRange
                    If RNG.Value2 = iHeader Then
                        columnIndex = RNG ' headerRange.Find(iHeader)
                        Exit For
                    End If
                Next
                If columnIndex IsNot Nothing Then
                    ' Guardar el índice en el diccionario y en la lista de resultados
                    Dim colIndex As Integer = columnIndex.Column
                    headerDict.Add(iHeader, colIndex)
                    indexList.Add(colIndex)
                Else
                    'Throw New Exception("La columna " & iHeader & " no fue encontrada en los encabezados.")
                End If
            End If
        Next

        Return indexList.ToArray()
    End Function
End Class
'Public Shared Sub uptedate(indexcolumnas As Integer, tbl As Range, ACCESO As Double, Optional ByRef handleProcessor As HandleDataProcessor = Nothing)
'    ' Iterar a través de las filas de la tabla
'    For Each TblRow As ListRow In tbl.ListRows
'        ' Obtener el rango de la fila actual
'        Dim RNG As Range = TblRow.Range

'        ' Verificar si el valor de la primera celda de la fila coincide con el código de acceso
'        If RNG.Cells(1).Value = ACCESO Then
'            Dim CIFRAS As Integer = 2
'            Dim Tolerancias As Double = 2
'            Dim ACTUALIZAR As Boolean = True

'            ' Iterar sobre los encabezados que se deben buscar
'            For i As Integer = LBound(indexcolumnas) To UBound(indexcolumnas)

'                ' Obtener el índice de la columna correspondiente
'                Dim columnIndex As Integer = indexcolumnas(i)

'                ' Verificar si los valores de la fila y la tabla coinciden, redondeando los valores
'                ' Considerando una tolerancia de dos
'                If Not Math.Abs(Math.Round(Row.Cells(columnas(i)).Value, CIFRAS) - Math.Round(RNG.Cells(columnIndex).Value, CIFRAS)) <= 2 Then

'                    ' Si los valores no coinciden, no se actualiza la tabla
'                    ACTUALIZAR = False
'                    Exit For
'                End If
'            Next

'            ' Si se debe actualizar, proceder con la actualización de valores desde el DataGridView
'            If ACTUALIZAR Then
'                'iteracion en los encabezados
'                Dim j As Integer = 1
'                For i As Integer = 0 To Headers.Count - 1
'                    Dim columnIndex As Integer = indexHeaders(i) ' Usamos el índice prealmacenado para Headers
'                    If Headers(i) <> "Comentarios" Then
'                        If Headers(i) <> "Layer" Then
'                            RNG.Cells(j, columnIndex).Value = Row.Cells(Headers(i)).Value
'                        Else
'                            'estoy contruyendo el Layer Name
'                            RNG.Cells(j, columnIndex).Value = "MED-" & Format(RNG.Cells(j, indexHeadersCLayer(1)).Value, "00") & "-" & RNG.Cells(j, indexHeadersCLayer(0)).Value

'                            ' set layer to Entity
'                            CLayerHelpers.ChangeLayersAcEnt(RNG.Cells(j, columnIndex - 1).Value, RNG.Cells(j, columnIndex).Value)
'                            Dim Cuneta As CunetasHandleDataItem = handleProcessor.DicHandlesByLength(RNG.Cells(j, columnIndex - 1).Value)
'                            handleProcessor.DicHandlesExistentes.Add(Cuneta.TabName, Cuneta)
'                            handleProcessor.DicHandlesByLength.Remove(Cuneta.TabName)
'                            'remover del listado de "Por Longitud"
'                        End If

'                    End If
'                Next
'                Exit For
'            End If
'        End If
'    Next
'End Sub