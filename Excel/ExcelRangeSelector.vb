Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class ExcelRangeSelector
    Public ExcelApp As ExcelAppTR
    Private selectedRange As Range = Nothing

    ' Constructor
    Public Sub New(Optional ByRef cExcelApp As ExcelAppTR = Nothing)
        ExcelApp = New ExcelAppTR()
        cExcelApp = ExcelApp
    End Sub
    Public Function SelectRowOnTbl(ByVal tableName As String, ByVal handle As String) As Range
        ' Obtener la tabla con el nombre proporcionado
        Dim tbl As ListObject = GetTableOnWorkBkByName(tableName)

        ' Verificar si se encontró la tabla
        If tbl Is Nothing Then
            'MsgBox "Tabla no encontrada: " & tableName
            Return Nothing
        End If

        ' Validar que se ha proporcionado un valor de Handle para buscar
        If handle = "" Then
            'MsgBox "No se proporcionó ningún valor de Handle"
            Return Nothing
        End If

        ' Encontrar el índice de la columna "Handle"
        Dim handleColumnIndex As Integer
        handleColumnIndex = 0

        For Each header As Range In tbl.HeaderRowRange
            If header.Value = "Handle" Then
                handleColumnIndex = header.Column - tbl.HeaderRowRange.Cells(1, 1).Column + 1
                Exit For
            End If
        Next header

        ' Verificar si se encontró la columna "Handle"
        If handleColumnIndex = 0 Then
            'MsgBox "No se encontró la columna 'Handle' en la tabla: " & tableName
            Return Nothing
        End If

        ' Buscar el valor del handle dentro de la columna "Handle" de la tabla
        Dim rngFound As Range = tbl.DataBodyRange.Columns(handleColumnIndex).Find(handle, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole)

        ' Verificar si se encontró el valor del Handle
        If rngFound Is Nothing Then
            'MsgBox "No se encontró el valor de Handle: " & handle
            Return Nothing
        End If

        ' Obtener el número de fila relativo en la tabla
        Dim rowNumber As Integer = rngFound.Row - tbl.DataBodyRange.Row + 1

        ' Devolver el rango de la fila seleccionada
        Return tbl.ListRows(rowNumber).Range
    End Function

    Public Function GetTableOnWorkBkByName(TblName As String, Optional WBFilePath As String = vbNullString) As ListObject
        Dim cExcelApp As New ExcelAppTR
        ' Obtener el libro activo
        Dim workBk As Workbook
        'chequeo de existencia de carpeta

        If FUrlCarpetaProcessor.ArchivoExiste(WBFilePath) Then
            workBk = ExcelApp.GetWorkB(WBFilePath)
        Else
            workBk = ExcelApp.GetWorkB("")
        End If

        Dim tbl As ListObject = Nothing

        ' Buscar la tabla con el nombre proporcionado
        For Each sh As Worksheet In workBk.Sheets
            tbl = GetTableOnSheetByName(TblName, sh)
            If tbl IsNot Nothing Then
                Return tbl
            End If
        Next

        ' Verificar si se encontró la tabla
        If tbl Is Nothing Then
            Return Nothing
        Else
            Return tbl
        End If
    End Function

    Public Function GetTableOnSheetByName(TblName As String, worksheet As Worksheet) As ListObject
        Dim tbl As ListObject = Nothing
        ' Buscar la tabla con el nombre proporcionado
        For Each t As ListObject In worksheet.ListObjects
            If t.Name = TblName Then
                tbl = t
                Return tbl
                'Exit For
            End If
        Next
        ' Verificar si se encontró la tabla

        If tbl Is Nothing Then
            Return Nothing
        Else
            Return tbl
        End If
    End Function


    Public Function SelectRowOnTbl(TblName As String, Headers As List(Of String), Optional ByRef Handle As String = "", Optional RNG As Range = Nothing) As List(Of Object)
        'Dim RNG As Range = Nothing
        ' Obtener la Tabla de proporcionado
        Dim tbl As ListObject = GetTableOnWorkBkByName(TblName)

        ' Verificar si se encontró la tabla
        If tbl Is Nothing Then
            Return Nothing
        End If

        If Handle = "" Then
            RNG = SelectRange()
        ElseIf Handle <> "" Then
            RNG = tbl.Range.Worksheet.Range(TblName & "[Handle]").Find(Handle)
        End If

        ' Validar si el rango es válido y si pertenece a la tabla especificada
        If RNG Is Nothing OrElse Not IsInTbl(TblName, RNG) Then
            Return Nothing
        End If

        ' Obtener los valores de la fila seleccionada en las columnas "TabName" y "FilePath"
        Dim rowNumber As Integer = RNG.Row - tbl.HeaderRowRange.Row

        Dim Datos As List(Of Object) = ConstructDatos(rowNumber, Headers, tbl)

        Return Datos
    End Function
    Public Function ConstructDatos(rowNumber As Integer, Headers As List(Of String), tbl As ListObject) As List(Of Object)
        Dim Datos As New List(Of Object)

        For Each col As ListColumn In tbl.ListColumns
            If Headers.Contains(col.Name) Then
                ' Obtener las columnas "TabName" y "FilePath" de la tabla
                Dim Column As ListColumn = col
                ' Verificar si ambas columnas se encontraron
                If Column Is Nothing Then
                    Return Nothing
                End If
                If rowNumber > 0 Then
                    Dim value As Object = tbl.ListRows(rowNumber).Range.Cells(1, Column.Index).Value
                    ' Añadir los valores a la lista de resultados los cuales pueden ser string o double  
                    Datos.Add(value)
                End If
            End If
        Next
        Return Datos
    End Function

    Public Function IsInTbl(TblName As String, RNG As Range) As Boolean
        ' Inicializamos el valor de retorno como False
        IsInTbl = False

        ' Verifica que el rango y el nombre de la tabla no sean nulos o vacíos
        If RNG Is Nothing OrElse String.IsNullOrEmpty(TblName) Then
            Return False
        End If

        ' Obtener la hoja de cálculo del rango proporcionado
        Dim worksheet As Worksheet = RNG.Worksheet

        ' Recorrer todas las tablas en la hoja de cálculo
        For Each tbl As ListObject In worksheet.ListObjects
            ' Verificar si el nombre de la tabla coincide con el nombre proporcionado
            If tbl.Name = TblName Then
                ' Verificar si el rango está dentro del rango de la tabla
                If ExcelApp.GetWorkB().Application.Intersect(tbl.Range, RNG) IsNot Nothing Then
                    IsInTbl = True
                    Exit For
                End If
            End If
        Next

        Return IsInTbl
    End Function
    Public Function SelectRange() As Range
        Try
            ' Intenta obtener la aplicación de Excel en ejecución
            ExcelApp.GetWorkB()

            If ExcelApp.xlWkbook Is Nothing Then
                MsgBox("No hay ningún libro activo en Excel.", MsgBoxStyle.Critical)
                Return Nothing
            End If

            Console.WriteLine("Libro activo encontrado: " & ExcelApp.xlWkbook.Name)

            'ExcelApp.SetInteractive(True)
            Console.WriteLine("Por favor, selecciona un rango de celdas en Excel y luego presiona OK.")
            selectedRange = CType(ExcelApp.InputBox("Selecciona un rango de celdas", Type:=8), Range)
            'ExcelApp.SetInteractive(False)
            Return selectedRange
        Catch ex As Exception
            MsgBox("Error al seleccionar el rango: " & ex.Message, MsgBoxStyle.Critical)
            Return Nothing
        Finally
            ' Liberar objetos COM
        End Try
    End Function

    Public Function GetdoubleData() As Double
        Dim Rng As Range = SelectRange()
        If Rng Is Nothing Then Return Nothing
        Dim Station As Double = Rng.Value

        ' Liberar objetos COM
        Return Station
    End Function

    Public Function HandleFromExcelTable(TBName As String, handle As String, Header As String) As Range
        Dim Table As ListObject = GetTableOnWorkBkByName(TBName)
        Dim RNG As Range = Nothing

        If Table IsNot Nothing Then
            ' Apagar optimizaciones de Excel para mejorar rendimiento
            'Dim excelOptimizer As New ExcelOptimizer()
            'excelOptimizer.TurnEverythingOff()
            Dim ExcelFilters As New ExcelFilters

            Try
                Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"Handle", Header}, Table.HeaderRowRange)
                Dim HandleColumnIndex As Integer = indexHeader(0)
                Dim ColumnIndex As Integer = indexHeader(1)

                Dim handleColumnRange As Range = Table.ListColumns(HandleColumnIndex).DataBodyRange

                ' Verificar si hay filtros aplicados antes de intentar eliminarlos

                ExcelFilters.GuardarFiltros(Table)

                If Table.AutoFilter.FilterMode Then
                    Table.AutoFilter.ShowAllData()
                End If

                Dim foundCell As Range = LookInRange(Table, handle, Header) ', HandleColumnIndex) 'handleColumnRange.Find(What:=handle, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole)



                ' Si `Find` encuentra un resultado, verificar si está visible
                If foundCell IsNot Nothing AndAlso foundCell.EntireRow.Hidden = False Then
                    Dim rowNumber As Integer = foundCell.Row - Table.DataBodyRange.Row + 1
                    RNG = Table.ListRows(rowNumber).Range.Cells(1, ColumnIndex)
                End If
            Catch ex As Exception
                Debug.WriteLine("Error en HandleFromExcelTable: " & ex.Message)
            Finally
                ExcelFilters.AplicarFiltros(Table)
                ' Restaurar optimizaciones
                'excelOptimizer.TurnEverythingOn()
                'excelOptimizer.Dispose()
            End Try
        End If
        Return RNG
    End Function
    Public Function LookInRange(Table As Microsoft.Office.Interop.Excel.ListObject, handle As String, Header As String) As Range ', ByRef HandleColumnIndex As Integer) As Range
        Try
            Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"Handle", Header}, Table.HeaderRowRange)

            Dim ColumnIndex As Integer = indexHeader(1)

            Dim HandleColumnIndex As Integer = indexHeader(0)

            Dim handleColumnRange As Range = Table.ListColumns(HandleColumnIndex).DataBodyRange

            HandleColumnIndex = indexHeader(0)
            ' Verificar si hay filtros aplicados antes de intentar eliminarlos



            Return handleColumnRange.Find(What:=handle, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole)

        Catch ex As Exception
            Return Nothing
        Finally

        End Try

    End Function

    Private Shared Sub ReleaseComObject(ByVal obj As Object)
        If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
            Marshal.ReleaseComObject(obj)
        End If
    End Sub

    Friend Function GetColumnsRange() As IEnumerable(Of Range)
        Throw New NotImplementedException()
    End Function

    'parte del conogo no utilizados 
    '_________________________________________________________________________________________________________

    Public Function GetData() As String
        Dim Rng As Range = SelectRange()
        If Rng Is Nothing Then Return Nothing
        Dim Station As String = Rng.Value

        ' Liberar objetos COM
        Return Station
    End Function

    Public Function SelectRowOnTblByhandle(TblName As String, Headers As List(Of String), StrHD As String) As List(Of Object)

        ' Obtener la Tabla de proporcionado
        Dim tbl As ListObject = GetTableOnWorkBkByName(TblName)
        ' Verificar si se encontró la tabla
        If tbl Is Nothing Then
            Return Nothing
        End If
        ' Obtener los valores de la fila seleccionada en las columnas "Handle" y "FilePath"
        Dim rowNumber As Integer = Range(TblName & "[Handle]").Find(StrHD).Row - tbl.HeaderRowRange.Row

        Dim Datos As List(Of Object) = ConstructDatos(rowNumber, Headers, tbl)

        Return Datos
    End Function

    Private Function Range(v As String) As Range
        Throw New NotImplementedException()
    End Function

    Public Function GetColumnsRange(TblName As String, ColumnNames As List(Of String)) As List(Of Range)
        Dim ColumnRanges As New List(Of Range)
        'Dim cExcelApp As New ExcelAppTR 'clase ya existente y definida 

        ' Obtener el libro activo 
        'Dim workBk As Workbook = cExcelApp.GetWorkB()
        Dim tbl As ListObject = GetTableOnWorkBkByName(TblName)


        ' Verificar si se encontró la tabla
        If tbl Is Nothing Then
            Return Nothing
        End If

        ' Buscar las columnas con los nombres proporcionados
        For Each columnName As String In ColumnNames
            Dim targetColumn As ListColumn = Nothing
            For Each col As ListColumn In tbl.ListColumns
                If col.Name = columnName Then
                    targetColumn = col
                    Exit For
                End If
            Next

            ' Verificar si se encontró la columna
            If targetColumn IsNot Nothing Then
                ColumnRanges.Add(targetColumn.DataBodyRange)
            Else
                ' Manejar columnas no encontradas: opcionalmente puedes agregar lógica para manejar esto
                ' Por ejemplo, puedes continuar, agregar un rango nulo, o detener el procesamiento
            End If
        Next

        ' Devolver la lista de rangos de columnas
        Return ColumnRanges
    End Function

    Public Function HandleFromExcelTable2(TBName As String, handle As String, Header As String) As Range
        Dim Table As Microsoft.Office.Interop.Excel.ListObject = GetTableOnWorkBkByName(TBName)
        Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing

        If Table IsNot Nothing Then
            Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"Handle", Header}, Table.HeaderRowRange)
            Dim HandleColumnIndex As Integer = indexHeader(0)
            Dim ColumnIndex As Integer = indexHeader(1)

            For Each row As Microsoft.Office.Interop.Excel.ListRow In Table.ListRows
                Dim cellValue As Object = row.Range.Cells(1, HandleColumnIndex).Value

                ' Verificar si la celda tiene valor antes de llamar a ToString()
                If cellValue IsNot Nothing AndAlso Not String.IsNullOrEmpty(cellValue.ToString()) Then
                    If cellValue.ToString() = handle Then
                        RNG = row.Range ' Return the matching row if found
                        Exit For
                    End If
                End If
            Next

            If RNG IsNot Nothing Then RNG = RNG.Cells(1, ColumnIndex)
        End If

        Return RNG
    End Function


    ' Liberar recursos COM
    Protected Overrides Sub Finalize()
        If selectedRange IsNot Nothing Then ReleaseComObject(selectedRange)
        MyBase.Finalize()
    End Sub
End Class
