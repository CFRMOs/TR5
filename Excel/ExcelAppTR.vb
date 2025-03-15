Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class ExcelAppTR
    Public xlSh As Worksheet = Nothing
    Public xlTable As ListObject = Nothing
    Public xlWkbook As Workbook
    Public ExcelApp As Application
    ' Método para seleccionar un rango
    Public Function InputBox(prompt As String, Type As Integer) As Object
        Return xlWkbook.Application.InputBox(prompt, Type:=Type)
    End Function

    ' Método para configurar la interactividad de Excel
    Public Sub SetInteractive(state As Boolean)
        xlWkbook.Application.Interactive = state
        xlWkbook.Application.ScreenUpdating = state
        xlWkbook.Application.DisplayAlerts = state
        xlWkbook.Application.EnableEvents = state
    End Sub

    ' Función para obtener la hoja de cálculo
    Public Function GetSH(Optional ShName As String = vbNullString) As Worksheet
        If ShName = vbNullString Then ShName = "REPORTE AUTOCAD"
        Try
            GetWorkB()
            If xlWkbook IsNot Nothing Then
                If WorksheetExists(xlWkbook, ShName) Then
                    xlSh = xlWkbook.Sheets(ShName)
                Else
                    xlSh = xlWkbook.Sheets.Add()
                    xlSh.Name = ShName
                End If
            End If
            Return xlSh
        Catch ex As Exception
            Console.WriteLine("Error al obtener la hoja: " & ex.Message)
            Return Nothing
        End Try
    End Function

    ' Función para obtener el libro de trabajo
    Public Function GetWorkB(Optional ByVal fileAddress As String = "") As Workbook
        Dim Xlwb As Workbook = Nothing

        Try
            If fileAddress <> "" AndAlso XlFileExists(fileAddress) Then
                Xlwb = XlFileOpen(fileAddress)
            Else
                Dim ExcelInstances As Process() = Process.GetProcessesByName("EXCEL")
                Dim xlApp As Application

                If ExcelInstances.Length = 0 Then
                    xlApp = New Application With {.Visible = True}
                    xlApp.Workbooks.Add()
                    Xlwb = xlApp.Workbooks(1)
                Else
                    Dim ExcelInstance As Application = TryCast(Marshal.GetActiveObject("Excel.Application"), Application)
                    If ExcelInstance IsNot Nothing AndAlso ExcelInstance.Workbooks.Count > 0 Then
                        Xlwb = ExcelInstance.Workbooks(1)
                    End If
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try

        xlWkbook = Xlwb
        Return Xlwb
    End Function

    ' Función para verificar si una hoja de cálculo existe
    Private Function WorksheetExists(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
        For Each ws As Worksheet In workbook.Sheets
            If ws.Name = sheetName Then Return True
        Next
        Return False
    End Function

    ' Función para verificar si una tabla existe
    Public Function CheckTable(ByRef table As ListObject, tableName As String) As Boolean
        If xlSh IsNot Nothing Then
            For Each tbl As ListObject In xlSh.ListObjects
                If tbl.Name = tableName Then
                    table = tbl
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    ' Función para crear una tabla
    Public Function CreateTable(ByVal worksheet As Worksheet, ByVal headers As List(Of String), ByVal tableName As String) As ListObject
        Try
            Dim startCell As Range = worksheet.Cells(1, 1)
            Dim endCell As Range = worksheet.Cells(1, headers.Count)
            Dim headerRange As Range = worksheet.Range(startCell, endCell)

            If CheckTable(xlTable, tableName) Then
                Return xlTable
            End If

            ' Establecer los encabezados en la hoja de Excel
            For i As Integer = 0 To headers.Count - 1
                worksheet.Cells(1, i + 1).Value = headers(i)
            Next

            ' Crear la tabla
            Dim tableRange As Range = worksheet.Range(startCell, worksheet.Cells(1, headers.Count))
            Dim vlistObject As ListObject = worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, tableRange, , XlYesNoGuess.xlYes)
            vlistObject.Name = tableName
            xlTable = vlistObject
            Return vlistObject
        Catch ex As Exception
            Console.WriteLine("Error al crear la tabla: " & ex.Message)
            Return Nothing
        End Try
    End Function

    ' Función para extraer datos de la tabla ya creada
    Public Function GetTransferData(ByVal data As Array, ByVal headers As Array, ByVal handle As String, ByVal plano As String, ByVal fileName As String) As Dictionary(Of String, String)
        Dim rowData As New Dictionary(Of String, String)()

        ' Identificar índices de los headers necesarios
        Dim handleIndex As Integer = Array.IndexOf(headers, "TabName")
        Dim planoIndex As Integer = Array.IndexOf(headers, "Plano")
        Dim fileNameIndex As Integer = Array.IndexOf(headers, "FileName")
        Dim comentariosIndex As Integer = Array.IndexOf(headers, "Comentarios")
        Dim tramosIndex As Integer = Array.IndexOf(headers, "Tramos")

        ' Iterar sobre cada fila de datos
        For Each row As Object In data
            ' Comparar valores de handle, plano y fileName
            If row(handleIndex).ToString() = handle AndAlso row(planoIndex).ToString() = plano AndAlso row(fileNameIndex).ToString() = fileName Then
                rowData("Comentarios") = row(comentariosIndex).ToString()
                rowData("Tramos") = row(tramosIndex).ToString()
                Exit For
            End If
        Next

        Return rowData
    End Function

    ' Establecer datos de transferencia
    Public Sub SetTransferData(ByVal data As Array, ByVal headers As Array, ByVal formats As Array, Optional ByVal xlTable As ListObject = Nothing, Optional ByVal comments As String = "")
        If data Is Nothing Then Throw New ArgumentNullException(NameOf(data))
        If headers Is Nothing Then Throw New ArgumentNullException(NameOf(headers))
        If formats Is Nothing Then Throw New ArgumentNullException(NameOf(formats))
        If xlSh Is Nothing OrElse xlTable Is Nothing Then Exit Sub
        Try
            Dim handleColIndex As Integer = Array.IndexOf(headers, "TabName") + 1
            If handleColIndex > 0 Then
                For Each row As ListRow In xlTable.ListRows
                    Dim cellValue As Object = row.Range.Cells(1, handleColIndex).Value
                    If cellValue IsNot Nothing AndAlso cellValue.ToString() = data(handleColIndex - 1).ToString() Then
                        UpdateRow(row, data, formats)
                        If Not String.IsNullOrEmpty(comments) Then
                            Dim commentsIndex = Array.IndexOf(headers, "Comentarios") + 1
                            If commentsIndex > 0 Then
                                row.Range.Cells(1, commentsIndex).Value = comments
                            End If
                        End If
                        Return
                    End If
                Next
            End If
            Dim newRow As ListRow = xlTable.ListRows.Add()
            UpdateRow(newRow, data, formats)
            If Not String.IsNullOrEmpty(comments) Then
                Dim commentsIndex = Array.IndexOf(headers, "Comentarios") + 1
                If commentsIndex > 0 Then
                    newRow.Range.Cells(1, commentsIndex).Value = comments
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Error al añadir datos a la tabla: " & ex.Message)
        End Try
    End Sub

    ' Método para actualizar la fila
    Private Sub UpdateRow(row As ListRow, data As Array, formats As Array)
        Try
            For i As Integer = 0 To data.Length - 1
                row.Range.Cells(1, i + 1).NumberFormat = formats(i)
                row.Range.Cells(1, i + 1).Value = data(i)
            Next
        Catch ex As InvalidCastException
            Console.WriteLine("Error al actualizar la fila: " & ex.Message)
        End Try
    End Sub

    ' Nueva función FTtransferdata para transferir datos
    Public Sub FTtransferdata(ByVal AR As Array, ByVal ARTitulos As Array, ByVal ARFT As Array)
        SetTransferData(AR, ARTitulos, ARFT)
    End Sub

    ' Función para guardar el archivo
    Public Sub SaveAs(ByVal Name As String, ByVal FileAddress As String)
        Try
            If xlWkbook IsNot Nothing AndAlso Not XlFileExists(Path.Combine(FileAddress, Name)) Then
                xlWkbook.SaveAs(Filename:=Path.Combine(FileAddress, Name))
            End If
        Catch ex As Exception
            Console.WriteLine("Error al guardar el archivo: " & ex.Message)
        End Try
    End Sub

    ' Función para verificar si el archivo existe
    Public Function XlFileExists(ByVal FileAddress As String) As Boolean
        Return File.Exists(FileAddress)
    End Function

    ' Función para abrir un archivo existente
    Public Function XlFileOpen(ByVal FileAddress As String) As Workbook
        Dim Xlwb As Workbook = Nothing
        Try
            ' Verifica si el archivo existe
            If Not File.Exists(FileAddress) Then
                Throw New FileNotFoundException("El archivo no fue encontrado.", FileAddress)
            End If

            Dim xlApp As Application
            ' Intenta obtener una instancia de Excel ya abierta
            Try
                xlApp = TryCast(Marshal.GetActiveObject("Excel.Application"), Application)
            Catch ex As COMException
                ' Si no hay ninguna instancia de Excel, crea una nueva
                xlApp = New Application With {.Visible = True}
            End Try

            ' Verifica si el archivo ya está abierto
            For Each wb As Workbook In xlApp.Workbooks
                If wb.FullName = FileAddress Then
                    Xlwb = wb
                    Console.WriteLine("El archivo ya está abierto.")
                    Exit For
                End If
            Next

            ' Si el archivo no está abierto, abrirlo
            If Xlwb Is Nothing Then
                ' Desactiva las alertas
                xlApp.DisplayAlerts = False

                ' Abre el archivo sin actualizar enlaces y sin notificaciones
                Xlwb = xlApp.Workbooks.Open(Filename:=FileAddress, UpdateLinks:=False, Notify:=False, AddToMru:=True)

                ' Espera hasta que el archivo esté completamente cargado
                Do While xlApp.Ready = False
                    Threading.Thread.Sleep(100) ' Espera 100 ms y luego verifica de nuevo
                Loop

                Console.WriteLine("El archivo se abrió correctamente y está listo para usarse.")
            End If

        Catch ex As FileNotFoundException
            Console.WriteLine("Archivo no encontrado: " & ex.Message)
        Catch ex As COMException
            Console.WriteLine("Error COM al abrir Excel: " & ex.Message)
        Catch ex As Exception
            Console.WriteLine("Error general al abrir el archivo: " & ex.Message)
        End Try

        Return Xlwb
    End Function

    ' Obtener comentarios de la tabla
    Public Function GetCommentsFromTable(ByVal worksheet As Worksheet, ByVal tableName As String, ByVal handle As String) As String
        If worksheet Is Nothing Then
            Throw New ArgumentNullException(NameOf(worksheet))
        End If

        If CheckTable(xlTable, tableName) Then
            Dim headers = xlTable.HeaderRowRange.Value.Cast(Of Object)().Select(Function(header) header.ToString()).ToArray()
            Dim data = xlTable.DataBodyRange.Value
            Dim commentsData = GetTransferData(data, headers, handle, "", "")
            If commentsData.ContainsKey("Comentarios") Then
                Return commentsData("Comentarios")
            End If
        End If
        Return String.Empty
    End Function

    ' Obtener o crear tabla
    Public Function GetOrCreateTable(ByVal worksheet As Worksheet, ByVal headers As List(Of String), ByVal tableName As String) As ListObject
        If CheckTable(xlTable, tableName) Then
            Return xlTable
        Else
            Return CreateTable(worksheet, headers, tableName)
        End If
    End Function

    ' Liberar recursos COM
    Protected Overrides Sub Finalize()
        If xlWkbook IsNot Nothing Then Marshal.ReleaseComObject(xlWkbook)
        If xlSh IsNot Nothing Then Marshal.ReleaseComObject(xlSh)
        If xlTable IsNot Nothing Then Marshal.ReleaseComObject(xlTable)
        MyBase.Finalize()
    End Sub
End Class
