Imports Autodesk.AutoCAD.Runtime
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Clase para gestionar una tabla en Excel, permitiendo copiarla en paralelo,
''' manejar los rangos de títulos, datos y totales, y verificar si ha sido copiada.
''' </summary>
Public Class RevisionSoporteExcel
    Private excelApp As Application
    Private workbook As Workbook
    Private worksheet As Worksheet

    ' Propiedades para almacenar los rangos y títulos
    Public Property TituloNombres As List(Of String)
    Public Property RangoTitulos As Range
    Public Property RangoDatos As Range
    Public Property RangoTotales As Range

    ''' <summary>
    ''' Inicializa una nueva instancia de la clase RevisionSoporteExcel.
    ''' </summary>
    ''' <param name="workbookName">Nombre del libro de trabajo (opcional).</param>
    ''' <param name="workbookPath">Ruta del archivo del libro de trabajo (opcional).</param>
    ''' <param name="worksheetName">Nombre de la hoja de trabajo (opcional).</param>
    ''' <param name="celdaInicio">Celda inicial de la tabla (por defecto es "A1").</param>
    ''' <param name="titulos">Lista de títulos o rango de títulos (opcional).</param>
    Public Sub New(Optional ByVal workbookName As String = Nothing,
                   Optional ByVal workbookPath As String = Nothing,
                   Optional ByVal worksheetName As String = Nothing,
                   Optional ByVal celdaInicio As String = "A1",
                   Optional ByVal titulos As Object = Nothing)

        ' Obtener la aplicación de Excel activa
        excelApp = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Application)

        ' Seleccionar el libro de trabajo
        workbook = ObtenerWorkbook(workbookName, workbookPath)

        ' Seleccionar la hoja de trabajo
        worksheet = ObtenerWorksheet(workbook, worksheetName)

        ' Configurar los rangos de la tabla
        IdentificarRangos(celdaInicio)

        ' Configurar títulos según el tipo de entrada
        If TypeOf titulos Is List(Of String) Then
            TituloNombres = CType(titulos, List(Of String))
            RangoTitulos = GenerarRangoDeTitulos(celdaInicio, TituloNombres)
        ElseIf TypeOf titulos Is Range Then
            RangoTitulos = CType(titulos, Range)
            TituloNombres = ExtraerTitulosDelRango(RangoTitulos)
        Else
            TituloNombres = ExtraerTitulosDelRango(RangoTitulos)
        End If
    End Sub

    ''' <summary>
    ''' Obtiene el libro de trabajo según el nombre o la ruta proporcionada, o el libro activo si no se especifica.
    ''' </summary>
    ''' <param name="workbookName">Nombre del libro de trabajo (opcional).</param>
    ''' <param name="workbookPath">Ruta del archivo del libro de trabajo (opcional).</param>
    ''' <returns>El objeto Workbook correspondiente.</returns>
    Private Function ObtenerWorkbook(ByVal workbookName As String, ByVal workbookPath As String) As Workbook
        If Not String.IsNullOrEmpty(workbookPath) Then
            Try
                Return excelApp.Workbooks.Open(workbookPath)
            Catch ex As Exception
                Throw New Exception("No se pudo abrir el archivo en la ruta especificada.")
            End Try
        ElseIf Not String.IsNullOrEmpty(workbookName) Then
            For Each wb As Workbook In excelApp.Workbooks
                If wb.Name = workbookName Then
                    Return wb
                End If
            Next
            Throw New Exception("El libro especificado no se encontró entre los libros abiertos.")
        Else
            Return excelApp.ActiveWorkbook
        End If
    End Function

    ''' <summary>
    ''' Obtiene la hoja de trabajo según el nombre proporcionado, o la hoja activa si no se especifica.
    ''' </summary>
    ''' <param name="workbook">El libro de trabajo en el que buscar la hoja.</param>
    ''' <param name="worksheetName">Nombre de la hoja de trabajo (opcional).</param>
    ''' <returns>El objeto Worksheet correspondiente.</returns>
    Private Function ObtenerWorksheet(ByVal workbook As Workbook, ByVal worksheetName As String) As Worksheet
        If Not String.IsNullOrEmpty(worksheetName) Then
            Try
                Return workbook.Sheets(worksheetName)
            Catch ex As Exception
                Throw New Exception("No se encontró la hoja especificada en el libro.")
            End Try
        Else
            Return workbook.ActiveSheet
        End If
    End Function

    ''' <summary>
    ''' Identifica los rangos de títulos, datos y totales en la tabla a partir de la celda de inicio.
    ''' </summary>
    ''' <param name="celdaInicio">Celda inicial de la tabla en la hoja de Excel.</param>
    Private Sub IdentificarRangos(ByVal celdaInicio As String)
        Dim rangoInicio As Range = worksheet.Range(celdaInicio)
        Dim rangoFin As Range = rangoInicio.End(XlDirection.xlToRight).End(XlDirection.xlDown)
        Dim rangoTabla As Range = worksheet.Range(rangoInicio, rangoFin)

        RangoTitulos = worksheet.Range(rangoInicio, rangoInicio.End(XlDirection.xlToRight))
        RangoDatos = worksheet.Range(rangoInicio.Offset(1, 0), rangoFin.Offset(-1, 0))
        RangoTotales = worksheet.Range(rangoFin.EntireRow.Cells(1, rangoInicio.Column), rangoFin)
    End Sub

    ''' <summary>
    ''' Extrae los nombres de los títulos del rango de títulos y los convierte en una lista de cadenas.
    ''' </summary>
    ''' <param name="rango">Rango que contiene los títulos en la hoja de Excel.</param>
    ''' <returns>Lista de nombres de los títulos.</returns>
    Private Function ExtraerTitulosDelRango(ByVal rango As Range) As List(Of String)
        Dim titulos As New List(Of String)
        For Each celda As Range In rango.Cells
            titulos.Add(celda.Value.ToString())
        Next
        Return titulos
    End Function

    ''' <summary>
    ''' Genera un rango de títulos en la hoja de cálculo a partir de una lista de nombres de títulos.
    ''' </summary>
    ''' <param name="celdaInicio">Celda inicial donde escribir los títulos en la hoja.</param>
    ''' <param name="titulos">Lista de nombres de títulos.</param>
    ''' <returns>El rango que abarca los títulos generados en la hoja de Excel.</returns>
    Private Function GenerarRangoDeTitulos(ByVal celdaInicio As String, ByVal titulos As List(Of String)) As Range
        Dim columnaInicio As Integer = worksheet.Range(celdaInicio).Column
        Dim filaInicio As Integer = worksheet.Range(celdaInicio).Row

        For i As Integer = 0 To titulos.Count - 1
            worksheet.Cells(filaInicio, columnaInicio + i).Value = titulos(i)
        Next

        Return worksheet.Range(worksheet.Cells(filaInicio, columnaInicio), worksheet.Cells(filaInicio, columnaInicio + titulos.Count - 1))
    End Function

    ''' <summary>
    ''' Copia la tabla (títulos, datos y totales) en una posición paralela a la derecha de la tabla original,
    ''' dejando una separación de dos columnas.
    ''' </summary>
    Public Sub CopiarTablaEnParalelo()
        Dim columnaDestino As Integer = RangoDatos.Columns(RangoDatos.Columns.Count).Column + 2
        Dim rangoDestino As Range = worksheet.Cells(RangoTitulos.Row, columnaDestino)

        Dim rangoTablaCompleta As Range = worksheet.Range(RangoTitulos, RangoTotales)
        rangoTablaCompleta.Copy(rangoDestino)

        For i As Integer = 0 To rangoTablaCompleta.Columns.Count - 1
            worksheet.Columns(columnaDestino + i).ColumnWidth = worksheet.Columns(RangoTitulos.Column + i).ColumnWidth
        Next

        excelApp.CutCopyMode = False
    End Sub

    ''' <summary>
    ''' Verifica si la tabla ha sido copiada en paralelo, comparando los títulos y totales.
    ''' </summary>
    ''' <returns>True si la tabla copiada coincide en títulos y totales, False en caso contrario.</returns>
    Public Function TablaHaSidoCopiada() As Boolean
        ' Calcular la posición de la copia paralela (a la derecha con 2 columnas de separación)
        Dim columnaDestino As Integer = RangoTitulos.Columns(RangoTitulos.Columns.Count).Column + 2
        Dim filaInicio As Integer = RangoTitulos.Row

        ' Definir los rangos de los títulos y los totales en la tabla copiada
        Dim rangoTitulosCopia As Range = worksheet.Range(worksheet.Cells(filaInicio, columnaDestino),
                                                         worksheet.Cells(filaInicio, columnaDestino + RangoTitulos.Columns.Count - 1))

        Dim rangoTotalesCopia As Range = worksheet.Range(worksheet.Cells(RangoTotales.Row, columnaDestino),
                                                         worksheet.Cells(RangoTotales.Row, columnaDestino + RangoTotales.Columns.Count - 1))

        ' Comparar los títulos de la tabla original con los de la copia
        If Not CompararRangos(RangoTitulos, rangoTitulosCopia) Then
            Return False
        End If

        ' Comparar los totales de la tabla original con los de la copia
        If Not CompararRangos(RangoTotales, rangoTotalesCopia) Then
            Return False
        End If

        ' Si ambos coinciden, la tabla ha sido copiada
        Return True
    End Function

    ''' <summary>
    ''' Compara dos rangos de celdas para verificar si tienen el mismo contenido y formato.
    ''' </summary>
    ''' <param name="rango1">Primer rango a comparar.</param>
    ''' <param name="rango2">Segundo rango a comparar.</param>
    ''' <returns>True si ambos rangos coinciden en contenido y formato; False en caso contrario.</returns>
    Private Function CompararRangos(ByVal rango1 As Range, ByVal rango2 As Range) As Boolean
        ' Verificar que ambos rangos tengan el mismo número de celdas
        If rango1.Cells.Count <> rango2.Cells.Count Then
            Return False
        End If

        ' Comparar cada celda individualmente
        For i As Integer = 1 To rango1.Cells.Count
            Dim celda1 As Range = rango1.Cells(i)
            Dim celda2 As Range = rango2.Cells(i)

            ' Comparar el contenido de las celdas
            If celda1.Value IsNot Nothing AndAlso celda2.Value IsNot Nothing Then
                If celda1.Value.ToString() <> celda2.Value.ToString() Then
                    Return False
                End If
            ElseIf celda1.Value IsNot Nothing OrElse celda2.Value IsNot Nothing Then
                Return False
            End If

            ' Comparar el formato de las celdas (tipo de fuente, tamaño de fuente y negrita)
            If celda1.Font.Name <> celda2.Font.Name OrElse
               celda1.Font.Size <> celda2.Font.Size OrElse
               celda1.Font.Bold <> celda2.Font.Bold Then
                Return False
            End If
        Next

        ' Si todo coincide, los rangos son iguales
        Return True
    End Function

    ''' <summary>
    ''' Libera los recursos de Excel utilizados en la clase, como los objetos de rango, hoja de trabajo y libro.
    ''' </summary>
    Public Sub LiberarRecursos()
        ReleaseObject(RangoTitulos)
        ReleaseObject(RangoDatos)
        ReleaseObject(RangoTotales)
        ReleaseObject(worksheet)
        ReleaseObject(workbook)
        ReleaseObject(excelApp)
    End Sub

    ''' <summary>
    ''' Libera un objeto COM de Excel para reducir el uso de memoria.
    ''' </summary>
    ''' <param name="obj">El objeto a liberar.</param>
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
Module Module1
    <CommandMethod("revisarsoporte")>
    Sub Main()
        Try
            ' Ruta al archivo de Excel para pruebas (asegúrate de cambiar esta ruta a un archivo de prueba en tu sistema)
            Dim archivoExcel As String = "C:\ruta\del\archivo\TablaPrueba.xlsx"
            Dim ExRNGSe As New ExcelRangeSelector()
            Dim RNG As Range = ExRNGSe.SelectRange()
            ' Inicializar la clase RevisionSoporteExcel
            Dim revisionExcel As New RevisionSoporteExcel(workbookPath:=RNG.Worksheet.Application.ActiveWorkbook.Path, worksheetName:=RNG.Worksheet.Name, celdaInicio:=RNG.Address)

            ' Copiar la tabla en paralelo
            Console.WriteLine("Copiando tabla en paralelo...")
            revisionExcel.CopiarTablaEnParalelo()
            Console.WriteLine("Tabla copiada.")

            ' Verificar si la tabla ha sido copiada correctamente
            Dim tablaCopiada As Boolean = revisionExcel.TablaHaSidoCopiada()
            If tablaCopiada Then
                Console.WriteLine("La tabla ha sido copiada correctamente.")
            Else
                Console.WriteLine("La tabla NO ha sido copiada correctamente.")
            End If

            ' Liberar recursos
            revisionExcel.LiberarRecursos()

        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        Finally
            ' Evitar que la consola se cierre automáticamente
            Console.WriteLine("Prueba finalizada. Presiona cualquier tecla para salir...")
            Console.ReadKey()
        End Try
    End Sub
End Module