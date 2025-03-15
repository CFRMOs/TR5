'Clase para manejar las operaciones de Excel
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class ExcelManager
    Dim excelApp As Excel.Application

    ' Obtiene una instancia activa de Excel o inicia una nueva
    Public Function GetExcelApp() As Excel.Application
        If IsExcelAppActive() Then
            Return excelApp
        Else
            Return StartNewExcelApp()
        End If
    End Function

    ' Encuentra una tabla por índice (n) en una hoja de cálculo específica
    Public Function FindTableByN(ByVal n As Integer, ByVal worksheet As Excel.Worksheet) As Excel.ListObject
        Dim sortedTables = worksheet.ListObjects.Cast(Of Excel.ListObject)().OrderBy(Function(tbl) tbl.Range.Row)
        Dim tableIndex As Integer = 0

        For Each table As Excel.ListObject In sortedTables
            tableIndex += 1
            If tableIndex = n Then
                Return table
            End If
        Next
        ' Si no se encuentra una tabla en la posición especificada, devolver Nothing
        Return Nothing
    End Function

    ' Verifica si una instancia de Excel ya está activa
    Private Function IsExcelAppActive() As Boolean
        Try
            Dim processes() As Process = Process.GetProcessesByName("EXCEL")
            If processes.Length > 0 Then
                excelApp = CType(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
                Return True
            Else
                ' Si no hay procesos de Excel en ejecución, intentamos crear una nueva instancia
                Return StartNewExcelApp() IsNot Nothing
            End If
        Catch ex As Exception
            ' Si se produce un error, devolvemos False
            Return False
        End Try
    End Function

    ' Inicia una nueva instancia de Excel
    Private Function StartNewExcelApp() As Excel.Application
        excelApp = New Excel.Application With {
            .Visible = True
        }
        Return excelApp
    End Function
End Class

'Clase para extraer datos de una tabla en Excel
Public Class ExcelHelper
    Private Shared ReadOnly MainExcelPath As String

    '#Const MainExcelPath = "D:\Desktop\Typsa - Las Placetas\1.0-Mediciones\Resumen Hormigones Typsa Med-13-R3 03-09-2024 1420 CF.xlsm"' "\\MSI-CARLOS\1.0-Mediciones\Resumen Hormigones Typsa Med-13-R3 03-09-2024 1420 CF.xlsm"
    ' F2: Extraer los datos de la tabla de Excel
    ' Esta función será Shared para que se pueda llamar sin necesidad de instanciar la clase
    ' Constructor
    Public Shared ExcelApp As ExcelAppTR

    Public Sub New()
        ExcelApp = New ExcelAppTR()
    End Sub
    Public Shared Function ObtenerDatosTabla(TableName As String, ACCESO As Double) As List(Of Object)
        Dim valores As New List(Of Object)
        Dim ExcelRangeSelector As New ExcelRangeSelector
        Dim listObject As Excel.ListObject = ExcelRangeSelector.GetTableOnWorkBkByName(TableName)
        Dim j As Integer = 0

        Try
            ' Recorrer la tabla y extraer HandleDataItems
            For Each row As Excel.ListRow In listObject.ListRows
                Console.WriteLine(row.Range.Row)
                If row.Range(1).Value2 = ACCESO Then
                    Debug.Assert(row.Range(2).Row <> 278)

                    ' Fix to handle empty or null cells
                    If row.Range(2).Value2 Is Nothing OrElse String.IsNullOrEmpty(row.Range(2).Value2.ToString()) Then
                        row.Range(2).Value2 = "NOT" & j
                        j += 1
                    End If

                    valores.Add(row.Range(2).Value2)
                End If
            Next

            ' Return the collected values
            Return valores

        Catch ex As Exception
            ' TabName exceptions gracefully
            Console.WriteLine("Error: " & ex.Message)
            Return New List(Of Object) ' Return an empty list instead of HandleDataItems(0)
        End Try
    End Function
    ' Esta función obtiene las filas de una tabla en Excel que corresponden al handle en base al valor de ACCESO
    Public Shared Function ObtenerDatosTabla2(TableName As String, ACCESO As Double, Optional WBFilePath As String = vbNullString) As Dictionary(Of String, CunetasHandleDataItem)  ', Headers As IEnumerable(Of String)) As List(Of Object)
        Dim HandleDataItems As New Dictionary(Of String, CunetasHandleDataItem)
        Dim ExcelRangeSelector As New ExcelRangeSelector(ExcelApp)
        Dim listObject As Excel.ListObject = ExcelRangeSelector.GetTableOnWorkBkByName(TableName, WBFilePath) ' Obtener la tabla de Excel por nombre
        Dim j As Integer = 0
        'ExcelApp = ExcelRangeSelector.ExcelApp
        Dim OP As New ExcelOptimizer(ExcelApp.GetWorkB().Application)
        OP.TurnEverythingOff()
        '' Aquí calculamos los índices de las columnas para `Headers` una sola vez
        Dim IndexHeaders As Integer() = Nothing

        Try
            'crear array de indexCol as integer()=

            ' Recorrer las filas de la tabla en Excel y extraer HandleDataItems
            For Each row As Excel.ListRow In listObject.ListRows
                ' Comprobar si la primera columna coincide con el valor de ACCESO
                If row.Range(1).Value2 = ACCESO Then
                    ' Si la celda está vacía o es nula, asignar un valor predeterminado
                    If row.Range(2).Value2 Is Nothing OrElse String.IsNullOrEmpty(row.Range(2).Value2.ToString()) Then
                        row.Range(2).Value2 = "NOT" & j
                        j += 1
                    End If

                    ' Crear una instancia de CunetasHandleDataItem
                    Dim handleDataItem As New CunetasHandleDataItem(IndexHeaders, listObject)

                    ' Convertir la lista de valores a un array y asignar a las propiedades del CunetasHandleDataItem
                    handleDataItem.SetPropertiesFromTableRng(row.Range)

                    ' Verificar si el handle ya existe en el diccionario antes de agregarlo
                    If Not HandleDataItems.ContainsKey(handleDataItem.Handle) Then
                        ' Agregar el CunetasHandleDataItem al diccionario usando el TabName como clave
                        HandleDataItems.Add(handleDataItem.Handle, handleDataItem)
                    Else
                        ' Si ya existe, puedes manejarlo de otra manera (como actualizar o ignorar)
                        Console.WriteLine($"El handle {handleDataItem.Handle} ya existe en el diccionario.")
                    End If
                End If
            Next

            ' Devolver los HandleDataItems recolectados
            Return HandleDataItems

        Catch ex As Exception
            ' Manejar las excepciones y devolver una lista vacía si ocurre un error
            Console.WriteLine("Error: " & ex.Message)
            Return New Dictionary(Of String, CunetasHandleDataItem) ' En caso de error, devolver una lista vacía
        Finally
            OP.TurnEverythingOn()
        End Try
    End Function

End Class
