Imports Microsoft.Office.Interop.Excel

'Public Class CExcelApp
'    Public Function ExcelApp() As Object

'        ' Specify the process ID of the Excel instance you want to connect to
'        'Dim excelProcessID As Integer = ' Replace with the actual process ID
'        Dim xlApp As Object = CreateObject("Excel.Application")
'        Dim xlwkbook As Workbook = xlApp.ActiveWorkbook
'        Dim xlSh As Worksheet = xlwkbook.ActiveSheet
'        Return xlSh
'    End Function
'End Class
' Get the Excel application instance using the process ID
'Dim excelApp As Application = Marshal.GetActiveObject("Excel.Application")

'Check If the process ID matches
'If excelApp.Hwnd = excelProcessID Then
'    MessageBox.Show("Excel instance found!")
'Else
'    MessageBox.Show("Excel instance not found.")
'End If