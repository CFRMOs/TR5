'Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.EditorInput
'Imports Microsoft.Office.Interop.Excel
'Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

'Public Class AutoCADExcelIntegration
'    <CommandMethod("CheckExcelRange")>
'    Public Sub CheckExcelRange()
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim ed As Editor = doc.Editor

'        Try
'            Dim selectedRange As Range = ExcelRangeSelector.SelectRange()

'            If selectedRange IsNot Nothing Then
'                Dim cellValues As New List(Of String)()

'                For Each cell As Range In selectedRange
'                    cellValues.Add(cell.Value.ToString())
'                Next

'                For Each value As String In cellValues
'                    ed.WriteMessage(value & vbLf)
'                Next
'            End If
'        Catch ex As Exception
'            ed.WriteMessage("Error: " & ex.Message & vbLf)
'        End Try
'    End Sub
'End Class
