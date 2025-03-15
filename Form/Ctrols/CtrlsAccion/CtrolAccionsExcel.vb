Imports System.Windows.Forms
'CtrolAccionsExcel.MoverCunetaAExistente
Public Class CtrolAccionsExcel
	Public Shared Sub MoverCunetaAExistente(DGdView1 As DataGridView, DGView2 As DataGridView, DGView3 As DataGridView, ByRef handleProcessor As HandleDataProcessor, Headers As List(Of String), CAcadHelp As ACAdHelpers) 'Handles Button3.Click
		Using OP As New ExcelOptimizer
			Try
				OP.TurnEverythingOff()
				'Dim Headers As List(Of String) = columnData.Select(Function(item) item.Item1).ToList()
				UpdateExcelTable.CheckExcelRange(CAcadHelp, Headers, DGView2, handleProcessor)
				handleProcessor.AddInfo(DGdView1, DGView2, DGView3, CAcadHelp)
			Catch ex As Exception
				OP.TurnEverythingOn()
				Diagnostics.Debug.WriteLine("ERROR: CLICK BUTTON3: " & ex.Message())
			Finally
				OP.TurnEverythingOn()
				OP.Dispose()
			End Try
		End Using
	End Sub

End Class
