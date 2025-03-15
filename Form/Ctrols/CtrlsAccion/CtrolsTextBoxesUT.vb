Imports System.Windows.Forms
Imports EXCEL = Microsoft.Office.Interop.Excel


Public Class CtrolsTextBoxesUT
	Public Shared Sub AdjustTextBoxValueWithMouseWheel(e As MouseEventArgs, TxTBox As TextBox, CAcadHelp As ACAdHelpers)
		Dim currentValue As Double = 0
		If Double.TryParse(TxTBox.Text.Replace("+", "").Replace(",", "."), currentValue) Then
			If e.Delta > 0 Then
				currentValue += 10
			Else
				currentValue -= 10
			End If
			TxTBox.Text = currentValue.ToString()
			CAcadHelp.CheckThisDrowing()
			Dim Station As Double
			Station = CDbl(TxTBox.Text.Replace("+", "").Replace(",", "."))
			CStationOffsetLabel.GotoStation(Station, CAcadHelp.Alignment, CAcadHelp.ThisDrawing)
		End If
	End Sub

	Public Shared Sub AdjustTextBoxValueWithMouseWheelColumn(e As MouseEventArgs, TxTBox As TextBox, CAcadHelp As ACAdHelpers, ByRef ColumnaKMINICIAL As EXCEL.Range, ByRef excelRange As EXCEL.Range, ByRef ListRNG As List(Of EXCEL.Range))
		Dim currentRow As Integer
		Dim newValue As String = TxTBox.Text

		If newValue = "" Then ColumnaKMINICIAL = Nothing 'OrElse excelRange.EntireRow.Hidden = True Or excelRange.EntireRow.Hidden = True
		'buscar en la hoja activa
		If ColumnaKMINICIAL Is Nothing Then
			FIRSTMOVE(excelRange, {"KM INCIAL", "KM INICIAL", "StartStation", "INICIO"}, ListRNG, ColumnaKMINICIAL)
		Else
			' Determina la dirección de desplazamiento
			If e.Delta > 0 Then
				'Array.IndexOf(ListRNG.ToArray(), excelRange)
				DESPLAZAMIENTOS(excelRange, ListRNG, -1)
			Else
				' Bajar una fila
				DESPLAZAMIENTOS(excelRange, ListRNG, 1)
			End If
		End If

		If excelRange Is Nothing Then Exit Sub


		' Obtiene el valor de la celda en la nueva fila
		Dim cellValue As Object = excelRange.Value

		' Verifica que el valor de la celda sea numérico y en el formato adecuado
		GotoStation(TxTBox, cellValue, CAcadHelp, currentRow)
		excelRange.Activate()
		excelRange.Worksheet.Activate()
	End Sub

	Public Shared Sub AdjustTextBoxValueWitharrowKeyColumn(e As KeyEventArgs, TxTBox As TextBox, CAcadHelp As ACAdHelpers, ByRef ColumnaKMINICIAL As EXCEL.Range, ByRef excelRange As EXCEL.Range, ByRef ListRNG As List(Of EXCEL.Range))
		Dim currentRow As Integer
		Dim newValue As String = TxTBox.Text


		If Array.IndexOf({Keys.Up, Keys.Down}, e.KeyCode) <> -1 Then
			If newValue = "" Then ColumnaKMINICIAL = Nothing 'OrElse excelRange.EntireRow.Hidden = True AndAlso excelRange.EntireRow.Hidden = True
			'buscar en la hoja activa
			If ColumnaKMINICIAL Is Nothing Then 'OrElse excelRange.EntireRow.Hidden = True
				FIRSTMOVE(excelRange, {"KM INCIAL", "KM INICIAL", "StartStation", "INICIO"}, ListRNG, ColumnaKMINICIAL)
			Else
				' Determina la dirección de desplazamiento
				If e.KeyCode = Keys.Up Then
					' Subir una fila
					DESPLAZAMIENTOS(excelRange, ListRNG, -1)
				ElseIf e.KeyCode = Keys.Down Then
					' Bajar una fila
					DESPLAZAMIENTOS(excelRange, ListRNG, 1)
				End If
			End If
		ElseIf Array.IndexOf({Keys.Enter}, e.KeyCode) = -1 Then
			'IfIsInListOfRange(ListRNG, excelRange, TxTBox)
			Exit Sub
		ElseIf Array.IndexOf({Keys.Enter}, e.KeyCode) <> -1 Then
			IfIsInListOfRange(ListRNG, excelRange, TxTBox)
		End If

		If excelRange Is Nothing Then Exit Sub

		' Obtiene el valor de la celda en la nueva fila
		Dim cellValue As Object = excelRange.Value

		' Verifica que el valor de la celda sea numérico y en el formato adecuado
		GotoStation(TxTBox, cellValue, CAcadHelp, currentRow)

		excelRange.Worksheet.Activate()
		excelRange.Activate()

	End Sub

	'crear rutita para buscar en listado un valor ointroducido aleatoreamente 
	Public Shared Sub IfIsInListOfRange(ListRNG As List(Of EXCEL.Range), ByRef excelRange As EXCEL.Range, txtbox As TextBox)
		Try

			If Math.Round(C_strStationToDbl(excelRange.Value), 2) <> Math.Round(C_strStationToDbl(txtbox.Text.ToString()), 2) Then
				For Each rng As EXCEL.Range In ListRNG
					If Math.Round(C_strStationToDbl(rng.Value), 2) = Math.Round(C_strStationToDbl(txtbox.Text), 2) Then
						excelRange = rng
						Exit For
					End If
				Next
			End If
		Catch ex As Exception

		End Try
	End Sub
	Public Shared Sub IfIsInListOfRangehandle(ListRNG As List(Of EXCEL.Range), ByRef excelRange As EXCEL.Range, txtbox As TextBox)
		Try

			If excelRange.Value.ToString() <> txtbox.Text.ToString() Then
				For Each rng As EXCEL.Range In ListRNG
					If rng.Value.ToString() = txtbox.Text Then
						excelRange = rng
						Exit For
					End If
				Next
			End If
		Catch ex As Exception

		End Try
	End Sub
	Public Shared Sub TxtBoxHandleWitharrowKeyColumn(e As KeyEventArgs, TxTBox As TextBox, CAcadHelp As ACAdHelpers, ByRef ColumnaKMINICIAL As EXCEL.Range, ByRef excelRange As EXCEL.Range, ByRef ListRNG As List(Of EXCEL.Range))
		'Dim currentRow As Integer
		Dim newValue As String = TxTBox.Text

		'If Array.IndexOf({Keys.Up, Keys.Down}, e.KeyCode) = -1 Then Exit Sub

		Dim EXCELAPP As New ExcelAppTR()

		If Array.IndexOf({Keys.Up, Keys.Down}, e.KeyCode) <> -1 Then
			If newValue = "" Then ColumnaKMINICIAL = Nothing 'OrElse excelRange.EntireRow.Hidden = True AndAlso excelRange.EntireRow.Hidden = True
			'buscar en la hoja activa
			If ColumnaKMINICIAL Is Nothing Then 'OrElse excelRange.EntireRow.Hidden = True
				FIRSTMOVE(excelRange, {"Handle"}, ListRNG, ColumnaKMINICIAL)
			Else
				' Determina la dirección de desplazamiento
				If e.KeyCode = Keys.Up Then
					' Subir una fila
					DESPLAZAMIENTOS(excelRange, ListRNG, -1)
				ElseIf e.KeyCode = Keys.Down Then
					' Bajar una fila
					DESPLAZAMIENTOS(excelRange, ListRNG, 1)
				End If
			End If
		ElseIf Array.IndexOf({Keys.Enter}, e.KeyCode) = -1 Then
			'IfIsInListOfRange(ListRNG, excelRange, TxTBox)
			Exit Sub
		ElseIf Array.IndexOf({Keys.Enter}, e.KeyCode) <> -1 Then
			'crear un metodo para handle de IfIsInListOfRange
			IfIsInListOfRangehandle(ListRNG, excelRange, TxTBox)
		End If

		If excelRange Is Nothing Then Exit Sub

		' Obtiene el valor de la celda en la nueva fila
		Dim cellValue As Object = excelRange.Value
		newValue = cellValue.ToString()
		TxTBox.Text = newValue
		'TxTBox.Tag = currentRow
		If CLHandle.CheckIfExistHd(TxTBox.Text) Then
			AcadZoomManager.SelectedZoom(TxTBox.Text, CAcadHelp.ThisDrawing)
		Else
			EXCELAPP.GetWorkB().Activate()
		End If
		excelRange.Worksheet.Activate()
		excelRange.Activate()
	End Sub
	Public Shared Sub DESPLAZAMIENTOS(ByRef excelRange As EXCEL.Range, ByRef ListRNG As List(Of EXCEL.Range), DIRECCION As Integer)
		' Determina la dirección de desplazamiento
		'If DIRECCION > 0 Then
		'Array.IndexOf(ListRNG.ToArray(), excelRange)
		If Array.IndexOf(ListRNG.ToArray(), excelRange) = 0 AndAlso DIRECCION < 0 Then
			excelRange = ListRNG(ListRNG.Count - 1)

		ElseIf Array.IndexOf(ListRNG.ToArray(), excelRange) = ListRNG.Count - 1 AndAlso DIRECCION > 0 Then
			excelRange = ListRNG(0)
		Else

			excelRange = ListRNG(Array.IndexOf(ListRNG.ToArray(), excelRange) + DIRECCION)
		End If
	End Sub
	Public Shared Sub FIRSTMOVE(ByRef excelRange As EXCEL.Range, ECABEZADOS As String(), ByRef ListRNG As List(Of EXCEL.Range), ByRef ColumnaKMINICIAL As EXCEL.Range)
		Dim EXCELAPP As New ExcelAppTR()
		Dim Sh As EXCEL.Worksheet = EXCELAPP.GetWorkB()?.ActiveSheet

		Dim DatosFound As EXCEL.Range = Nothing
		For Each Enc As String In ECABEZADOS
			DatosFound = Sh.Cells.Find(What:=Enc, MatchCase:=True, LookAt:=EXCEL.XlLookAt.xlWhole)
			If DatosFound IsNot Nothing Then Exit For
		Next

		If DatosFound Is Nothing Then Exit Sub


		ColumnaKMINICIAL = Sh.Range(DatosFound.Cells(2, 1).Address & ":" & DatosFound.Cells(1, 1).Offset(10000, 0).End(EXCEL.XlDirection.xlUp).Address).SpecialCells(EXCEL.XlCellType.xlCellTypeVisible)

		If ListRNG Is Nothing Then
			ListRNG = New List(Of EXCEL.Range)
		Else
			ListRNG.Clear()
		End If

		For Each Cell As EXCEL.Range In ColumnaKMINICIAL.Cells
			If Cell.EntireRow.Hidden = False Then ListRNG.Add(Cell)
		Next

		excelRange = ListRNG(0)
	End Sub
	Public Shared Sub GotoStation(TxTBox As TextBox, cellValue As Object, ByRef CAcadHelp As ACAdHelpers, ByRef currentRow As Integer)

		' Verifica que el valor de la celda no sea Nothing y conviértelo a número
		Dim cellValueNumeric As Double = C_strStationToDbl(If(cellValue?.ToString(), "0"))

		' Si el valor es válido (mayor que cero)
		If cellValueNumeric > 0 Then
			' Formatea y asigna el valor al TextBox
			TxTBox.Text = cellValueNumeric.ToString("0+000.00")
			TxTBox.Tag = currentRow ' Guarda la fila actual

			' Verifica el estado del dibujo en AutoCAD
			CAcadHelp.CheckThisDrowing()

			' Llama a GotoStation con la estación ya convertida
			CStationOffsetLabel.GotoStation(cellValueNumeric, CAcadHelp.Alignment, CAcadHelp.ThisDrawing)
		End If

	End Sub

	Public Function RNGOffset(RNG As EXCEL.Range, Header As String) As EXCEL.Range

		Dim RNGFound As EXCEL.Range = RNG?.Worksheet.Cells.Find(What:=Header, MatchCase:=True, LookAt:=EXCEL.XlLookAt.xlWhole)
		If RNGFound IsNot Nothing Then
			Return RNG.Offset(0, RNGFound?.Column - RNG?.Column)
		Else
			Return Nothing
		End If
	End Function

	' TabName TextBox key down event
	Public Shared Sub KeyDown(TBox1 As TextBox, TBox2 As TextBox, CAcadHelp As ACAdHelpers, e As KeyEventArgs)

		If e.KeyCode = Keys.Enter Then
			CAcadHelp.CheckThisDrowing()
			If TBox1.Focused = True Then
				KeyDownStation(TBox1, CAcadHelp, e)
			ElseIf TBox2.Focused = True Then
				Dim hdString As String = TBox2.Text
				AcadZoomManager.SelectedZoom(hdString, CAcadHelp.ThisDrawing)
			End If
		End If
	End Sub


	Public Shared Sub KeyDownStation(TBox1 As TextBox, CAcadHelp As ACAdHelpers, e As KeyEventArgs)
		Dim Station As Double
		If e.KeyCode = Keys.Enter Then 'e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Then
			Station = CDbl(Replace(TBox1.Text, "+", ""))
			CStationOffsetLabel.CGotoStation(Station, CAcadHelp)
		End If
	End Sub

	Public Shared Sub GotoSt(StrStation As String, CAcadHelp As ACAdHelpers)
		Dim Station As Double = C_strStationToDbl(StrStation)
		CStationOffsetLabel.CGotoStation(Station, CAcadHelp)
	End Sub
	Public Shared Function C_strStationToDbl(StrStation As String)
		Dim Station As Double
		Double.TryParse(StrStation.Replace("+", "").Replace(",", "."), Station)
		Return Station
	End Function
End Class

Public Class UtlTextB
	Public Property EventHandlerParmetrict As New List(Of EventHandlerParametricClass)
	Public Sub New()

		'	''Eventos de key en textboxes Nested parameters (sender As Object, e As KeyEventArgs)
		'	EventHandlerParmetrict.Add(New EventHandlerParmetrictClass({TextBox1, TextBox2}, "KeyDown", Sub(sender As Object, e As KeyEventArgs)
		'																									CtrolsTextBoxesUT.KeyDown(TextBox1, TextBox2, CAcadHelp, sender, e)
		'																								End Sub))
	End Sub
End Class
