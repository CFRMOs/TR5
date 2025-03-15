Imports System.Diagnostics
Imports System.Linq
Imports System.Windows.Controls.Primitives
Imports System.Windows.Forms
'Imports ExcelInteropManager2
Public Class PolylineMG
	Inherits Form

	' Class variables
	Public selectedLayers As New List(Of String)

	Dim selectedLayer As String

	Dim hdString As String

	Public nombresCapas As String()

	Public DCapas As New Dictionary(Of String, String())

	Private Const filePath As String = ""

	Public WithEvents VBListRelateTramos As New ListadoRelacionados

	Public WithEvents CAcadHelp As New ACAdHelpers

	Public Globalccv As CCVDataClosedPL

	Private dataGridViewEventHandlers As New List(Of DataGridViewEventHandler)

	Public PDrenajeSuperficial As PosicionDrenajeSuperficial

	Dim columnData As List(Of (String, String, String)) = HandleDataProcessor.Headers

	Dim columnaformatos As List(Of String) = HandleDataProcessor.Headers.Select(Function(t) t.Item3).ToList()

	Dim ListMediciones As String() = {"MED-09", "MED-10", "MED-11", "MED-12", "MED-13", "MED-14", "MED-15", "MED-16"}

	Dim TiposCunetas As String() = {"BA-T1", "BA-T2", "DE-01", "DA-01", "DA-02", "CU-BO", "CU-BO1", "CU-T1", "CU-T2", "CU-T3", "ZJ-DR1",
									"ZJ-DR2", "ZJ-DR3", "PV-T3", "PV-T2", "PV-T1", "BADEN", "CU-TIPO-LIBRITO", "LOSAS", "Paso Vehicular",
									"Losa-Aproximacion", "Aceras", "Bordillos de PV"
										}
	Public handleProcessor As HandleDataProcessor

	Public EventsHandler As New List(Of EventHandlerClass)

	Public EventHandlerParmetrict As New List(Of EventHandlerParametricClass)

	Public WBFilePath As String = CAcadHelp.WBFilePath

	Public Utl As CommondUtl

	Public ControlManager As New ControlManager(Me)

	Dim tabControlDGView As TabControl = ControlManager._tabControlDGView

	Public PanelDGViews As TabDGViews 'New TabDGViews(tabControlDGView, columnFormats, CAcadHelp, Me, handleProcessor) With {.TabName = "Cunetas Existentes", .DGViewName = "CunetasExistentesDGView"}

	Public Property CunetasExistentesDGView As DataGridView = Nothing

	Public Property DGViewByLeght As DataGridView = Nothing

	Public Property CunetasExternasDGView As DataGridView = Nothing

	Public Property CunetasFueraDeAlcanceDGView As DataGridView = Nothing

	Public Property PuntosCCVDGView As DataGridView = Nothing

	Public Property Diseños As DataGridView = Nothing

	Public UTLBaseDAtos As UTLBaseDAtos

	Public DataTransferManager As New DataTransferManager(tabControlDGView, columnData)

	' Variable para manejar AutoCAD
	'Private autoCADManager As AutoCADEmbedder
	' Constructor del formulario
	Public Sub New()
		InitializeComponent()


		handleProcessor = New HandleDataProcessor(CAcadHelp, ControlManager._tabControlMenu, tabControlDGView, WBFilePath, ComboBox1, ComMediciones)

		PanelDGViews = New TabDGViews(tabControlDGView, columnaformatos, CAcadHelp, Me, handleProcessor)

		Dim tabVars As New List(Of Tuple(Of String, String)) From {
																	   Tuple.Create("CunetasExistentesDGView", "Cunetas Existentes"),
																	   Tuple.Create("DGViewByLeght", "Cunetas por Longitud"),
																	   Tuple.Create("CunetasExternasDGView", "Cunetas Referencias Externas"),
																	   Tuple.Create("CunetasFueraDeAlcanceDGView", "Cunetas Fuera de Proyecto"),
																	   Tuple.Create("PuntosCCVDGView", "Puntos CCV")
				}
		' Itera sobre cada tupla en la lista
		For Each tabVar As Tuple(Of String, String) In tabVars
			' Configura las propiedades de PanelDGViews con los valores correspondientes
			PanelDGViews.TabName = tabVar.Item2       ' Asigna "Cunetas por Longitud"
			PanelDGViews.DGViewName = tabVar.Item1    ' Asigna "DGViewByLeght"

			' Usa reflexión para asignar PanelDGViews.DGView a la propiedad correspondiente
			Dim prop = Me.GetType().GetProperty(tabVar.Item1)  ' Obtiene la propiedad por nombre
			If prop IsNot Nothing AndAlso prop.CanWrite Then   ' Verifica si la propiedad existe y es asignable
				prop.SetValue(Me, PanelDGViews.DGView)         ' Asigna PanelDGViews.DGView al valor de la propiedad
			Else
				' Manejo de error: la propiedad no existe o no es asignable
				Console.WriteLine($"La propiedad '{tabVar.Item1}' no existe o no se puede asignar.")
			End If
		Next


		ControlManager.MoveTabCtrolMittleRighttDown(Me, TabControl1, 5, tabControlDGView)


		PDrenajeSuperficial = New PosicionDrenajeSuperficial(CAcadHelp, tabControlDGView, ControlManager._tabControlMenu)


		'handleProcessor.PanelMediciones1.buttonActions.Add("Pasar Dato a Excel", Sub() ExportDataToExcel2())
		Utl = New CommondUtl(ControlManager._tabControlMenu, tabControlDGView, CAcadHelp, ComboBox1, ComMediciones, handleProcessor)

		'pestaña base de Datos 
		UTLBaseDAtos = New UTLBaseDAtos(ControlManager._tabControlMenu, tabControlDGView, CAcadHelp, ComboBox1, ComMediciones, handleProcessor)

		'esta parte del codigo se encarga de dividir las polilineas en un plano 
		'queda pendientes fallas en el codigo cuando se manejan 2d y 3d polyline
		EventsHandler.Add(New EventHandlerClass(DividirL, "click", Sub() CtrolsButtonsUT.DividePL(CAcadHelp.Alignment, TextBox3)))

		EventsHandler.Add(New EventHandlerClass(btSetAllLayerColor, "click", Sub() LayerColorChanger.SetAllGPLYColor(ListMediciones)))

		EventsHandler.Add(New EventHandlerClass(IDDatagView, "click", Sub() DateViewSet.IdDataFEnt(GetControlsDGView.GetDGView(tabControlDGView))))

		''Eventos de key en textboxes Nested parameters (sender As Object, e As KeyEventArgs)
		EventHandlerParmetrict.Add(New EventHandlerParametricClass({TextBox1, TextBox2}, "KeyDown", Sub(sender As Object, e As KeyEventArgs)
																										CtrolsTextBoxesUT.KeyDown(TextBox1, TextBox2, CAcadHelp, e)
																									End Sub))

		EventsHandler.Add(New EventHandlerClass(SetCurrent, "click", Sub() CLayerHelpers.SetCurrentLayer(ComboBox1.Text)))

		EventsHandler.Add(New EventHandlerClass(BtDataClear, "click", Sub() GetControlsDGView.GetDGView(tabControlDGView).Rows.Clear()))

	End Sub

	'' Form Load Event
	' Evento Load del formulario
	Private Sub Polyline_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Seleccion.Initialize()
	End Sub
	Public Sub SetDatosIni()
		'Dim rows As DataGridViewRowCollection = CunetasExistentesDGView.Rows
		CAcadHelp.Form = Me
		'rows.Clear()
		DCapas = ClGroupLayerHelper.CrearGruposCapas(ListMediciones, TiposCunetas)

		For Each DGView As DataGridView In {CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView}
			DataGridViewHelper.Addcolumns(DGView, columnData)
		Next
		ComboBox2.DataSource = nombresCapas
		ComboBox1.DataSource = nombresCapas
		ComMediciones.DataSource = ListMediciones.ToList()
		If CAcadHelp.List_ALing.Count <> 0 Then
			CoBoxAlignments.DataSource = CAcadHelp.List_ALingName
		End If
	End Sub

	' TabName ComboBox key down event
	Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown, ComboBox1.KeyDown
		HandleComboBoxKeyDown(e)
		CLayerHelpers.SetCurrentLayer(ComboBox1.Text)
	End Sub

	' TabName ComboBox selection changed event
	Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
		UpdateDataDGV.UpdateDataGridViewForSelectedLayer(CunetasExistentesDGView, CAcadHelp.Alignment, ComboBox2.SelectedValue, CAcadHelp.ThisDrawing)
	End Sub

	' TabName ComboBox key down event for adding new layer
	Private Sub HandleComboBoxKeyDown(Optional e As KeyEventArgs = Nothing)

		If CAcadHelp.Alignment Is Nothing OrElse e.KeyCode <> Keys.Enter Then Exit Sub

		nombresCapas = DCapas(ComMediciones.SelectedItem.ToString())

		Dim CBox As ComboBox = If(ComboBox2.Focused, ComboBox2, ComboBox1)

		Dim selectedLayer As String = Trim(CBox.Text)

		UpdateDataDGV.SETLYGroup(selectedLayer, nombresCapas, DCapas, ComMediciones.SelectedItem.ToString())

		nombresCapas = DCapas(ComMediciones.SelectedItem.ToString())

		SetComboList(CBox, nombresCapas)

		CunetasExistentesDGView.Rows.Clear()

		GetDataByLayer(selectedLayer, CunetasExistentesDGView)
	End Sub

	Public Sub GetDataByLayer(selectedLayer As String, DataGridView As DataGridView)
		If selectedLayer = String.Empty Then Exit Sub
		Dim form As Form = DataGridView.FindForm()
		form.Update()
		DateViewSet.SetViewCellEnt(CAcadHelp.ThisDrawing, CAcadHelp.Alignment, form, selectedLayer, DataGridView)
		form.Show()
		form.Activate()
	End Sub

	Private Sub GotoStationIN_Click(sender As Object, e As EventArgs)
		Dim ExRNGSe As New ExcelRangeSelector()
		Dim Station As Double = ExRNGSe.GetdoubleData() ' Start from the first cell

		If Station = 0 Then Exit Sub
		AlignmentLabelHelper.CFindLabelStation(Station, CAcadHelp)
		TextBox1.Text = Station
	End Sub

	Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComMediciones.SelectedIndexChanged
		Dim CBox As ComboBox = If(ComboBox2.Focused, ComboBox2, ComboBox1)
		nombresCapas = DCapas(ComMediciones.SelectedItem.ToString())
		SetComboList(CBox, nombresCapas)
		LayerColorChanger.SetcurrentMed(ComMediciones)
	End Sub
	Public Sub SetComboList(CBox As ComboBox, nombresCapas As String())
		Dim CIndex As Integer = CBox.SelectedIndex
		CBox.DataSource = nombresCapas
		CBox.SelectedIndex = CIndex
	End Sub
	Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComMediciones.KeyDown
		Dim Index As Integer = If(e.KeyCode = Keys.Up Or e.KeyCode = Keys.Enter, ComMediciones.SelectedIndex - 1, ComMediciones.SelectedIndex + 1)
		If Index >= 0 AndAlso Index < ComMediciones.Items.Count Then
			LayerColorChanger.SetcurrentMed(ComMediciones)
		End If
	End Sub

	Private Sub BtFdLineByLen_Click(sender As Object, e As EventArgs)
		Dim ExRNGSe As New ExcelRangeSelector()
		Dim Handle As String = ""
		IdentifyPL.FindPolyline(CAcadHelp, Handle)
		AcadZoomManager.SelectedZoom(Handle, CAcadHelp.ThisDrawing)
		'agregar info a DGview1
		If Handle = "" Then Exit Sub
		DateViewSet.CAddAnEntity(CAcadHelp, CunetasExistentesDGView, Handle)
	End Sub
	Private Sub Button4_Click(sender As Object, e As EventArgs)
		Dim DGView1 As DataGridView = Me.CunetasExistentesDGView
		Dim DGView2 As DataGridView = Me.DGViewByLeght
		' Verificar si los hndle existen en AutoCAD
		Dim Listhandles As (List(Of String), List(Of String)) = AutoCADHelper.ProcesarHandlesDesdeDoc(CAcadHelp.ThisDrawing.Name, "CunetasGenerales")

		' Check if either of the lists inside the tuple is Nothing or empty
		If Listhandles.Item1 Is Nothing OrElse Listhandles.Item2 Is Nothing OrElse
Listhandles.Item1.Count = 0 OrElse Listhandles.Item2.Count = 0 Then
			Exit Sub
		End If
		For Each DGView As DataGridView In {DGView1, DGView2}
			DGView.Rows.Clear()
		Next
		For Each HD As String In Listhandles.Item1
			IdentifyPL.GetDataFromTbl(CAcadHelp, HD)
			DateViewSet.CAddAnEntity(CAcadHelp, DGView1, HD)
		Next
		For Each HD As String In Listhandles.Item2
			DateViewSet.CAddAnEntity(CAcadHelp, DGView2, IdentifyPL.FindPolyline(CAcadHelp, HD), True)
		Next
		For Each DGView As DataGridView In {DGView1, DGView2}
			DGView.Update()
			DGView.Show()
		Next
	End Sub

	Private Sub CoBoxAlignments_SelectedValueChanged(sender As Object, e As EventArgs) Handles CoBoxAlignments.SelectedValueChanged
		Dim index As Integer = CoBoxAlignments.SelectedIndex
		CAcadHelp.Alignment = CAcadHelp.List_ALing(index)
	End Sub
End Class
