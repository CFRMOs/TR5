' todos los comentarios hechos por mi deberan de permanece
' en esta clase no se manejaran librerias com de autocad. esos proceso se maneja en clase dedicadas para ese objetivo 
' devuelme siempre la actualizacion de esta pagina competa 
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Public Class HandleDataProcessor
    Public EventsHandler As New List(Of EventHandlerClass)
    Public EventHandlerParmetrict As New List(Of EventHandlerParametricClass)
    Public Shared ReadOnly Headers As New List(Of (String, String, String)) From {
                                                ("Accesos", "Accesos", """Acceso ""00"),
                                                ("Handle", "Handle", "@"),
                                                ("Layer", "Layer", "@"),
                                                ("MinX", "Min X", "0.00"),
                                                ("MinY", "Min Y", "0.00"),
                                                ("MaxX", "Max X", "0.00"),
                                                ("MaxY", "Max Y", "0.00"),
                                                ("StartStation", "Start Station", "0+000.00"),
                                                ("EndStation", "End Station", "0+000.00"),
                                                ("Longitud", "Longitud", "0.00"),
                                                ("Area", "Area", "0.00"),
                                                ("Side", "Side", "@"),
                                                ("AlignmentHDI", "Alignment HDI", "@"),
                                                ("Comentarios", "Comentarios", "@"),
                                                ("Tramos", "Tramos", "@"),
                                                ("PLANO", "PLANO", "@"),
                                                ("FileName", "File Name", "@"),
                                                ("FilePath", "File Path", "@"),
                                                ("CodigoCuneta", "Codigo Cuneta", "@"),
                                                ("PosicionFormatted", "Posicion Formatted", "@"),
                                                ("Type", "Type", "@"),
                                                ("IDNum", "IDNum", "00"),
                                                ("Closed", "Closed", "@")
        }
    Dim InHeaders As List(Of String) = Headers.Select(Function(item) item.Item1).ToList()
    ' Main dictionary to store handle data (key: handle, value: object or structured data)
    ' Diccionario principal para almacenar los datos del handle
    Public ReadOnly HandleData As Dictionary(Of String, CunetasHandleDataItem)

    ' Public dictionaries for specific analyses
    ' Diccionarios públicos para análisis específicos

    ' esto proviene desde la tabla de excel 
    Public DicHandlesExistentes As New Dictionary(Of String, CunetasHandleDataItem) ' Handles that exist or meet a certain condition
    Public DicHandlesNoExistentes As New Dictionary(Of String, CunetasHandleDataItem) ' Handles that exist or meet a certain condition
    Public DicHandlesByLength As New Dictionary(Of String, CunetasHandleDataItem) ' Handles categorized by length
    Public DicHandlesExternos As New Dictionary(Of String, CunetasHandleDataItem) ' External handles
    Public DicHandlesFueradeProyecto As New Dictionary(Of String, CunetasHandleDataItem) ' External handles

    ' proviene desde el archivo de autocad 
    Public DicHandleCalculados As New Dictionary(Of String, CunetasHandleDataItem) ' estos pasaran por el proceso de calculo segun grupos y layers 
    Public FilePath As String = String.Empty
    Public CAcadHelp As ACAdHelpers = Nothing
    Public CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView As DataGridView
    Public AligmentsCertify As New AlignmentHelper()
    Public PanelMediciones1 As TabsMenu
    Public PanelMediciones2 As TabsMenu
    Public EncabezadosSoporte As Range
    Public cListType As New List(Of List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId))

    Public excelRange As Excel.Range = Nothing
    Public excelRangeHandle As Excel.Range = Nothing
    Public ColumnRange As Excel.Range = Nothing
    Public ColumnRangeHandle As Excel.Range = Nothing
    Public ListRNG As List(Of Excel.Range)
    Public ListRNGHandle As List(Of Excel.Range)

    ' Constructor to initialize the main dictionary
    ' Constructor para inicializar el diccionario principal
    Public Sub New(CAcadHelp As ACAdHelpers, TabCtrlPanol As TabControl, TabCtrolView As TabControl, WBFilePath As String, ComboBox As ComboBox, ComMediciones As ComboBox)
        HandleData = New Dictionary(Of String, CunetasHandleDataItem)
        FilePath = WBFilePath
        CAcadHelp = CAcadHelp

        PanelMediciones1 = New TabsMenu(TabCtrlPanol) With {.TabName = "Mediciones1"}
        PanelMediciones2 = New TabsMenu(TabCtrlPanol) With {.TabName = "Mediciones2"}

        Dim columnData As List(Of String) = Headers.Select(Function(item) item.Item3).ToList()

        Dim Pcl As New ParcelCommands

        'Pcl.CreateParcelPolyline()
        Dim CCunetaPlanta As New CCunetasAreaPlanta

        Dim BoundaryProcessor As New BoundaryProcessor

        Dim buttonActions1 As New Dictionary(Of String, System.Action) From {
            {"Add To Active DGView", Sub() CtrolsButtonsUT.AddAnEntty(TabCtrolView, CAcadHelp)},
            {"Add By Layer To Active DGView", Sub() CtrolsButtonsUT.AddAnEntty(TabCtrolView, CAcadHelp, CSelectionHelper.GetLayerByEnt())},
            {"Add From DGView to Excel", Sub() UpDateExcel(TabCtrolView)},
            {"Reemplazar dato en Excel", Sub() ReplaceDataExcel(TabCtrolView)},
            {"UP Date DGView", Sub() UpdateDataDGV.CUPDateDGView(GetControlsDGView.GetDGView(TabCtrolView), CAcadHelp)},
            {"Cambiar Layer", Sub() CtrolsButtonsUT.CambiarLayer(ComboBox, ComMediciones, TabCtrolView, CAcadHelp)},
            {"Area en planta de cuneta", Sub() CCunetaPlanta.CrearAreaEnPlanta()},
            {"BOUNDARY_DIFFERENCE", Sub() BoundaryProcessor.ExecuteBoundaryDifference()},
            {"Convet to Polyline", Sub() CtrolsButtonsUT.CFeatureLine()},
            {"Copy Entity From File Externo", Sub() CtrolsButtonsUT.CopyEntityFromC(GetControlsDGView.GetDGView(TabCtrolView), CAcadHelp, Me)},
            {"Convertir Parcela", Sub() Pcl.CreateParcelPolyline()},
            {"Create Strip", Sub() PolylineOperations.CreateStripAndOrderVertices()}'CreateStripAndOrderVertices()
                }
        'BOUNDARY_DIFFERENCE
        'ExecuteBoundaryDifference()
        Dim buttonActions2 As New Dictionary(Of String, System.Action) From {
                                            {"SelectedByentity", Sub() SelectedByentity()},'CAcadHelp.ThisDrawing.Name, "CunetasGeneral")},'WBFilePath
                                            {"Selected By entity From Soporte", Sub() SelectedByentitySoporte(TabCtrolView)},
                                            {"Match Soporte vs Table", Sub() MatchSoportevsTable(TabCtrolView)},
                                            {"Add Data", Sub() CLoadData(CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView, CAcadHelp)},
                                            {"MoverCunetaAExistente", Sub() CtrolAccionsExcel.MoverCunetaAExistente(CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView, Me, Headers.Select(Function(item) item.Item1).ToList(), CAcadHelp)},
                                            {"check Aliment By Proximida", Sub() AligmentsCertify.CheckAlimentByProximida(cSelect:=True)},
                                            {"Test By Lenght Only", Sub() AutoCADHelper2.ByLenghtAndStation()},
                                            {"UPDATEDICTIONARIOS", Sub() LoadData(CAcadHelp.ThisDrawing.Name, "CunetasGeneral", WBFilePath)},
                                            {"UpData By Leght", Sub() ByLeght(CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView, CAcadHelp)}
                                        }
        'GetDataByLayer
        PanelMediciones1.ButtonActions = buttonActions1

        PanelMediciones2.ButtonActions = buttonActions2

        ' caja de texto para rapido acceso a las progresivas he entidades relacionadas con listado de excel 
        With New GroupBoxManager("Station:", New System.Drawing.Size(170, 25), New System.Drawing.Point(0, 0), PanelMediciones1.TabPage)
            EventHandlerParmetrict.Add(New EventHandlerParametricClass(.TextBoxes(0), "mousewheel", Sub(sender As Object, e As MouseEventArgs)
                                                                                                        CtrolsTextBoxesUT.AdjustTextBoxValueWithMouseWheel(e, .TextBoxes(0), CAcadHelp)
                                                                                                    End Sub))


            EventHandlerParmetrict.Add(New EventHandlerParametricClass(.TextBoxes(0), "KeyDown", Sub(sender As Object, e As KeyEventArgs)
                                                                                                     CtrolsTextBoxesUT.KeyDownStation(.TextBoxes(0), CAcadHelp, e)
                                                                                                 End Sub))

            .AddNewGBoxTextBox("Station From KM INICIAL:", New System.Drawing.Size(170, 25), New System.Drawing.Point(0, 0))


            EventHandlerParmetrict.Add(New EventHandlerParametricClass(.TextBoxes(1), "mousewheel", Sub(sender As Object, e As MouseEventArgs)
                                                                                                        CtrolsTextBoxesUT.AdjustTextBoxValueWithMouseWheelColumn(e, .TextBoxes(1), CAcadHelp, ColumnRange, excelRange, ListRNG)
                                                                                                    End Sub))

            EventHandlerParmetrict.Add(New EventHandlerParametricClass(.TextBoxes(1), "KeyDown", Sub(sender As Object, e As KeyEventArgs)
                                                                                                     CtrolsTextBoxesUT.AdjustTextBoxValueWitharrowKeyColumn(e, .TextBoxes(1), CAcadHelp, ColumnRange, excelRange, ListRNG)
                                                                                                 End Sub))
            .AddNewGBoxTextBox("Handle From Table:", New System.Drawing.Size(170, 25), New System.Drawing.Point(0, 0))



            EventHandlerParmetrict.Add(New EventHandlerParametricClass(.TextBoxes(2), "KeyDown", Sub(sender As Object, e As KeyEventArgs)
                                                                                                     If Array.IndexOf({Keys.Enter}, e.KeyCode) = -1 Then Exit Sub
                                                                                                     AcadZoomManager.SelectedZoom(.TextBoxes(2).Text, CAcadHelp.ThisDrawing)
                                                                                                 End Sub))
            Dim TXTBOX As String = ""

            EventHandlerParmetrict.Add(New EventHandlerParametricClass(.TextBoxes(2), "KeyDown", Sub(sender As Object, e As KeyEventArgs)
                                                                                                     CtrolsTextBoxesUT.TxtBoxHandleWitharrowKeyColumn(e, .TextBoxes(2), CAcadHelp, ColumnRangeHandle, excelRangeHandle, ListRNGHandle)
                                                                                                     TXTBOX = .TextBoxes(2).Text()
                                                                                                     CtrolsButtonsUT.AddAnEntty(TabCtrolView, CAcadHelp, String.Empty, False, .TextBoxes(2).Text())
                                                                                                 End Sub))

        End With

    End Sub

    ''' <summary>
    ''' Selecciona una entidad en AutoCAD y realiza operaciones basadas en los datos de una tabla de Excel.
    ''' El método busca handles en la tabla de Excel y realiza zoom sobre la entidad encontrada en AutoCAD.
    ''' </summary>
    Public Sub SelectedByentity()
        ' Diccionarios para manejar datos de handles y los resultados procesados
        Dim DicHandles As New Dictionary(Of String, CunetasHandleDataItem)
        Dim DicHandlesreSultado As New Dictionary(Of String, CunetasHandleDataItem)

        ' Selector de rangos en Excel
        Dim ExRNGSe As New ExcelRangeSelector()

        ' Encabezados necesarios para seleccionar datos de la tabla de Excel
        Dim iHeaders As List(Of String) = Headers.Select(Function(item) item.Item1).ToList()
        Dim EXheaders As New List(Of String) From {"CodigoCuneta", "PosicionFormatted", "Type"}

        ' Obtener una fila de datos de la tabla "CunetasGeneral" usando los encabezados
        Dim Data As List(Of Object) = ExRNGSe.SelectRowOnTbl("CunetasGeneral", iHeaders)

        ' Seleccionar el rango de la fila en la hoja de cálculo
        Dim RNG As Excel.Range = ExRNGSe.SelectRowOnTbl("CunetasGeneral", Data(0).ToString())
        RNG.Worksheet.Activate()

        ' Crear una nueva instancia de CunetasHandleDataItem y asignar datos de Excel
        Dim Cuneta As New CunetasHandleDataItem(listObject:=ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral"))

        ' Obtener todas las propiedades de CunetasHandleDataItem
        Dim props = Cuneta.GetType().GetProperties()

        Dim i As Integer

        Try
            ' Asignar valores a las propiedades del objeto Cuneta usando datos de Excel
            For Each prop In props
                If iHeaders.Contains(prop.Name) And Not EXheaders.Contains(prop.Name) Then
                    prop.SetValue(Cuneta, Convert.ChangeType(Data(i), prop.PropertyType))
                    i += 1
                End If
            Next
        Catch ex As Exception
            ' Capturar y manejar cualquier excepción que ocurra
            Debug.WriteLine("Error al asignar propiedades: " & ex.Message)
        End Try

        ' Agregar el objeto Cuneta al diccionario de handles
        DicHandles.Add(Cuneta.Handle, Cuneta)

        ' Buscar entidades en AutoCAD basadas en la longitud usando el diccionario de handles
        AutoCADHelper2.LookByLength(DicHandles, DicHandlesreSultado)

        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        If DicHandlesreSultado.Count <> 0 Then
            ' Si se encontraron resultados, obtener el primer Cuneta y hacer zoom en la entidad
            Cuneta = DicHandlesreSultado.Values(0)
            AcadZoomManager.SelectedZoom(Cuneta?.Handle, doc)
        Else
            ' Si no se encuentran resultados, intentar usar la estación de inicio para otra operación
            Dim Station As Double = Cuneta.StartStation

            ' Código comentado para verificar alineaciones
            ' Dim AligmentsCertify As New AlignmentHelper()
            ' Dim Al As Alignment = AligmentsCertify.CheckCunetasAlignment(
            '     CType(CLHandle.GetEntityByStrHandle(Cuneta.Handle), Autodesk.AutoCAD.DatabaseServices.Polyline))
            ' CStationOffsetLabel.GotoStation(Station, CAcadHelp.Alignment, doc)
        End If
    End Sub

    Public Sub SelectedByentitySoporte(Optional TabCtrolView As TabControl = Nothing)

        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabCtrolView)

        Dim iHeaders As List(Of String) = Headers.Select(Function(item) item.Item1).ToList()

        ' Construct the dictionary only once
        Dim ListType As List(Of List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId)) = AutoCADHelper2.ConstructDiccionario(cListType)

        If EncabezadosSoporte Is Nothing Then
            Dim ExRNGSe As New ExcelRangeSelector()
            ' Select the range and search for headers
            Dim RNGS As Excel.Range = ExRNGSe.SelectRange()
            Using optimizer As New ExcelOptimizer(ExRNGSe.ExcelApp.GetWorkB().Application)
                Try
                    optimizer.TurnEverythingOff()

                    ' Loop through the Excel range
                    For Each Rng As Range In RNGS
                        Try
                            ' Create a new CunetasHandleDataItem and populate properties
                            Dim Cuneta As New CunetasHandleDataItem

                            SetProFromSoporteCuneta(Cuneta, Rng)

                            ' Handle observations
                            Dim RNGOBSERVACION As Range = RNGOffset(Rng, "OBSERVACIONES TYPSA")

                            ' Sync the data and update the DataGridView
                            SyncSoporteTabla(Cuneta, Cuneta.IDNum, doc, ExRNGSe, RNGOBSERVACION, ListType)


                            ' Only add to DGView if there's no observation value
                            If RNGOBSERVACION.Value = "" Then
                                Cuneta.AddToDGView(DGView)
                            End If
                        Catch ex As Exception
                            ' Log or handle errors specific to this row
                            Debug.WriteLine($"Error processing row {Rng.Address}: {ex.Message}")
                        End Try
                    Next
                    ' Explicitly release the Excel COM objects
                    Marshal.ReleaseComObject(RNGS)
                Catch ex As Exception
                    ' Handle overall errors here
                    Debug.WriteLine($"Error in SelectedByentitySoporte: {ex.Message}")
                Finally
                    optimizer.TurnEverythingOn()
                    ExRNGSe = Nothing
                End Try
            End Using
        End If

    End Sub
    Public Sub MatchSoportevsTable(Optional TabCtrolView As TabControl = Nothing)

        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabCtrolView)

        Dim iHeaders As List(Of String) = Headers.Select(Function(item) item.Item1).ToList()

        ' Construct the dictionary only once
        Dim ListType As List(Of List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId)) = AutoCADHelper2.ConstructDiccionario(cListType)

        If EncabezadosSoporte Is Nothing Then
            Dim ExRNGSe As New ExcelRangeSelector()
            ' Select the range and search for headers
            Dim RNGS As Excel.Range = ExRNGSe.SelectRange()
            Using optimizer As New ExcelOptimizer(ExRNGSe.ExcelApp.GetWorkB().Application)
                Try
                    optimizer.TurnEverythingOff()

                    ' Loop through the Excel range
                    For Each Rng As Range In RNGS
                        Try
                            ' Create a new CunetasHandleDataItem and populate properties
                            Dim Cuneta As New CunetasHandleDataItem

                            SetProFromSoporteCuneta(Cuneta, Rng)
                            Cuneta.SetAccesoNum()
                            ' Handle observations
                            Dim TB As Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")
                            Dim RngTable As Range

                            With Cuneta
                                RngTable = GetRNGByLeght(TB, .Accesos, .StartStation, .EndStation, .Longitud, .Side, .IDNum)
                            End With

                            If RngTable Is Nothing Then
                                Continue For
                            Else

                                Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"IDNum", "Comentarios"}, TB.HeaderRowRange)

                                RNGOffset(Rng, "ID").Formula = "='" & RngTable.Worksheet.Name & "'!" & RngTable.Cells(indexHeader(0)).address(False, False)

                                RNGOffset(Rng, "OBSERVACIONES TYPSA").Formula = "='" & RngTable.Worksheet.Name & "'!" & RngTable.Cells(indexHeader(1)).address(False, False)
                            End If

                            'Dim RNGOBSERVACION As Range = RNGOffset(Rng, "OBSERVACIONES TYPSA")

                        Catch ex As Exception
                            ' Log or handle errors specific to this row
                            Debug.WriteLine($"Error processing row {Rng.Address}: {ex.Message}")
                        End Try
                    Next
                    ' Explicitly release the Excel COM objects
                    Marshal.ReleaseComObject(RNGS)
                Catch ex As Exception
                    ' Handle overall errors here
                    Debug.WriteLine($"Error in SelectedByentitySoporte: {ex.Message}")
                Finally
                    optimizer.TurnEverythingOn()
                    ExRNGSe = Nothing
                End Try
            End Using
        End If

    End Sub

    Public Sub SyncSoportevsTableDGView(Optional TabCtrolView As TabControl = Nothing)

        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

        Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabCtrolView)

        Dim iHeaders As List(Of String) = Headers.Select(Function(item) item.Item1).ToList()

        ' Construct the dictionary only once
        Dim ListType As List(Of List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId)) = AutoCADHelper2.ConstructDiccionario(cListType)

        If EncabezadosSoporte Is Nothing Then
            Dim ExRNGSe As New ExcelRangeSelector()
            ' Select the range and search for headers
            Dim RNGS As Excel.Range = ExRNGSe.SelectRange()
            Using optimizer As New ExcelOptimizer(ExRNGSe.ExcelApp.GetWorkB().Application)
                Try
                    optimizer.TurnEverythingOff()

                    ' Loop through the Excel range
                    For Each Rng As Range In RNGS
                        Try
                            ' Create a new CunetasHandleDataItem and populate properties
                            Dim Cuneta As New CunetasHandleDataItem

                            SetProFromSoporteCuneta(Cuneta, Rng)
                            Cuneta.SetAccesoNum()
                            ' Handle observations
                            Dim TB As Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")
                            Dim RngTable As Range

                            With Cuneta
                                RngTable = GetRNGByLeght(TB, .Accesos, .StartStation, .EndStation, .Longitud, .Side, .IDNum)
                            End With

                            If RngTable Is Nothing Then
                                Continue For
                            Else

                                Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"IDNum", "Comentarios"}, TB.HeaderRowRange)

                                RNGOffset(Rng, "ID").Formula = "='" & RngTable.Worksheet.Name & "'!" & RngTable.Cells(indexHeader(0)).address(False, False)

                                RNGOffset(Rng, "OBSERVACIONES TYPSA").Formula = "='" & RngTable.Worksheet.Name & "'!" & RngTable.Cells(indexHeader(1)).address(False, False)
                            End If

                            'Dim RNGOBSERVACION As Range = RNGOffset(Rng, "OBSERVACIONES TYPSA")

                        Catch ex As Exception
                            ' Log or handle errors specific to this row
                            Debug.WriteLine($"Error processing row {Rng.Address}: {ex.Message}")
                        End Try
                    Next
                    ' Explicitly release the Excel COM objects
                    Marshal.ReleaseComObject(RNGS)
                Catch ex As Exception
                    ' Handle overall errors here
                    Debug.WriteLine($"Error in SelectedByentitySoporte: {ex.Message}")
                Finally
                    optimizer.TurnEverythingOn()
                    ExRNGSe = Nothing
                End Try
            End Using
        End If

    End Sub

    Public Sub SetProFromSoporteCuneta(ByRef Cuneta As CunetasHandleDataItem, Rng As Range)

        Dim STATIONS As New List(Of Double)

        For Each HD As String In {"KM INICIAL", "KM FINAL"}
            Dim celdaValor As Object = RNGOffset(Rng, HD).Value
            If TypeOf celdaValor Is Double Then
                STATIONS.Add(RNGOffset(Rng, HD).Value)

            ElseIf TypeOf celdaValor Is String Then

                STATIONS.Add(CDbl(RNGOffset(Rng, HD).Value.replace("+", "").replace(",", ".")))
            End If
        Next

        'Dim STATIONS As New List(Of Double) From {CDbl(RNGOffset(Rng, "KM INICIAL").Value.replace("+", "").replace(",", ".")), CDbl(RNGOffset(Rng, "KM FINAL").Value.replace("+", "").replace(",", "."))}

        Cuneta.StartStation = STATIONS.Min

        Cuneta.EndStation = STATIONS.Max

        Dim RNGFound As Range = RNGOffset(Rng, "LADO")

        Select Case RNGFound?.Value.ToString()
            Case "DER"
                Cuneta.Side = "Right"', "Left")
            Case "IZQ"
                Cuneta.Side = "Left"
            Case Else
                Cuneta.Side = ""
        End Select

        Cuneta.Longitud = Math.Abs(CDbl(RNGOffset(Rng, "LONG.(m)").Value))

        Cuneta.IDNum = CDbl(Replace(RNGOffset(Rng, "ID").Value, "CU-", ""))
    End Sub
    Public Function RNGOffset(RNG As Range, Header As String) As Range

        Dim RNGFound As Range = RNG?.Worksheet.Cells.Find(What:=Header, MatchCase:=True, LookAt:=XlLookAt.xlWhole)
        If RNGFound IsNot Nothing Then
            Return RNG.Offset(0, RNGFound?.Column - RNG?.Column)
        Else
            Return Nothing
        End If
    End Function
    Public Sub SyncSoporteTabla(Cuneta As CunetasHandleDataItem, Codnum As Double, doc As Document, ExRNGSe As ExcelRangeSelector, ByRef RNGnotM As Range, Optional ByRef ListType As List(Of List(Of Autodesk.AutoCAD.DatabaseServices.ObjectId)) = Nothing)
        Dim DicHandles As New Dictionary(Of String, CunetasHandleDataItem)

        Dim DicHandlesreSultado As New Dictionary(Of String, CunetasHandleDataItem)

        DicHandles.Add("Not", Cuneta)
        'buscart primero por numero en la tabla 


        AutoCADHelper2.LookByLength(DicHandles, DicHandlesreSultado, cListType:=ListType)


        If DicHandlesreSultado.Count <> 0 Then
            Cuneta = DicHandlesreSultado.Values(0)
            Cuneta.SetPropertiesFromDWG(Cuneta.Handle, Cuneta.Alignment.Handle.ToString())

            Cuneta.IDNum = Codnum

            If Not DicHandlesExistentes.ContainsKey(Cuneta.Handle) Then DicHandlesExistentes.Add(Cuneta.Handle, Cuneta)

            UpDateByExcelCuneta(Cuneta)

            'RNGnotM.Value = "" ' "=IFNA(OFFSET($AC$8,MATCH(A81,$Q$9:$Q$121,0),0),"""")"
        Else
            Dim TB As Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")

            With Cuneta

                .SetAccesoNum()

                .IDNum = Codnum

                RNG = GetRNGByLeght(TB, .Accesos, .StartStation, .EndStation, .Longitud, .Side, .IDNum)

                If RNG Is Nothing Then RNG = GetRNGByHandle(.Handle, TB)
            End With

            If RNG Is Nothing Then RNGnotM.Value = "No se encontro"
        End If
    End Sub
    'crear una funcion para dividir las polilineas y agregar una de esta a el diccionario DicHandlesFueradeProyecto
    'EventsHandler.Add(New EventHandlerClass(DividirL, "click", Sub() CtrolsButtonsUT.DividePL(CAcadHelp.Alignment, TextBox3)))

    Public Sub CLoadData(CunetasExistentesDGView As DataGridView, DGViewByLeght As DataGridView, CunetasExternasDGView As DataGridView, CAcadHelp As ACAdHelpers)
        If DicHandlesExistentes.Count > 0 Or DicHandlesNoExistentes.Count > 0 Then
            AddInfo(CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView, CAcadHelp)
        Else
            LoadData(CAcadHelp.ThisDrawing.Name, "CunetasGeneral", FilePath)
            AddInfo(CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView, CAcadHelp)
        End If
    End Sub
    ' Method to load data into the HandleData dictionary
    ' The hndleData represents data processed by AutoCADHelper.ProcesarHandlesDesdeDoc(docName, tableName)
    ' Método para cargar datos en el diccionario HandleData
    Public Sub LoadData(docName As String, tableName As String, Optional WBFilePath As String = vbNullString)
        Try
            ' Obtener el diccionario con los datos de TabName

            'Obtener la carpeta del archivo WBFilePath
            Dim Carpeta As String = Left(WBFilePath, InStrRev(WBFilePath, "\"))

            'Crear una segunda clase para el manejo de base de datos 
            'Dim SQLmase As New DataBSQLManager(Carpeta & "\" & "BasededatosCunetas")
            HandleData.Clear()
            ClearData()
            Dim hndleData As Dictionary(Of String, CunetasHandleDataItem) = AutoCADHelper2.ProcesarHandlesDesdeDoc(docName, tableName, WBFilePath)

            ' Agregar los datos al diccionario principal si no están ya presentes
            For Each kvp As KeyValuePair(Of String, CunetasHandleDataItem) In hndleData
                If Not HandleData.ContainsKey(kvp.Key) Then
                    HandleData.Add(kvp.Key, kvp.Value)
                    'SQLmase.InsertHandleData(kvp.Value)    
                End If
            Next
            AutoCADHelper2.VerificarHandles(HandleData, DicHandlesExistentes, DicHandlesNoExistentes)
            AutoCADHelper2.LookByLength(DicHandlesNoExistentes, DicHandlesByLength, True)
        Catch ex As Exception
            Console.WriteLine("Error al cargar los datos: " & ex.Message)
        End Try
    End Sub
    Public Sub ByLeght(ByRef DGView1 As DataGridView, ByRef DGView2 As DataGridView, ByRef DGView3 As DataGridView, ByRef CAcadHelp As ACAdHelpers)

        'ClearData()
        'AutoCADHelper2.VerificarHandles(HandleData, DicHandlesExistentes, DicHandlesNoExistentes)
        AutoCADHelper2.LookByLength(DicHandlesNoExistentes, DicHandlesByLength, True)

        AddInfo(DGView1, DGView2, DGView3, CAcadHelp)

    End Sub

    ' Method to get data for a specific handle
    ' Método para obtener datos de un handle específico
    Public Function GetHandleData(handle As String) As CunetasHandleDataItem
        Dim dataItem As CunetasHandleDataItem = Nothing
        If HandleData.TryGetValue(handle, dataItem) Then
            Return dataItem
        Else
            Console.WriteLine("Handle no encontrado")
            Return Nothing
        End If
    End Function

    ' Method to transfer a handle from one dictionary to another (without removing from sourceDict)
    ' Método para transferir un handle de un diccionario a otro sin eliminarlo del diccionario fuente
    Public Sub TransferHandle(sourceDict As Dictionary(Of String, CunetasHandleDataItem), targetDict As Dictionary(Of String, CunetasHandleDataItem), handle As String)
        TransferOrMoveHandle(sourceDict, targetDict, handle, False)
    End Sub

    ' Method to move a handle from one dictionary to another (removes from sourceDict)
    ' Método para mover un handle de un diccionario a otro, eliminándolo del diccionario fuente
    Public Sub MoveHandle(sourceDict As Dictionary(Of String, CunetasHandleDataItem), targetDict As Dictionary(Of String, CunetasHandleDataItem), handle As String)
        TransferOrMoveHandle(sourceDict, targetDict, handle, True)
    End Sub

    ' Método privado para transferir o mover un handle entre diccionarios
    Private Sub TransferOrMoveHandle(sourceDict As Dictionary(Of String, CunetasHandleDataItem), targetDict As Dictionary(Of String, CunetasHandleDataItem), handle As String, Optional removeFromSource As Boolean = False)
        Dim sourceItem As CunetasHandleDataItem = Nothing
        If sourceDict.TryGetValue(handle, sourceItem) Then
            targetDict(handle) = sourceItem
            If removeFromSource Then sourceDict.Remove(handle)
        Else
            Console.WriteLine("TabName no encontrado en el diccionario fuente")
        End If
    End Sub

    ' Method to clear all dictionaries (useful for resetting data)
    ' Método para limpiar todos los diccionarios
    Public Sub ClearData()
        'HandleData.Clear()
        DicHandlesExistentes.Clear()
        DicHandlesByLength.Clear()
        DicHandlesExternos.Clear()
    End Sub
    Public Sub AddInfo(DGView1 As DataGridView, DGView2 As DataGridView, DGView3 As DataGridView, CAcadHelp As ACAdHelpers)

        'con esta linea quiero simplificar el proceso en un solo bucle
        ' Inicializar las listas para simplificar el proceso en un solo bucle

        Dim DicHandles() As (Dictionary(Of String, CunetasHandleDataItem), DataGridView) = {
                                                                                            (DicHandlesExistentes, DGView1),
                                                                                            (DicHandlesByLength, DGView2),
                                                                                            (DicHandlesNoExistentes, DGView3)
                                                                                                }

        UpdateDataDGV.UpDateDGVByHandleDataProcessor(DicHandles, CAcadHelp, Me)
    End Sub
    'metodos para agregar datos a tabla en excel 
    Public Sub UpDateExcel(Optional TabCtrol As TabControl = Nothing, Optional Handle As String = "")
        Using excelOptimizer As New ExcelOptimizer()
            excelOptimizer.TurnEverythingOff()

            Dim ExcelFilters As New ExcelFilters
            Dim ExRNGSe As New ExcelRangeSelector()

            Dim Table As ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")

            ExcelFilters.GuardarFiltros(Table)

            If String.IsNullOrEmpty(Handle) AndAlso TabCtrol IsNot Nothing Then
                'Devuelve el datagridview que pertenece al tabcontrol
                Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabCtrol) 'TabCtrol)



                Dim selectedCells As DataGridViewSelectedCellCollection = DGView.SelectedCells

                For Each cell As DataGridViewCell In selectedCells
                    Handle = DGView.Rows(cell.RowIndex).Cells("Handle").Value.ToString()
                    UpDateByhandleFromCDI(Handle)
                Next
            Else
                UpDateByhandleFromCDI(Handle)
            End If
            ExcelFilters.AplicarFiltros(Table)
            excelOptimizer.TurnEverythingOn()
            excelOptimizer.Dispose()
        End Using
    End Sub
    'metodo para sustituir un dato por otro en la tabla de excel 
    Public Sub ReplaceDataExcel(Optional TabCtrol As TabControl = Nothing, Optional Handle As String = "")
        If String.IsNullOrEmpty(Handle) AndAlso TabCtrol IsNot Nothing Then
            'Devuelve el datagridview que pertenece al tabcontrol
            Dim DGView As DataGridView = GetControlsDGView.GetDGView(TabCtrol) 'TabCtrol)

            Dim selectedCells As DataGridViewSelectedCellCollection = DGView.SelectedCells

            For Each cell As DataGridViewCell In selectedCells
                Handle = DGView.Rows(cell.RowIndex).Cells("Handle").Value.ToString()
                UpDateByhandleFromCDI(Handle)
            Next
        Else
            'UpDateByhandleFromCDI(Handle)
        End If
    End Sub

    Public Sub UpDateByhandleFromCDI(Handle As String, Optional RNGHandle As Range = Nothing)

        Dim ExRNGSe As New ExcelRangeSelector()
        Dim TB As Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")
        If TB Is Nothing Then Exit Sub
        Dim CUNETA As New CunetasHandleDataItem(listObject:=TB)

        If DicHandlesExistentes.ContainsKey(Handle) Then
            'Dim AligmentsCertify As New AlignmentHelper()

            'AligmentsCertify.CheckAlimentByProximida(CLHandle.GetEntityByStrHandle(Handle))
            'AligmentsCertify.CheckByStations()
            'CUNETA.AlignmentHDI = CUNETA.GetFromExcel("AlignmentHDI")

            'CUNETA.SetPropertiesFromDWG(Handle, AligmentsCertify.Alignment?.Handle.ToString())
            CUNETA.Handle = Handle

            CUNETA.SetPropertiesFromDWG(Handle, CUNETA.AlignmentHDI)

            If DicHandlesExistentes.ContainsKey(Handle) Then
                DicHandlesExistentes.Remove(Handle)
                DicHandlesExistentes.Add(CUNETA.Handle, CUNETA)
            End If
        Else
            Dim AligmentsCertify As New AlignmentHelper()

            AligmentsCertify.CheckAlimentByProximida(CLHandle.GetEntityByStrHandle(Handle))

            If AligmentsCertify.Alignment Is Nothing Then
                'AligmentsCertify.CheckAlignmentWithStationEquation(CAcadHelp, CLHandle.GetEntityByStrHandle(Handle))
                CUNETA.Handle = Handle

                CUNETA.SetPropertiesFromDWG(Handle, CUNETA.AlignmentHDI)
            Else
                'CUNETA.AlignmentHDI = CUNETA.GetFromExcel("AlignmentHDI")

                CUNETA.SetPropertiesFromDWG(Handle, AligmentsCertify.Alignment?.Handle.ToString())

                DicHandlesExistentes.Add(CUNETA.Handle, CUNETA)
            End If

        End If

        'Dim optimizer As New ExcelOptimizer(ExRNGSe.ExcelApp.GetWorkB().Application)
        'optimizer.TurnEverythingOff()

        'proceso para añadir datos a excel 
        Dim Exp As New List(Of String) From {"Comentarios", "Ubicacion", "Medicion de Pago", "IDNum"}
        Dim RNG As Microsoft.Office.Interop.Excel.Range
        If RNGHandle IsNot Nothing Then
            RNG = RNGHandle
        Else
            RNG = GetRNGByHandle(CUNETA.Handle, TB)
        End If


        With CUNETA
            If RNG Is Nothing Then RNG = GetRNGByLeght(TB, .Accesos, .StartStation, .EndStation, .Longitud, .Side)
        End With


        Dim Formatos As List(Of String) = Headers.Select(Function(item) item.Item3).ToList()

        'Formatos.Insert(0, """Acceso ""00") ' Formatos.Insert(0, $"Acceso {0.ToString("00")}")

        CUNETA.AddToExcelTable(TB, Exp, RNG, Formatos)
        'TB.AutoFilter.ApplyFilter()

        'optimizer.TurnEverythingOn()
        'optimizer.Dispose()
    End Sub
    Public Sub UpDateByExcelCuneta(CUNETA As CunetasHandleDataItem, Optional RECALCULATE As Boolean = False)

        Dim ExRNGSe As New ExcelRangeSelector()

        Dim TB As Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")

        If RECALCULATE Then
            Dim AligmentsCertify As New AlignmentHelper()

            AligmentsCertify.CheckAlimentByProximida(CLHandle.GetEntityByStrHandle(CUNETA.Handle))

            CUNETA.SetPropertiesFromDWG(CUNETA.Handle, AligmentsCertify.Alignment?.Handle.ToString())
        End If

        Dim optimizer As New ExcelOptimizer(ExRNGSe.ExcelApp.GetWorkB().Application)

        optimizer.TurnEverythingOff()

        'proceso para añadir datos a excel 
        Dim Exp As New List(Of String) From {"Comentarios", "Ubicacion", "Medicion de Pago"}

        'CUNETA.Accesos = 9
        Dim RNG As Microsoft.Office.Interop.Excel.Range

        With CUNETA
            RNG = GetRNGByLeght(TB, .Accesos, .StartStation, .EndStation, .Longitud, .Side)
            If RNG Is Nothing Then RNG = GetRNGByHandle(CUNETA.Handle, TB)
        End With

        Dim Formatos As List(Of String) = Headers.Select(Function(item) item.Item3).ToList()

        CUNETA.AddToExcelTable(TB, Exp, RNG, Formatos)

        optimizer.TurnEverythingOn()
        optimizer.Dispose()
    End Sub
    Function GetRNGByIDNUM(ID As Integer, Acceso As Integer, Table As Microsoft.Office.Interop.Excel.ListObject) As Microsoft.Office.Interop.Excel.Range

        'buscar a partir de un Handle en la cloumna Handle
        Dim ExRNGSe As New ExcelRangeSelector()

        Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"IDNum", "Accesos"}, Table.HeaderRowRange)

        Dim IDColumnIndex As Integer = indexHeader(0)

        Dim AccesoColumnIndex As Integer = indexHeader(1)

        Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing

        For Each row As Microsoft.Office.Interop.Excel.ListRow In Table.ListRows
            If row.Range.Cells(1, IDColumnIndex).Value = ID AndAlso row.Range.Cells(1, AccesoColumnIndex).Value = Acceso Then
                RNG = row.Range ' Return the matching row if found
                Exit For
            End If
        Next
        Return RNG
    End Function
    Function GetRNGByHandle(handle As String, Table As Microsoft.Office.Interop.Excel.ListObject) As Microsoft.Office.Interop.Excel.Range

        'buscar a partir de un Handle en la cloumna Handle
        Dim ExRNGSe As New ExcelRangeSelector()

        Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"Handle", "Comentarios"}, Table.HeaderRowRange)

        Dim handleColumnIndex As Integer = indexHeader(0)

        Dim ComentariosColumnIndex As Integer = indexHeader(1)

        Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing

        For Each row As Microsoft.Office.Interop.Excel.ListRow In Table.ListRows
            Try
                If row.Range.Cells(1, handleColumnIndex).Value.ToString() = handle Then
                    RNG = row.Range ' Return the matching row if found
                    Exit For
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        Next
        Return RNG
    End Function

    Function GetRNGByLeght(Table As Microsoft.Office.Interop.Excel.ListObject,
                           Acceso As Integer,
                           Startstation As Double,
                           Endstation As Double,
                           Longitud As Double,
                           Side As String, Optional ByRef IDNum As Integer = 0) As Microsoft.Office.Interop.Excel.Range

        'buscar a partir de un Handle en la cloumna Handle
        Dim ExRNGSe As New ExcelRangeSelector()
        Dim Tolerancia As Double = 2
        Dim Presicion As Integer = 2


        Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"Accesos", "StartStation", "EndStation", "Longitud", "Side", "IDNum"}, Table.HeaderRowRange)

        Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing

        For Each row As Microsoft.Office.Interop.Excel.ListRow In Table.ListRows
            If row.Range.Cells(1, indexHeader(0)).Value = Acceso AndAlso
              Math.Abs(Math.Round(CDbl(row.Range.Cells(1, indexHeader(1)).Value), Presicion) - Startstation) <= Tolerancia AndAlso
              Math.Abs(Math.Round(CDbl(row.Range.Cells(1, indexHeader(2)).Value), Presicion) - Endstation) <= Tolerancia AndAlso
               Math.Abs(Math.Round(CDbl(row.Range.Cells(1, indexHeader(3)).Value), Presicion) - Longitud) <= Tolerancia Then

                Dim EvaNum As Boolean = True
                Dim EvaSide As Boolean = True

                If IDNum <> 0 Then EvaNum = IDNum = row.Range.Cells(1, indexHeader(5)).Value

                If Side <> "" Then EvaSide = row.Range.Cells(1, indexHeader(4)).Value.ToString() = Side

                If EvaNum OrElse EvaSide Then
                    RNG = row.Range ' Return the matching row if found
                    Exit For
                End If
            End If
        Next
        Return RNG
    End Function
End Class

'Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
'    If handleProcessor.DicHandlesExistentes.Count > 0 Or handleProcessor.DicHandlesNoExistentes.Count > 0 Then
'        handleProcessor.AddInfo(CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView, CAcadHelp)
'    Else
'        handleProcessor.LoadData(CAcadHelp.ThisDrawing.Name, "CunetasGeneral", WBFilePath)
'        handleProcessor.AddInfo(CunetasExistentesDGView, DGViewByLeght, CunetasExternasDGView, CAcadHelp)
'    End If
'End Sub


