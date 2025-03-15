' Clase para representar los datos del handle
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.GraphicsSystem
Imports Autodesk.AutoCAD.Windows.Data
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data.Entity.Core.Common.EntitySql
Imports System.Diagnostics
Imports System.Linq
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox

Public Class CunetasHandleDataItem
    'base de datos para esta clase
    'Public Property DbManager As New DataBSQLManager()

    ' Evento que se disparará cuando la propiedad Name cambie
    Public Event EndStationChanged As EventHandler
    Public Event HandleChanged As EventHandler

    ' Columnas visibles en el DataGridView o tabla
    Public Property Accesos As Double

    Public CatchTableName As String
    Public Property TableName As String
        Get
            Return CatchTableName
        End Get
        Set(value As String)
            If CatchHandle <> value Then
                CatchTableName = value
                'SettingBaseDDatos()
            End If
        End Set
    End Property


    Public Property CatchHandle As String
    ''' <summary>
    ''' Propiedad para manejar cambios en el Handle.
    ''' Incluye lógica para manejar conversiones y actualizaciones.
    ''' </summary>
    Public Property Handle As String

        Get
            Return CatchHandle
        End Get
        Set(value As String)
            ' Trigger the event only if the value changes
            If CatchHandle <> value Then
                CatchHandle = value
                ' Trigger the EndStationChanged event
                RaiseEvent HandleChanged(Me, EventArgs.Empty)
                If CLHandle.CheckIfExistHd(CatchHandle) Then
                    Type = TypeName(CLHandle.GetEntityByStrHandle(CatchHandle))

                    SetAccesoNum()

                    PLANO = PlanoNombre(CLHandle.GetEntityIdByStrHandle(CatchHandle))

                    If Type = "Polyline" Then
                        Polyline = CLHandle.GetEntityByStrHandle(CatchHandle)
                    ElseIf Type = "FeatureLine" OrElse Type = "Polyline3d" OrElse Type = "Polyline2d" OrElse Type = "line" Then
                        ConvertToPolyline(False, Polyline:=Polyline)
                        CatchHandle = Polyline?.Handle.ToString()
                        Type = TypeName(CLHandle.GetEntityByStrHandle(CatchHandle))
                        PLANO = PlanoNombre(CLHandle.GetEntityIdByStrHandle(CatchHandle))
                    ElseIf Type = "Parcel" Then
                        Dim ParcelCommands As New ParcelCommands
                        Polyline = ParcelCommands.GetParcelBoudary(CatchHandle, IDNum)
                        CatchHandle = Polyline?.Handle.ToString()

                        Type = TypeName(CLHandle.GetEntityByStrHandle(CatchHandle))

                        PLANO = PlanoNombre(CLHandle.GetEntityIdByStrHandle(CatchHandle))
                    Else

                        GoTo HandleErroneo
                    End If

                    Dim modifier As New PolylineModifier()

                    'modifier.ChangePolylineProperties(CatchHandle, "FLUXO3", 4.5, 0.5, Nothing)

                    Dim AligmentsCertify As New AlignmentHelper()

                    If EndStation <> 0 AndAlso StartStation <> 0 Then
                        AligmentsCertify.CheckByStations(CType(Polyline, Autodesk.AutoCAD.DatabaseServices.Entity), Side, StartStation, EndStation, 2)
                    Else
                        AligmentsCertify.CheckAlimentByProximida(Polyline, False)
                    End If

                    Alignment = AligmentsCertify.Alignment

                    If Alignment IsNot Nothing Then
                        AlignmentHDI = Alignment?.Handle.ToString()
                    Else
                        Alignment = AligmentsCertify.CheckCunetasAlignment(Polyline)
                        AlignmentHDI = Alignment?.Handle.ToString()
                    End If
                End If
HandleErroneo:
            End If
        End Set
    End Property

    Public Sub SetAccesoNum()
        Dim docName As String = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name
        ' Si no se encuentra el código, salir
        Dim RefAcceso As String = CodigoExtractor.AccesoNum(docName)
        If String.IsNullOrEmpty(RefAcceso) Then
            Console.WriteLine("El archivo no presenta referencia de accesos.")
            Exit Sub
        End If

        ' Obtener el número de acceso a partir de la referencia extraída
        Accesos = CDbl(CodigoExtractor.ExtraerCodigo(RefAcceso, "\d+"))
        FilePath = docName


        FileName = Right(docName, Len(docName) - InStrRev(docName, "\"))
    End Sub

    ' Enhanced property implementation for updating entity layers intelligently
    Public Event LayerChanged As EventHandler

    Private _catchLayer As String

    ' Property to get or set the layer name and trigger updates accordingly
    Public Property Layer As String
        Get
            Return _catchLayer
        End Get
        Set(value As String)
            ' Trigger the event and update only if the layer value has actually changed
            If _catchLayer <> value Then
                _catchLayer = value

                ' Raise the LayerChanged event safely, checking for null subscribers
                RaiseEvent LayerChanged(Me, EventArgs.Empty)

                ' Check if the entity handle is valid and the layer needs to be updated
                If Not String.IsNullOrEmpty(Handle) AndAlso CLHandle.CheckIfExistHd(Handle) Then
                    Dim currentLayer As String = CLayerHelpers.GetLayersAcEnt(Handle)

                    ' Update the entity layer only if it's different from the desired layer
                    If currentLayer <> _catchLayer Then
                        CLayerHelpers.ChangeLayersAcEnt(Handle, _catchLayer)
                        ' Si no existe en el diccionario, buscar en la tabla de Excel
                        'Dim ExRNGSe As New ExcelRangeSelector()
                        'Dim RNG As Microsoft.Office.Interop.Excel.Range = Nothing
                        'RNG = ExRNGSe.HandleFromExcelTable(TBName:="CunetasGeneral", handle:=Handle, "Layer") ' "Comentarios")
                        'If RNG IsNot Nothing Then RNG.Value = _catchLayer
                    End If

                End If
                'SyncWithDatabase()
            End If
        End Set
    End Property

    Public Property MinX As Double
    Public Property MinY As Double
    Public Property MaxX As Double
    Public Property MaxY As Double
    Public Property StartStation As Double

    Private _endStation As Double

    ' Public property EndStation with validation logic
    Public Property EndStation As Double
        Get
            Return _endStation
        End Get
        Set(value As Double)
            ' Trigger the event only if the value changes
            If _endStation <> value Then
                _endStation = value

                ' Trigger the EndStationChanged event
                RaiseEvent EndStationChanged(Me, EventArgs.Empty)

                ' Check if StartStation > EndStation and perform Excel operations
                If StartStation > _endStation Then
                    CStationOffsetLabel.GetMxMnOPPL(StartStation, _endStation)
                    SyncPropExcel({"StartStation", "EndStation"}, "Get")
                End If
            End If
        End Set
    End Property

    Public Property Longitud As Double
    Public Property Area As Double
    Public Property Side As String
    Public Property CatchAlignmentHDI As String

    ''' <summary>
    ''' Propiedad para manejar el AlignmentHDI y verificar alineaciones relacionadas.
    ''' </summary>
    Public Property AlignmentHDI As String
        Get
            Return CatchAlignmentHDI
        End Get
        Set(value As String)
            If CatchAlignmentHDI <> value Then
                CatchAlignmentHDI = value
                EnsureAlignmentIsRight(value:=value)
            End If
        End Set
    End Property

    ''' <summary>
    ''' Asegura que el alineamiento del handle actual sea correcto y, si no lo es,
    ''' intenta obtener y asignar el alineamiento adecuado desde Excel o mediante validación.
    ''' </summary>
    ''' <param name="value">El valor esperado para el AlignmentHDI.</param>
    Private Sub EnsureAlignmentIsRight(value As String)
        ' Crear una nueva instancia de AlignmentHelper para manejar validaciones de alineamientos
        Dim AligmentsCertify As New AlignmentHelper()

        'Verificar si el handle existe en AutoCAD y realizar una certificación de proximidad
        If CLHandle.CheckIfExistHd(Handle) Then
            AligmentsCertify.AlimentByProximida(Handle, False)
        End If

        ' Verificar si se encontró un alineamiento válido y si es diferente del actual
        If AligmentsCertify.Alignment IsNot Nothing AndAlso AlignmentHDI <> AligmentsCertify.Alignment?.Handle.ToString() Then
            ' Asignar el handle del nuevo alineamiento
            CatchAlignmentHDI = AligmentsCertify.Alignment.Handle.ToString()
        Else
            ' Si el valor actual de CatchAlignmentHDI es diferente del proporcionado, intentar obtenerlo desde Excel
            If CatchAlignmentHDI <> value Then
                CatchAlignmentHDI = GetFromExcel("AlignmentHDI")
            End If

            ' Si aún no se ha asignado un alineamiento, buscar datos adicionales en la tabla de Excel
            If String.IsNullOrEmpty(CatchAlignmentHDI) Then
                ' Crear un selector de rangos de Excel y definir los encabezados de las columnas relevantes
                Dim ExRNGSe As New ExcelRangeSelector()
                Dim Headers As New List(Of String) From {"Handle", "AlignmentHDI", "StartStation", "StartStation", "Longitud", "Side"}

                ' Seleccionar una fila de la tabla "CunetasGeneral" usando los encabezados definidos
                Dim Data As List(Of Object) = ExRNGSe.SelectRowOnTbl("CunetasGeneral", Headers)
            End If
        End If
        ' Si FilePath está vacío, configurar el número de acceso
        If FilePath = "" Then
            SetAccesoNum()
        End If
        'SyncWithDatabase(DbManager, "CunetasDatos")
    End Sub

    'option de de seleccion de alineamiento para asignacion
    Public Property Comentarios As String
    Public Property Tramos As String
    Public Property PLANO As String
    Public Property FileName As String
    Public Property FilePath As String
    Public Property CodigoCuneta As String

    Public Property CantidadPorAnalisis As New List(Of Integer)
    ' Evento que se disparará cuando la propiedad cambie
    Public Event PosicionChanged As EventHandler
    ' Campo privado que almacena el valor numérico
    Private _posicionValue As Integer
    ' Propiedad pública que recibe un número entero y lo convierte en un string
    Public Property Posicion As Integer
        Get
            Return _posicionValue
        End Get
        Set(ByVal value As Integer)
            ' Establecer el valor del campo privado
            _posicionValue = value

            ' Disparar el evento cuando la posición cambia
            RaiseEvent PosicionChanged(Me, EventArgs.Empty)
        End Set
    End Property

    ' Propiedad calculada que devuelve el valor formateado como string
    Public ReadOnly Property PosicionFormatted As String
        Get
            If _posicionValue = 0 Then
                Return "No asignada"
            ElseIf _posicionValue = 1 Then
                Return "Vía"
            ElseIf _posicionValue >= 2 And _posicionValue < Me.CantidadPorAnalisis.Max Then
                Return "Berma" & (_posicionValue - 1).ToString()
            Else
                Return "Corona"
            End If
        End Get
    End Property
    ''' <summary>
    ''' Representa el tipo de la entidad en AutoCAD, como "Polyline", "FeatureLine", "Polyline3d", etc.
    ''' Esta propiedad se asigna cuando se identifica la entidad por su handle y se convierte a un tipo específico si es necesario.
    ''' </summary>
    Public Property Type As String

    ''' <summary>
    ''' Referencia a un objeto relacionado de tipo CunetasHandleDataItem.
    ''' Se utiliza para enlazar cunetas que están conectadas o relacionadas de alguna manera.
    ''' </summary>
    Public Property RelateCuneta As CunetasHandleDataItem

    ''' <summary>
    ''' Lista de objetos CunetasHandleDataItem que representan tramos relacionados.
    ''' Se utiliza para almacenar y gestionar tramos que están conectados, como cuando el final de un tramo coincide con el inicio de otro.
    ''' </summary>
    Public Property RelatedTramos As List(Of CunetasHandleDataItem) = New List(Of CunetasHandleDataItem)()


    ''' <summary>
    ''' Identifica y asigna tramos relacionados en el conjunto de cunetas filtradas.
    ''' Un tramo es considerado relacionado si el final de uno coincide con el inicio de otro.
    ''' </summary>
    ''' <param name="cunetasFiltradas">Conjunto de cunetas ya filtradas por rango y lado.</param>
    Public Sub IdentificarTramosRelacionadosParaRango(cunetasFiltradas As Dictionary(Of String, CunetasHandleDataItem))
        ' Recorrer el conjunto filtrado para encontrar tramos relacionados
        For Each otherCuneta As CunetasHandleDataItem In cunetasFiltradas.Values
            If otherCuneta IsNot Me Then
                ' Verificar si el final de esta cuneta coincide con el inicio de otro tramo, o viceversa
                If (EndStation = otherCuneta.StartStation) OrElse
                   (StartStation = otherCuneta.EndStation) AndAlso
                   otherCuneta.Layer = Layer Then
                    ' Añadir el otro tramo a la lista de tramos relacionados si aún no está incluido
                    If Not RelatedTramos.Contains(otherCuneta) Then
                        RelatedTramos.Add(otherCuneta)
                    End If
                    ' Asegurar que el otro tramo también tenga esta cuneta en sus tramos relacionados
                    If Not otherCuneta.RelatedTramos.Contains(Me) Then
                        otherCuneta.RelatedTramos.Add(Me)
                    End If
                End If
            End If
        Next
    End Sub


    ''' <summary>
    ''' Objeto Polyline de AutoCAD asociado a este handle.
    ''' Esta propiedad se usa para realizar operaciones geométricas y de modificación sobre la polilínea.
    ''' </summary>
    Public Property Polyline As Autodesk.AutoCAD.DatabaseServices.Polyline

    ''' <summary>
    ''' Rango de datos en Excel asociado a esta instancia.
    ''' Esta propiedad es privada y se utiliza internamente para manejar el rango de la fila de datos en Excel.
    ''' </summary>
    Private Property FilaData As Range

    ''' <summary>
    ''' Rango de la celda en Excel que contiene los comentarios relacionados con este handle.
    ''' Se usa para gestionar y actualizar los comentarios en el archivo Excel vinculado.
    ''' </summary>
    Public Property RNGComentarios As Range

    ''' <summary>
    ''' Rango de la celda en Excel que contiene el handle.
    ''' Esta propiedad es útil para localizar y gestionar el handle en la tabla de Excel.
    ''' </summary>
    Public Property RNGHandle As Range

    ' Headers para las columnas
    Public Shared ReadOnly Property Headers As String() = HandleDataProcessor.Headers.Select(Function(t) t.Item1).ToArray()

    ' Indices de los Headers en la fila de Excel
    Public Property IndexHeaders As Integer()
    Public Alignment As Alignment
    Public Property IDNum As Integer

    Public Property Close As Boolean

    'Parameterless constructor
    Public Sub New()
        ' Initialization code if needed
        TableName = "CunetasDatos"
    End Sub
    ' Constructor para inicializar con los índices de los encabezados o vacío
    Public Sub New(Optional ByRef iIndexHeaders As Integer() = Nothing, Optional listObject As ListObject = Nothing)
        ' Verificar si iIndexHeaders es Nothing o está vacío
        If iIndexHeaders Is Nothing OrElse iIndexHeaders.Length = 0 Then
            ' Si está vacío o es nulo y listObject no es Nothing, calcular los índices desde la tabla de Excel
            If listObject IsNot Nothing Then
                Me.IndexHeaders = UpdateExcelTable.ConArray(Headers, listObject.HeaderRowRange)
            End If
        Else
            ' Si no está vacío, usar los índices proporcionados
            Me.IndexHeaders = iIndexHeaders
        End If
        'DataHandler.InitializeAndInsertData(Of me)()
        TableName = "CunetasDatos"
    End Sub

    'metodo para la conveersion de entidades 
    Public Sub ConvertToPolyline(Optional Sync As Boolean = True, Optional ByRef Polyline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing)
        Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity

        Ent = CLHandle.GetEntityByStrHandle(CatchHandle)
        If TypeOf Ent Is FeatureLine Then
            Dim LY As String = GetFromExcel("Layer")
            If String.IsNullOrEmpty(LY) Then CLayerHelpers.GetCurrentLayer(LY)
            Polyline = FeaturelineToPolyline.ConvertFeaturelineToPolyline(TryCast(Ent, FeatureLine), True)
            CLayerHelpers.ChangeLayersAcEnt(Polyline.Handle.ToString(), LY)
            Handle = Polyline.Handle.ToString()
            If Sync Then SyncPropExcel({"Handle"}, "Get")

        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
            Dim LY As String = GetFromExcel("Layer")
            If String.IsNullOrEmpty(LY) Then CLayerHelpers.GetCurrentLayer(LY)
            Polyline = CPoLy3dToPL(TryCast(Ent, Autodesk.AutoCAD.DatabaseServices.Polyline3d), True)
            CLayerHelpers.ChangeLayersAcEnt(Polyline.Handle.ToString(), LY)
            Handle = Polyline.Handle.ToString()
            If Sync Then SyncPropExcel({"Handle"}, "Get")
        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
            Dim LY As String = GetFromExcel("Layer")
            If String.IsNullOrEmpty(LY) Then CLayerHelpers.GetCurrentLayer(LY)
            Polyline = CPoLy2dToPL(TryCast(Ent, Autodesk.AutoCAD.DatabaseServices.Polyline2d), True)
            CLayerHelpers.ChangeLayersAcEnt(Polyline.Handle.ToString(), LY)
            Handle = Polyline.Handle.ToString()
            If Sync Then SyncPropExcel({"Handle"}, "Get")
        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Line Then
            Dim LY As String = GetFromExcel("Layer")
            If String.IsNullOrEmpty(LY) Then CLayerHelpers.GetCurrentLayer(LY)
            Polyline = ConvertLineToPolyline(TryCast(Ent, Autodesk.AutoCAD.DatabaseServices.Line), True)
            CLayerHelpers.ChangeLayersAcEnt(Polyline.Handle.ToString(), LY)
            Handle = Polyline.Handle.ToString()
            If Sync Then SyncPropExcel({"Handle"}, "Get")
        End If
    End Sub
    'Método para asignar valores a las propiedades usando un arreglo de valores y los headers
    Public Sub SetPropertiesFromTableRng(RNG As Range)
        ' Obtener todas las propiedades de la clase CunetasHandleDataItem
        Dim props = Me.GetType().GetProperties()
        FilaData = RNG
        ' Iterar sobre cada propiedad


        For Each prop In props
            ' Obtener el nombre de la propiedad (que debe coincidir con un nombre en Headers)
            Dim propName = prop.Name

            SetProperties(propName, RNG, prop)
        Next
    End Sub

    Public Sub SetProperties(propName As String, RNG As Range, prop As Reflection.PropertyInfo) ', Headers As List(Of String))
        ' Buscar el índice correspondiente en Headers
        Dim headerIndex As Integer = Array.IndexOf(Headers, propName)

        ' Si el encabezado existe y el índice es válido
        If headerIndex >= 0 AndAlso headerIndex < IndexHeaders?.Length Then
            Dim valueIndex As Integer = IndexHeaders(headerIndex) - 1

            ' Verificar que el índice de la fila no sea negativo y esté dentro del rango de los valores de la fila
            If valueIndex >= 0 AndAlso valueIndex < (IndexHeaders.Length + 1) Then
                Dim value = RNG.Cells(1, valueIndex).value2

                If propName = "Comentarios" Then RNGComentarios = RNG.Cells(1, valueIndex)
                If propName = "Handle" Then RNGHandle = RNG.Cells(1, valueIndex)
                'Dim exc As List(Of String) from {""}

                ' Asignar el valor a la propiedad correspondiente usando reflection
                If value IsNot Nothing Then
                    Try
                        ' Convertir el valor y asignarlo a la propiedad
                        prop.SetValue(Me, Convert.ChangeType(value, prop.PropertyType))
                    Catch ex As Exception
                        Console.WriteLine($"Error al asignar el valor a la propiedad {propName}: {ex.Message}")
                    End Try
                End If
            End If
        End If
    End Sub

    'esta clase maneja la transferencias de informacion de la clase y excel
    ' Esta clase maneja la transferencia de información entre la clase y Excel
    Public Sub SyncPropExcel(PropNames As Array, F As String)
        ' Asegurarse de manejar las propiedades vía reflexión
        Dim props = Me.GetType().GetProperties()
        'Dim op As New ExcelOptimizer

        Try
            ' Desactivar funciones de Excel para optimizar el rendimiento
            'op.TurnEverythingOff()

            ' Recorrer las propiedades y actualizar Excel en función de los encabezados
            For Each propName As String In PropNames
                ' Buscar el índice del encabezado que coincida con el nombre de la propiedad
                Dim headerIndex As Integer = Array.IndexOf(Headers, propName) + 1

                ' Verificar que el índice del encabezado es válido
                If headerIndex >= 0 Then
                    ' Encontrar la propiedad usando reflexión
                    Dim prop = props.FirstOrDefault(Function(p) p.Name = propName)

                    ' Asegurarse de que la propiedad fue encontrada
                    ' Selección de acción según el valor de F ("Get" o "Set")
                    Select Case F
                        Case "Get"
                            ' Obtener el valor de la propiedad y escribirlo en la celda correspondiente en Excel
                            If prop IsNot Nothing Then

                                FilaData.Cells(1, headerIndex).Value = prop.GetValue(Me, Nothing)
                            Else
                                Debug.WriteLine($"Propiedad '{propName}' no encontrada.")
                            End If
                        Case "Set"
                            ' Leer el valor de la celda en Excel y asignarlo a la propiedad de la clase
                            Dim excelValue = FilaData.Cells(1, headerIndex).Value
                            ' Intentar convertir el valor al tipo de la propiedad antes de asignarlo
                            Dim convertedValue = Convert.ChangeType(excelValue, prop.PropertyType)
                            prop.SetValue(Me, convertedValue, Nothing)
                        Case Else
                            ' En caso de que F no tenga un valor válido
                            Debug.WriteLine($"Acción '{F}' no reconocida.")
                    End Select

                End If
            Next

        Finally
            ' Restaurar las configuraciones de Excel después de la operación
            'op.TurnEverythingOn()
            'op.Dispose()
        End Try
    End Sub

    Public Function GetFromExcel(PropName As String) As Object
        ' Asegurarse de manejar las propiedades vía reflexión
        Dim props = Me.GetType().GetProperties()
        Dim op As New ExcelOptimizer

        Try

            ' Desactivar funciones de Excel para optimizar el rendimiento
            op.TurnEverythingOff()

            ' Recorrer las propiedades y actualizar Excel en función de los encabezados
            ' Buscar el índice del encabezado que coincida con el nombre de la propiedad
            Dim headerIndex As Integer = Array.IndexOf(Headers, PropName) + 1

            ' Verificar que el índice del encabezado es válido
            If headerIndex >= 0 Then
                ' Encontrar la propiedad usando reflexión
                Dim prop = props.FirstOrDefault(Function(p) p.Name = PropName)

                ' Obtener el valor de la propiedad y escribirlo en la celda correspondiente en Excel
                If prop IsNot Nothing Then
                    ' Leer el valor de la celda en Excel y asignarlo a la propiedad de la clase
                    Dim excelValue = FilaData.Cells(1, headerIndex).Value
                    ' Intentar convertir el valor al tipo de la propiedad antes de asignarlo
                    Return Convert.ChangeType(excelValue, prop.PropertyType)
                End If
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        Finally
            ' Restaurar las configuraciones de Excel después de la operación
            op.TurnEverythingOn()
            op.Dispose()
        End Try
    End Function

    ' Método para calcular las propiedades en relación a la entidad, alignment y archivo
    ' Se pretende calcular las propiedades de una entidad dada

    Public Sub SetPropertiesFromDWG(Handle As String, AlingHandle As String, Optional ByRef RelateCuneta As CunetasHandleDataItem = Nothing)
        With Me
            CStationOffsetLabel.StrProcessEntity(Handle, .Layer, .MinX, .MinY, .MaxX, .MaxY, .StartStation, .EndStation, .Side, .Longitud, .Area, AlingHandle, FileName, FilePath, .Close)
            Me.Handle = Handle
            Me.AlignmentHDI = AlingHandle
            If Not CLHandle.CheckIfExistHd(Handle) Then Exit Sub
            Dim textExplo As String = GeneralSegmentLabelHelper.GetExplodedLabelText(
                                        AlignmentLabelHelper.GetLabelForLineOrCurve(
                                        CLHandle.GetEntityIdByStrHandle(Handle)))

            Me.CodigoCuneta = CodigoExtractor.ExtraerCodigo(textExplo, "CU\d+-\d+")

            If RelateCuneta IsNot Nothing Then
                Me.RelateCuneta = RelateCuneta
                Me.Comentarios = RelateCuneta.Comentarios
                Me.FilaData = RelateCuneta.FilaData
            End If
        End With
    End Sub
    'metodo para agregar info a un DGView 
    'en este metodo se pasara un DGView y el listado de propiedades y este agregara la informacion al DGView

    Public Sub AddToDGView(DGView As DataGridView, Optional RowIndex As Integer = -1)
        ' Crear una nueva fila para agregar al DataGridView
        Dim rowValues As New List(Of Object)
        Dim Headers As List(Of String) = HandleDataProcessor.Headers.Select(Function(item) item.Item1).ToList()
        ' Obtener todas las propiedades de la clase CunetasHandleDataItem, excepto las excluidas
        Dim props = Me.GetType().GetProperties()
        ' Iterar sobre cada propiedad
        For Each prop In props
            ' Excluir las propiedades que no deben estar en el DataGridView
            If Headers.Contains(prop.Name) Then 'AndAlso prop.Name <> "Accesos" Then
                ' Obtener el valor de la propiedad
                Dim value = prop.GetValue(Me, Nothing)

                ' Añadir el valor a la lista de valores para la fila
                rowValues.Add(value)
            End If
        Next
        ' Agregar o actualizar la fila en el DataGridView
        If RowIndex = -1 Then
            ' Si no se especifica un índice, agregar una nueva fila
            DGView.Rows.Add(rowValues.ToArray())
        Else
            ' Si se especifica un índice, actualizar la fila existente
            For i As Integer = 0 To rowValues.Count - 1
                DGView.Rows(RowIndex).Cells(i).Value = rowValues(i)
            Next
        End If
    End Sub

    'metodo para agregar info a un DGView 
    'en este metodo se pasara un DGView y el listado de propiedades y este agregara la informacion al DGView
    ' Método para agregar los datos de la clase a una tabla de Excel

    Public Sub AddToExcelTable(ByRef listObject As ListObject, Optional Exception As List(Of String) = Nothing, Optional RNGRow As Excel.Range = Nothing, Optional Formatos As List(Of String) = Nothing)
        ' Crear una nueva fila al final de la tabla
        'Dim RNGRow As Excel.Range
        If RNGRow Is Nothing Then RNGRow = listObject.ListRows.Add().Range

        Dim Headers As List(Of String) = HandleDataProcessor.Headers.Select(Function(item) item.Item1).ToList()
        Dim HeadersArray As String() = HandleDataProcessor.Headers.Select(Function(item) item.Item1).ToArray()
        'Headers.Insert(0, "Accesos")
        ' Obtener todas las propiedades de la clase CunetasHandleDataItem
        Dim props = Me.GetType().GetProperties()

        ' Iterar sobre cada encabezado definido en la tabla (Headers)
        For Each Header As String In Headers
            ' Buscar la propiedad correspondiente en la clase
            Dim prop = props.FirstOrDefault(Function(p) p.Name = Header)

            ' Obtener el índice del header
            Dim headerIndex As Integer = Array.IndexOf(HeadersArray, Header) + 1

            ' Verificar que la propiedad y el índice existen
            If prop IsNot Nothing AndAlso headerIndex > 0 AndAlso Not Exception.Contains(prop.Name) Then
                ' Obtener el valor de la propiedad
                Dim value As Object = prop.GetValue(Me, Nothing)

                ' Asignar el valor a la celda de la nueva fila
                If value IsNot Nothing Then
                    'Dim RNGIheader As Range = listObject.HeaderRowRange.Find(What:=Headers(headerIndex - 1))
                    RNGIheader = listObject.HeaderRowRange.Cells.Cast(Of Range)().FirstOrDefault(Function(c) c.Value = Headers(headerIndex - 1))

                    If RNGIheader IsNot Nothing Then
                        Dim IndexHeaderTBRNG As Integer = RNGIheader.Column - listObject.HeaderRowRange(1).column + 1
                        If prop.Name = "Layer" AndAlso RNGRow.Cells(1, IndexHeaderTBRNG).Value <> "" Then
                            prop.SetValue(Me, RNGRow.Cells(1, IndexHeaderTBRNG).Value)
                            value = prop.GetValue(Me, Nothing)
                        End If

                        RNGRow.Cells(1, IndexHeaderTBRNG).Value = value

                        ' Opcional: Formatear la celda según el tipo de dato segun item3 de headers
                        If Formatos IsNot Nothing AndAlso headerIndex < Formatos.Count Then
                            RNGRow.Cells(1, headerIndex).NumberFormat = Formatos(headerIndex - 1)
                        ElseIf prop.PropertyType Is GetType(Double) OrElse prop.PropertyType Is GetType(Decimal) Then 'AndAlso Exception.Contains(prop.Name) Then
                            RNGRow.Cells(1, headerIndex).NumberFormat = "0.00"
                        ElseIf prop.PropertyType Is GetType(Date) Then
                            RNGRow.Cells(1, headerIndex).NumberFormat = "mm/dd/yyyy"
                        Else
                            ' Asignar formato general para otro tipo de datos
                            RNGRow.Cells(1, headerIndex).NumberFormat = "@"
                        End If
                    Else

                    End If

                End If
            End If
        Next
    End Sub

    Public Shared Function LookForHandleInTB(Handle As String) As Range
        Dim ExRNGSe As New ExcelRangeSelector()

        Dim Table As Microsoft.Office.Interop.Excel.ListObject = ExRNGSe.GetTableOnWorkBkByName("CunetasGeneral")

        If Table IsNot Nothing Then
            Dim indexHeader As Integer() = UpdateExcelTable.ConArray({"Handle", "Comentarios"}, Table.HeaderRowRange)

            Dim handleColumnIndex As Integer = indexHeader(0)

            Dim ComentariosColumnIndex As Integer = indexHeader(1)

            For Each row As Microsoft.Office.Interop.Excel.ListRow In Table.ListRows
                If row.Range.Cells(1, handleColumnIndex).Value.ToString() = Handle Then
                    Return row.Range ' Return the matching row if found
                End If
            Next
        End If

        Return Nothing
    End Function


    '<Sumary>
    'manejo de base de datos 
    '<Sumary>
    Public Sub SettingBaseDDatos()
        'Comentarios de código adicional para manejar sincronización con la base de datos (desactivado)
        ' Obtener la carpeta de la ruta del archivo usando Path.GetDirectoryName
        'Crear la tabla basada en las propiedades de CunetasHandleDataItem si no existe
        'DbManager.CreateTableIfNotExists(Of CunetasHandleDataItem)(TableName, "ID_Cunetas")
    End Sub

    ''' <summary>
    ''' Sincroniza los datos de la instancia actual con la base de datos.
    ''' Si el handle ya existe, actualiza los datos; si no, inserta una nueva fila.
    ''' </summary>
    ''' <param name="cdbManager">Instancia de DataBSQLManager para gestionar la base de datos.</param>
    ''' <param name="ctableName">El nombre de la tabla en la que se deben sincronizar los datos.</param>
    Public Sub SyncWithDatabase(Optional cdbManager As DataBSQLManager = Nothing, Optional ctableName As String = "")
        If String.IsNullOrEmpty(ctableName) Then ctableName = TableName
        'If cdbManager Is Nothing Then cdbManager = DbManager

        ' Nombre de la clave primaria
        Dim primaryKeyName As String = "ID_Cunetas"

        ' Intentar obtener el ID por Handle
        Dim id As Integer = cdbManager.GetRowByHandle(ctableName, Me, primaryKeyName)

        ' Si no se encontró por Handle, buscar por las propiedades
        If id = -1 Then
            id = cdbManager.GetRowByProperties(ctableName, Me, primaryKeyName)
        End If

        ' Actualizar o insertar según corresponda
        If id <> -1 Then
            cdbManager.UpdateDataID(ctableName, Me, primaryKeyName, id)
        Else
            cdbManager.InsertData(Me, ctableName)
        End If
    End Sub

    ''' <summary>
    ''' Carga los datos de la base de datos en la instancia actual.
    ''' </summary>
    Public Sub LoadFromDatabase()
        ' Attempt to retrieve the data item by handle
        'Dim dataFromDb As CunetasHandleDataItem = DbManager.GetDataItemByHandle(Of CunetasHandleDataItem)(TableName, Me.Handle)

        'If dataFromDb IsNot Nothing Then
        '    ' Use reflection to copy properties from dataFromDb to this instance
        '    For Each prop In Me.GetType().GetProperties()
        '        ' Ensure property exists and is not null in dataFromDb
        '        Dim dbValue = prop.GetValue(dataFromDb)
        '        If dbValue IsNot Nothing Then
        '            Try
        '                ' Assign value from dataFromDb to current instance
        '                prop.SetValue(Me, dbValue)
        '            Catch ex As Exception
        '                Console.WriteLine($"Error setting property '{prop.Name}': {ex.Message}")
        '            End Try
        '        End If
        '    Next
        'Else
        '    ' Log if handle is not found in the database
        '    Console.WriteLine($"Handle '{Me.Handle}' not found in the database '{TableName}'.")
        'End If
    End Sub
End Class

Public Class LosasDataItem
    'base de datos para esta clase
    Public Property DbManager As New DataBSQLManager()

    ' Evento que se disparará cuando la propiedad Name cambie
    Public Event EndStationChanged As EventHandler
    Public Event HandleChanged As EventHandler
    Public Property Type As String
    Public Property CatchHandle As String
    ''' <summary>
    ''' Propiedad para manejar cambios en el Handle.
    ''' Incluye lógica para manejar conversiones y actualizaciones.
    ''' </summary>
    Public Property Handle As String

        Get
            Return CatchHandle
        End Get
        Set(value As String)
            ' Trigger the event only if the value changes
            If CatchHandle <> value Then
                CatchHandle = value
                ' Trigger the EndStationChanged event
                RaiseEvent HandleChanged(Me, EventArgs.Empty)
                If CLHandle.CheckIfExistHd(CatchHandle) Then
                    Type = TypeName(CLHandle.GetEntityByStrHandle(CatchHandle))
                    SetAccesoNum()

                    PLANO = PlanoNombre(CLHandle.GetEntityIdByStrHandle(CatchHandle))

                    'If Type = "Polyline" Then
                    '    Polyline = CLHandle.GetEntityByStrHandle(CatchHandle)
                    'ElseIf Type = "FeatureLine" OrElse Type = "Polyline3d" OrElse Type = "Polyline2d" OrElse Type = "line" Then
                    '    ConvertToPolyline(False, Polyline:=Polyline)
                    'Else

                    '    GoTo HandleErroneo
                    'End If

                    Dim modifier As New PolylineModifier()

                    'modifier.ChangePolylineProperties(CatchHandle, "FLUXO3", 4.5, 0.5, Nothing)

                    'Dim AligmentsCertify As New AlignmentHelper()

                    'If EndStation <> 0 AndAlso StartStation <> 0 Then
                    '    AligmentsCertify.CheckByStations(CType(Polyline, Autodesk.AutoCAD.DatabaseServices.Entity), Side, StartStation, EndStation, 2)
                    'Else
                    '    AligmentsCertify.CheckAlimentByProximida(Polyline, False)
                    'End If

                    'Alignment = AligmentsCertify.Alignment
                    'If Alignment IsNot Nothing Then
                    '    AlignmentHDI = Alignment.Handle.ToString()
                    'Else
                    '    Alignment = AligmentsCertify.CheckCunetasAlignment(Polyline)
                    '    AlignmentHDI = Alignment.Handle.ToString()
                    'End If
                End If
HandleErroneo:
            End If
        End Set
    End Property

    ' Enhanced property implementation for updating entity layers intelligently
    Public Event LayerChanged As EventHandler

    ' Columnas visibles en el DataGridView o tabla
    Public Property Acceso As Double

    Public CatchTableName As String

    Private _catchLayer As String

    ' Property to get or set the layer name and trigger updates accordingly
    Public Property Layer As String
        Get
            Return _catchLayer
        End Get
        Set(value As String)
            ' Trigger the event and update only if the layer value has actually changed
            If _catchLayer <> value Then
                _catchLayer = value

                ' Raise the LayerChanged event safely, checking for null subscribers
                RaiseEvent LayerChanged(Me, EventArgs.Empty)

                ' Check if the entity handle is valid and the layer needs to be updated
                'If Not String.IsNullOrEmpty(Handle) AndAlso CLHandle.CheckIfExistHd(Handle) Then
                '    Dim currentLayer As String = CLayerHelpers.GetLayersAcEnt(Handle)

                '    ' Update the entity layer only if it's different from the desired layer
                '    If currentLayer <> _catchLayer Then
                '        CLayerHelpers.ChangeLayersAcEnt(Handle, _catchLayer)
                '    End If

                'End If
                'SyncWithDatabase()
            End If
        End Set
    End Property

    Public Property TableName As String
        Get
            Return CatchTableName
        End Get
        Set(value As String)
            If CatchHandle <> value Then
                CatchTableName = value
                SettingBaseDDatos()
            End If
        End Set
    End Property
    Public Property StartStation As Double

    Private _endStation As Double

    ' Public property EndStation with validation logic
    Public Property EndStation As Double
        Get
            Return _endStation
        End Get
        Set(value As Double)
            ' Trigger the event only if the value changes
            If _endStation <> value Then
                _endStation = value

                ' Trigger the EndStationChanged event
                RaiseEvent EndStationChanged(Me, EventArgs.Empty)

                ' Check if StartStation > EndStation and perform Excel operations
                If StartStation > _endStation Then
                    CStationOffsetLabel.GetMxMnOPPL(StartStation, _endStation)
                    'SyncPropExcel({"StartStation", "EndStation"}, "Get")
                End If
            End If
        End Set
    End Property
    'option de de seleccion de alineamiento para asignacion
    Public Property Perimetro As Double
    Public Property Area As Double
    Public Property Side As String
    Public Property Comentarios As String
    Public Property Tramos As String
    Public Property PLANO As String
    Public Property FileName As String
    Public Property FilePath As String

    Public Sub New()
        ' Initialization code if needed
        TableName = "LosasDatos"
        CorreccionesTBName = "LosasDatosCorregidas"
        SetAccesoNum()
    End Sub
    Public Sub SetAccesoNum()
        Dim docName As String = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name
        ' Si no se encuentra el código, salir
        Dim RefAcceso As String = CodigoExtractor.AccesoNum(docName)
        If String.IsNullOrEmpty(RefAcceso) Then
            Console.WriteLine("El archivo no presenta referencia de accesos.")
            Exit Sub
        End If
        ' Obtener el número de acceso a partir de la referencia extraída
        Accesos = CDbl(CodigoExtractor.ExtraerCodigo(RefAcceso, "\d+"))
        FilePath = docName
        FileName = Right(docName, Len(docName) - InStrRev(docName, "\"))
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    '<Sumary>
    'manejo de base de datos 
    '<Sumary>
    Public Sub SettingBaseDDatos()
        'Comentarios de código adicional para manejar sincronización con la base de datos (desactivado)
        ' Obtener la carpeta de la ruta del archivo usando Path.GetDirectoryName
        'Crear la tabla basada en las propiedades de CunetasHandleDataItem si no existe
        DbManager.CreateTableIfNotExists(Of CunetasHandleDataItem)(TableName, "ID_Cunetas")
    End Sub

    '''' <summary>
    '''' Sincroniza los datos de la instancia actual con la base de datos.
    '''' Si el handle ya existe, actualiza los datos; si no, inserta una nueva fila.
    '''' </summary>
    '''' <param name="cdbManager">Instancia de DataBSQLManager para gestionar la base de datos.</param>
    '''' <param name="ctableName">El nombre de la tabla en la que se deben sincronizar los datos.</param>
    'Public Sub SyncWithDatabase(Optional cdbManager As DataBSQLManager = Nothing, Optional ctableName As String = "")
    '    If String.IsNullOrEmpty(ctableName) Then ctableName = TableName
    '    If cdbManager Is Nothing Then cdbManager = DbManager

    '    ' Nombre de la clave primaria
    '    Dim primaryKeyName As String = "ID_Cunetas"

    '    ' Intentar obtener el ID por Handle
    '    Dim id As Integer = cdbManager.GetRowByHandle(ctableName, Me, primaryKeyName)

    '    ' Si no se encontró por Handle, buscar por las propiedades
    '    If id = -1 Then
    '        id = cdbManager.GetRowByProperties(ctableName, Me, primaryKeyName)
    '    End If

    '    ' Actualizar o insertar según corresponda
    '    If id <> -1 Then
    '        cdbManager.UpdateDataID(ctableName, Me, primaryKeyName, id)
    '    Else
    '        cdbManager.InsertData(Me, ctableName)
    '    End If
    'End Sub

    ''' <summary>
    ''' Carga los datos de la base de datos en la instancia actual.
    ''' </summary>
    Public Sub LoadFromDatabase()
        ' Attempt to retrieve the data item by handle
        Dim dataFromDb As CunetasHandleDataItem = DbManager.GetDataItemByHandle(Of CunetasHandleDataItem)(TableName, Handle)

        If dataFromDb IsNot Nothing Then
            ' Use reflection to copy properties from dataFromDb to this instance
            For Each prop In Me.GetType().GetProperties()
                ' Ensure property exists and is not null in dataFromDb
                Dim dbValue = prop.GetValue(dataFromDb)
                If dbValue IsNot Nothing Then
                    Try
                        ' Assign value from dataFromDb to current instance
                        prop.SetValue(Me, dbValue)
                    Catch ex As Exception
                        Console.WriteLine($"Error setting property '{prop.Name}': {ex.Message}")
                    End Try
                End If
            Next
        Else
            ' Log if handle is not found in the database
            Console.WriteLine($"Handle '{Handle}' not found in the database '{TableName}'.")
        End If
    End Sub

End Class

