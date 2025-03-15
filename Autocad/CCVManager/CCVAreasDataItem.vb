Imports System.Linq
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Public Class CCVDataClosedPL


    ' Evento que se disparará cuando la propiedad Name cambie
    Public Event NameChanged As EventHandler

    ' Campo privado para almacenar el valor de la propiedad Name
    Private _name As String

    Public Property Handle As String
    ' Propiedad pública Name
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            ' Solo dispara el evento si el valor cambia
            If _name <> value Then
                _name = value
                ' Disparar el evento NameChanged
                RaiseEvent NameChanged(Me, EventArgs.Empty)
                ' Actualizar el índice automáticamente cuando el nombre cambia
                ActualizarPropByIndex()

            End If
        End Set
    End Property

    Public Property GrupoPointName As String

    'Propiedades sync with BorderPLines properties en dynimicly change
    Public Property Layer As String
    Public Property MinX As Double
    Public Property MinY As Double
    Public Property MaxX As Double
    Public Property MaxY As Double
    Public Property StartStation As Double
    Public Property EndStation As Double
    Public Property Longitud As Double
    Public Property Area As Double
    Public Property Side As String
    Public Property AlignmentHDI As String
    'propiedades internas 
    Public Property Index As Integer

    Public Property Border As Polyline

    Private ReadOnly dataGridViewEventHandlers As New List(Of DataGridViewEventHandler)
    Public Property GogoPoints As List(Of CogoPoint)
    Public Property GroupCOGOPoints As New Dictionary(Of String, ObjectId)()
    Public Property BorderPLines As New Dictionary(Of String, BorderPLines)
    Public Shared ReadOnly Property Headers As String() = {"Handel", "Name", "Grupo Point Name",
                                                            "Layer", "MinX", "MinY", "MaxX",
                                                            "MaxY", "StartStation", "EndStation",
                                                            "Longitud", "Area", "Side", "AlignmentHDI"}

    Public ReadOnly columnFormats As New List(Of String) From {
        "@",'handle
        "@",
        "@",
        "@",'layer
        "0.000",
        "0.000",
        "0.000",
        "0.000",
        "0+000.00",'StartSTation
        "0+000.00",'EndStation
        "0.000",
        "0.000",
        "@",
        "@",
        "@",
        "@",
        "@",
        "@",
        "@"
    }
    Public EventsHandler As New List(Of EventHandlerClass)
    Public PanelDGViews As TabDGViews
    Public Sub New(DGView As System.Windows.Forms.DataGridView,
                   ByRef CAcadHelp As ACAdHelpers,
                   parentForm As System.Windows.Forms.Form)
        ''PanelDGViews = New TabDGViews(tabControlDGView, columnFormats, CAcadHelp, Me, handleProcessor) With {.TabName = "Puntos CCV", .DGViewName = "PuntosCCVDGView"}
        ''PuntosCCVDGView = PanelDGViews.DGView

        ''Dim PanelMediciones1 As New TabsMenu(ControlManager._tabControlMenu) With {.TabName = "CCV"}
        'cargar polyline cerradas generadas de datos CCV
        ''EventsHandler.Add(New EventHandlerClass(btContructCCVAreas, "click", Sub() RemoveVertice(PuntosCCVDGView)))


        ''Obtener la colección de grupos de puntos de Civil 3D
        ''GroupCOGOPoints=(nombre de grupo, objectid del grupo)
        'GroupCOGOPoints = GetPointGroupIds()
        'Dim i As Integer = 0
        'For Each key As String In GroupCOGOPoints.Keys
        '    'GogoPoints = DirectCast(GroupCOGOPoints(GroupName).GetObject(OpenMode.ForRead), CogoPoint)
        '    'GogoPoints is a list of the   CogoPoints in group named by string key.value
        '    'If key = "a-ACC13T1 - LV. REGULARIZACION LOSA LD - KM 1+300 - 1+304" Then
        '    '    Stop
        '    'End If
        '    GogoPoints = GetCOGOPointCollections(key)
        '    'clase para representar los bordes creados con los puntos y que formaran un poligono cerrado 
        '    Dim BorderPLine As New BorderPLines(GogoPoints, CAcadHelp.Alignment)

        '    BorderPLines.Add("Area-" & i, BorderPLine)
        '    i += 1
        'Next
        '' Configurar cada DataGridView
        'DateViewSet.ConfigureDataGridView(DGView, columnFormats)

        '' Crear y agregar el manejador de eventos a la lista
        'Dim eventHandler As New DataGridViewEventHandler(DGView, CAcadHelp, parentForm, Nothing)
        'dataGridViewEventHandlers.Add(eventHandler)
        ''AÑADIR LOS ELEMENTOA A AL DGVIEW

        'Dim Headers As New List(Of (String, String))()
        'Dim CleanedHeaders As New List(Of String)()

        'For Each header As String In CCVDataClosedPL.Headers
        '    ' Clean header by removing spaces and trimming
        '    Dim cleanedHeader As String = Trim(Replace(header, " ", ""))

        '    ' Add the cleaned header to both lists
        '    CleanedHeaders.Add(cleanedHeader)
        '    i = Array.IndexOf(CCVDataClosedPL.Headers, header)
        '    'Headers.Add((cleanedHeader, header, CCVDataClosedPL.columnFormats(i)))
        'Next

        '' Add columns using the pair headers
        'DataGridViewHelper.Addcolumns(DGView, Headers, columnFormats)

        '' Add data to the DataGridView
        'For index As Integer = 0 To GroupCOGOPoints.Count
        '    If index < GroupCOGOPoints.Count Then
        '        Me.Name = GroupCOGOPoints.Keys.ElementAt(index)
        '        AddToDGView(DGView, CleanedHeaders.ToArray())
        '    End If
        'Next
    End Sub

    Public Function GetCOGOPointCollections(pointGroupName As String) As List(Of CogoPoint)
        ' Obtener el documento actual de Civil 3D
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Using LockDoc As DocumentLock = doc.LockDocument
            ' Iniciar una transacción
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Obtener la colección de grupos de puntos de Civil 3D
                    Dim civildoc As CivilDocument = CivilApplication.ActiveDocument
                    Dim pointGroups As PointGroupCollection = civildoc.PointGroups
                    Dim copt As New List(Of CogoPoint)()

                    ' Verificar si hay grupos de puntos
                    If pointGroups.Count > 0 Then
                        ' Obtener el nombre del primer grupo de puntos

                        If pointGroupName = [String].Empty Then
                            Return Nothing
                        End If

                        ' Obtener el grupo de puntos por nombre
                        Dim pointGroupId As ObjectId = GetPointGroupIdByName(pointGroupName)
                        Dim group As PointGroup = TryCast(pointGroupId.GetObject(OpenMode.ForRead), PointGroup)

                        ' Obtener los números de los puntos del grupo
                        Dim pointNumbers As UInteger() = group.GetPointNumbers()

                        ' Recorrer los números de puntos y obtener los CogoPoints correspondientes
                        For Each pointNumber As UInteger In pointNumbers
                            Dim colCogop As CogoPointCollection = civildoc.CogoPoints()
                            Dim ccogoPoint As CogoPoint = trans.GetObject(colCogop.GetPointByPointNumber(pointNumber), OpenMode.ForRead)
                            copt.Add(ccogoPoint)
                        Next
                    Else
                        ed.WriteMessage(vbLf & "No hay grupos de puntos disponibles.")
                        Return Nothing
                    End If

                    ' Completar la transacción
                    trans.Commit()
                    Return copt
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message)
                    Return Nothing
                Finally
                    trans.Dispose()
                End Try
            End Using
        End Using

    End Function

    ' Obtener el ObjectId de un grupo de puntos por su nombre
    ' Function to get the ObjectId of a PointGroup by its name
    Private Function GetPointGroupIdByName(groupName As String) As ObjectId
        Dim civildoc As CivilDocument = CivilApplication.ActiveDocument
        Dim pointGroups As PointGroupCollection = civildoc.PointGroups

        If pointGroups.Contains(groupName) Then
            Dim pointGroupId As ObjectId = pointGroups(groupName)
            Return pointGroupId
        Else
            Throw New ArgumentException("No point group found with the specified name.")
        End If
    End Function

    ' Function to get a dictionary of all PointGroup names and their ObjectIds
    Private Function GetPointGroupIds() As Dictionary(Of String, ObjectId)
        ' Iniciar una transacción
        ' Obtener el documento actual de Civil 3D
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Using LockDoc As DocumentLock = doc.LockDocument
            ' Iniciar una transacción
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    Dim civildoc As CivilDocument = CivilApplication.ActiveDocument
                    Dim pointGroups As PointGroupCollection = civildoc.PointGroups
                    Dim result As New Dictionary(Of String, ObjectId)()
                    ' Verificar si hay grupos de puntos disponibles
                    If pointGroups.Count = 0 Then
                        ed.WriteMessage(vbLf & "No hay grupos de puntos disponibles en el documento.")
                        Return result
                    End If
                    For Each pointGroupId As ObjectId In pointGroups
                        ' Get the PointGroup object from its ObjectId
                        Dim pointGroup As PointGroup = TryCast(trans.GetObject(pointGroupId, OpenMode.ForRead), PointGroup)

                        ' Asegurarse de que el grupo de puntos no es nulo
                        If pointGroup IsNot Nothing Then
                            ' Añadir el nombre del grupo de puntos y su ObjectId al diccionario

                            result.Add(pointGroup.Name, pointGroupId)
                        Else
                            ed.WriteMessage(vbLf & $"No se pudo leer el grupo de puntos con ObjectId: {pointGroupId}.")
                        End If
                    Next
                    Return result
                    trans.Commit()
                Catch ex As Exception
                    Return New Dictionary(Of String, ObjectId)
                Finally
                    trans.Dispose()
                End Try
            End Using
        End Using
    End Function
    ' Método que actualiza el índice y otras áreas cuando cambia el nombre
    Private Sub ActualizarPropByIndex()
        ' Aquí puedes añadir la lógica para actualizar el índice
        ' o realizar cualquier otro cálculo cuando el nombre cambia
        Me.Index = GroupCOGOPoints.Keys.ToList().IndexOf(Me.Name)
        Console.WriteLine("El índice ha sido actualizado a: " & Me.Index)
        ' Aquí también puedes recalcular las áreas, actualizar polilíneas, etc.
        ' Ejemplo:
        If BorderPLines.ContainsKey("Area-" & Me.Index) Then
            Dim selectedBorder As BorderPLines = BorderPLines("Area-" & Me.Index)
            Console.WriteLine("El área correspondiente ha sido seleccionada.")
            Me.Handle = selectedBorder.Border.Handle.ToString()
            'Me.GrupoPointName = GogoPoints(Me.Index).PointName
            Me.GrupoPointName = GroupCOGOPoints.Keys.ElementAt(Index)
            UpdatePropertiesFromBorder(selectedBorder)
        End If
    End Sub
    ' Método para actualizar las propiedades de esta clase con las de selectedBorder
    Public Sub UpdatePropertiesFromBorder(selectedBorder As BorderPLines)
        ' Obtiene todas las propiedades de la clase actual (Me)
        Dim props = Me.GetType().GetProperties()

        ' Itera a través de las propiedades de la clase actual (Me)
        For Each prop In props
            ' Obtiene el nombre de la propiedad actual en Me
            Dim propName As String = prop.Name

            ' Obtiene la propiedad correspondiente en selectedBorder
            Dim correspondingProperty = selectedBorder.GetType().GetProperty(propName)

            ' Si existe una propiedad con el mismo nombre en selectedBorder y los tipos son compatibles
            If correspondingProperty IsNot Nothing AndAlso prop.PropertyType.IsAssignableFrom(correspondingProperty.PropertyType) Then
                ' Asigna el valor de la propiedad de selectedBorder a la propiedad correspondiente en Me
                prop.SetValue(Me, correspondingProperty.GetValue(selectedBorder))
            End If
        Next
    End Sub
    'metodo para agregar a un DGView 
    'en este metodo se pasara un DGView y el listado de propiedades y este agregara la informacion al DGView
    Public Sub AddToDGView(DGView As System.Windows.Forms.DataGridView, CleanedHeaders As String())
        ' Crear una nueva fila para agregar al DataGridView
        Dim rowValues As New List(Of Object)

        ' Obtener todas las propiedades de la clase CunetasHandleDataItem, excepto las excluidas
        Dim props = Me.GetType().GetProperties()
        ' Iterar sobre cada propiedad
        For Each prop In props
            ' Excluir las propiedades que no deben estar en el DataGridView
            If CleanedHeaders.Contains(prop.Name) AndAlso prop.Name <> "Accesos" Then
                ' Obtener el valor de la propiedad
                rowValues.Add(prop.GetValue(Me, Nothing))
            End If
            ' Añadir el valor a la lista de valores para la fila
        Next
        ' Agregar la fila al DataGridView usando los valores de las propiedades
        DGView.Rows.Add(rowValues.ToArray())
    End Sub

    'RemoveVertice()
    'proceso pra eliminar un punto de un border definido y redrawing el borde 
    'identificar el punto by selecte row in dgview 
    'hylight the group of points para la posterior seleccion de uno o varios puntos 
    'enter para aceptar la selecction 
    'mensase de error si se selecciona un punto que no pertenece al listado y permaneciento enla seleccion
    'cancel option
    Public Sub RemoveVertice(DGView As System.Windows.Forms.DataGridView)
        ' Ensure a row is selected in the DataGridView
        If DGView.SelectedCells.Count <> 1 Then
            MsgBox("Please select a row corresponding to the polyline to remove a vertex.", MsgBoxStyle.Information, "Selection Required")
            Return
        End If

        ' Get the selected row and associated polyline

        'Dim selectedRow As System.Windows.Forms.DataGridViewRow = DGView.Rows(DGView.SelectedCells(0).RowIndex).Cells()

        Dim GroupName As String = DGView.Rows(DGView.SelectedCells(0).RowIndex).Cells("Name").Value


        Dim borderIndex As Integer = GroupCOGOPoints.Keys.ToList().IndexOf(GroupName) 'selectedRow.Index
        If Not BorderPLines.ContainsKey("Area-" & borderIndex) Then
            MsgBox("The selected polyline could not be found.", MsgBoxStyle.Critical, "Error")
            Return
        End If

        Dim selectedBorder As BorderPLines = BorderPLines("Area-" & borderIndex)
        Me.Name = GroupCOGOPoints.Keys.ElementAt(borderIndex)

        SetPointStyle(GroupName)


        '' Highlight the points of the polyline
        'HighlightPoints(GogoPoints)

        '' Prompt the user to select a point (vertex) from the polyline in AutoCAD
        'Dim selectedPoint As Point3d
        'Try
        '    selectedPoint = SelectPointFromAutoCAD()
        'Catch ex As InvalidOperationException
        '    MsgBox("Point selection was canceled or invalid.", MsgBoxStyle.Information, "Canceled")
        '    Return
        'End Try

        '' Validate the selected point: Check if it belongs to the polyline (border)
        'Dim validVertex As Boolean = False
        'Dim vertexIndex As Integer = -1

        'For i As Integer = 0 To selectedBorder.Border.NumberOfVertices - 1
        '    Dim vertex As Point3d = selectedBorder.Border.GetPoint3dAt(i)
        '    If selectedPoint.Equals(vertex) Then
        '        validVertex = True
        '        vertexIndex = i
        '        Exit For
        '    End If
        'Next

        'If Not validVertex Then
        '    MsgBox("The selected point does not belong to the polyline. Please select a valid vertex.", MsgBoxStyle.Critical, "Invalid Selection")
        '    ' Option to retry or cancel
        '    If MsgBox("Would you like to try again?", MsgBoxStyle.YesNo, "Retry?") = MsgBoxResult.Yes Then
        '        RemoveVertice(DGView)
        '    End If
        '    Return
        'End If

        '' Remove the selected vertex from the polyline
        'Using tr As Transaction = selectedBorder.Border.Database.TransactionManager.StartTransaction()
        '    Try
        '        ' Remove the vertex at the specified index
        '        selectedBorder.Border.RemoveVertexAt(vertexIndex)

        '        ' Optionally, update properties like area, length, etc.
        '        UpdatePropertiesFromBorder(selectedBorder)

        '        ' Commit the transaction to redraw the polyline in AutoCAD
        '        tr.Commit()

        '        MsgBox("Vertex removed successfully.", MsgBoxStyle.Information, "Success")
        '    Catch ex As Exception
        '        MsgBox("An error occurred while removing the vertex: " & ex.Message, MsgBoxStyle.Critical, "Error")
        '    End Try
        'End Using
    End Sub
    Public Sub SetPointStyle(GroupName As String)
        'agregar los style de GroupCOGOPoints
        'Obtener el ObjectId del grupo de puntos correspondiente
        Dim groupId As ObjectId = GroupCOGOPoints(GroupName)

        Dim styleManager As New CogoPointGroupStyleManager()

        Dim pointStyleManager As New CogoPointStyleManager(CivilApplication.ActiveDocument)

        ' Asignar un estilo de punto al grupo de puntos antes de realizar cualquier modificación
        Dim pointStyleId As ObjectId = pointStyleManager.GetPointStyleByName("EstiloPersonalizado") ' Ajustar el nombre del estilo según sea necesario
        'Crear el estilo nuevo si no existe 
        'Usar el método para crear un nuevo estilo de punto personalizado
        ' Verificar si el estilo existe, si no, crearlo
        If pointStyleId = ObjectId.Null Then
            pointStyleId = pointStyleManager.CreateCustomPointStyle("EstiloPersonalizado", markerType:=CustomMarkerType.CustomMarkerDot, markerSuperimposeType:=CustomMarkerSuperimposeType.Circle)
        End If

        Dim labelStyleId As ObjectId = pointStyleManager.GetPointStyleByName("EtiquetaPersonalizada") ' Ajustar el nombre de la etiqueta según sea necesario

        If labelStyleId = ObjectId.Null Then
            labelStyleId = pointStyleManager.CreateCustomPointStyle("EtiquetaPersonalizada", markerType:=CustomMarkerType.CustomMarkerDot, markerSuperimposeType:=CustomMarkerSuperimposeType.Circle, markerColor:=Color.FromColorIndex(ColorMethod.ByAci, 2), labelColor:=Color.FromColorIndex(ColorMethod.ByAci, 3))
        End If

        styleManager.AssignPointStyleToGroup(GroupName, pointStyleId, labelStyleId)
    End Sub
    'HighlightPoints
    'en esta clase se pretende seleccionar y desceseccionar los CogoPoints  a eleiminar o dejar para la creaccion de perimetros 
    'debe de haber un la seleccion recurciva con solo lo puntos en GogoPoints
    'ExcludedGogoPoints points seria los no seleccionados 
    'para laseleccion sele debe de aggregar un style temporal para la visualizacion del punto 
    Public Sub HighlightPoints(GogoPoints As List(Of CogoPoint))
        ' Highlight each CogoPoint by temporarily changing its color or style
        For Each cogoPoint As CogoPoint In GogoPoints
            HighlightCogoPoint(cogoPoint, Color.FromRgb(255, 0, 0)) ' Highlight in red
        Next
    End Sub
    Private Sub HighlightCogoPoint(cogoPoint As CogoPoint, highlightColor As Color)
        ' Change the color of the CogoPoint temporarily for highlighting
        Using tr As Transaction = cogoPoint.Database.TransactionManager.StartTransaction()
            Try
                Dim cogoEnt As CogoPoint = CType(tr.GetObject(cogoPoint.ObjectId, OpenMode.ForWrite), CogoPoint)
                cogoEnt.Color = highlightColor ' Set to red or another color
                tr.Commit()
            Catch ex As Exception
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error highlighting CogoPoint: " & ex.Message)
            End Try
        End Using
    End Sub
    ' The SelectPointFromAutoCAD method
    Public Function SelectPointFromAutoCAD() As Point3d
        ' Get the current AutoCAD document editor
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        ' Prompt the user to select a point
        Dim promptResult As PromptPointResult = ed.GetPoint("Select a vertex point from the polyline:")

        ' Check if the selection was successful
        If promptResult.Status = PromptStatus.OK Then
            Return promptResult.Value ' Return the selected Point3d
        Else
            Throw New InvalidOperationException("Point selection was canceled or invalid.")
        End If
    End Function
End Class


'esta clase representara los borde representados por el area delimitado por los punto CogoPoits
'esta solo tendra valor si se considera que los puntos correspponden a un poligono cerrado 
Public Class BorderPLines
    'Public Property Color As Color
    Public Color As Color
    Public Property Border As Polyline

    ' Evento que se disparará cuando la propiedad Layer cambie
    Public Event LayerChanged As EventHandler

    ' Campo privado para almacenar el valor de la propiedad Name
    Private _Layer As String

    Public Property Handle As String
    ' Propiedad pública Name
    Public Property Layer As String
        Get
            Return _Layer
        End Get
        Set(value As String)
            ' Solo dispara el evento si el valor cambia
            If _Layer <> value Then
                _Layer = value
                ' Disparar el evento NameChanged
                RaiseEvent LayerChanged(Me, EventArgs.Empty)
                ' Actualizar el índice automáticamente cuando el nombre cambia
                CrearGrupoYCapas(_Layer, {_Layer})
                'verde 
                LayerColorChanger.ChColorLY(_Layer, Color)
                CLayerHelpers.ChangeLayersAcEnt(Handle, _Layer)
                'Border.Layer = CatchLayer

            End If
        End Set
    End Property

    Public Property MinX As Double
    Public Property MinY As Double
    Public Property MaxX As Double
    Public Property MaxY As Double
    Public Property StartStation As Double
    Public Property EndStation As Double
    Public Property Longitud As Double
    Public Property Area As Double
    Public Property Side As String

    ' Evento que se disparará cuando la propiedad Layer cambie
    Private _AlignmentHDI As String

    Public Event AlignmentHDIChanged As EventHandler
    Public Property AlignmentHDI As String
        Get
            Return _AlignmentHDI
        End Get
        Set(value As String)
            ' Solo dispara el evento si el valor cambia
            If _AlignmentHDI <> value Then
                _AlignmentHDI = value
                ' Disparar el evento NameChanged
                RaiseEvent AlignmentHDIChanged(Me, EventArgs.Empty)
                ' Actualizar el índice automáticamente cuando el nombre cambia
                SetPropertiesFromDWG(Handle, _AlignmentHDI)
            End If
        End Set
    End Property
    Public Property PLANO As String
    Public Property FileName As String
    Public Property FilePath As String

    Public Sub New(GogoPoints As List(Of CogoPoint), Optional ByRef Alignment As Alignment = Nothing)

        Border = CrearBorders(GogoPoints)
        'Public Function CrearGrupoYCapas(LayerFNAme As String, nombresCapas As String()) As List(Of String)
        Color = Color.FromRgb(153, 255, 0)
        Handle = Border.Handle.ToString()
        Layer = "Aceras"
        AlignmentHDI = Alignment.Handle.ToString()
    End Sub
    Public Function CrearBorders(GogoPoints As List(Of CogoPoint))
        Dim Vertices As New List(Of Point3d)()

        ' Línea que converge en un solo vértice
        For Each Cogo As CogoPoint In GogoPoints
            Vertices.Add(Cogo.Location)
        Next

        Dim VerticesOrdenados As Dictionary(Of String, Point3d) = CGPointHelper.OrderVertices(Vertices)
        Dim Border As Polyline = CGPointHelper.CrearPL(VerticesOrdenados.Values.ToList())
        Border.Closed = True
        CGPointHelper.AddToModal(Border)
        Return Border

    End Function

    ' Método para calcular las propiedades en relación a la entidad, alignment y archivo
    ' Se pretende calcular las propiedades de una entidad dada
    Public Sub SetPropertiesFromDWG(Handle As String, AlingHandle As String)
        With Me
            CStationOffsetLabel.StrProcessEntity(Handle, .Layer, .MinX, .MinY, .MaxX, .MaxY, .StartStation, .EndStation, .Side, .Longitud, .Area, AlingHandle)
            Me.Handle = Handle
        End With
    End Sub
End Class

'esta clase manejara la base de dato sqlite de para las instacias creadas tanto de la clase ccvdataClosedpl y BorderPLines


'Public Class SqlBase
'    Private _connectionString As String

'    ' Constructor to initialize the database connection
'    Public Sub New(dbFilePath As String)
'        ' Check if the database file exists, otherwise create it
'        If Not File.Exists(dbFilePath) Then
'            SQLiteConnection.CreateFile(dbFilePath)
'        End If

'        ' Set the connection string
'        _connectionString = $"Data Source={dbFilePath};Version=3;"
'    End Sub

'    ' Method to open a connection to the SQLite database
'    Private Function OpenConnection() As SQLiteConnection
'        Dim connection As New SQLiteConnection(_connectionString)
'        connection.Open()
'        Return connection
'    End Function

'    ' Method to create a table for CCVDataClosedPL and BorderPLines if it doesn't exist
'    Public Sub CreateTables()
'        Using connection As SQLiteConnection = OpenConnection()
'            Dim sqlCreateCCVTable As String = "
'                CREATE TABLE IF NOT EXISTS CCVDataClosedPL (
'                    TabName TEXT PRIMARY KEY,
'                    Name TEXT,
'                    GrupoPointName TEXT,
'                    Layer TEXT,
'                    MinX REAL,
'                    MinY REAL,
'                    MaxX REAL,
'                    MaxY REAL,
'                    StartStation REAL,
'                    EndStation REAL,
'                    Longitud REAL,
'                    Area REAL,
'                    Side TEXT,
'                    AlignmentHDI TEXT
'                );"

'            Dim sqlCreateBorderTable As String = "
'                CREATE TABLE IF NOT EXISTS BorderPLines (
'                    TabName TEXT PRIMARY KEY,
'                    Layer TEXT,
'                    MinX REAL,
'                    MinY REAL,
'                    MaxX REAL,
'                    MaxY REAL,
'                    StartStation REAL,
'                    EndStation REAL,
'                    Longitud REAL,
'                    Area REAL,
'                    Side TEXT,
'                    AlignmentHDI TEXT
'                );"

'            Using command As New SQLiteCommand(sqlCreateCCVTable, connection)
'                command.ExecuteNonQuery()
'            End Using
'            Using command As New SQLiteCommand(sqlCreateBorderTable, connection)
'                command.ExecuteNonQuery()
'            End Using
'        End Using
'    End Sub

'    ' Method to insert CCVDataClosedPL into the database
'    Public Sub InsertCCVData(closedPL As CCVDataClosedPL)
'        Using connection As SQLiteConnection = OpenConnection()
'            Dim sql As String = "
'                INSERT OR REPLACE INTO CCVDataClosedPL (
'                    TabName, Name, GrupoPointName, Layer, MinX, MinY, MaxX, MaxY,
'                    StartStation, EndStation, Longitud, Area, Side, AlignmentHDI)
'                VALUES (@TabName, @Name, @GrupoPointName, @Layer, @MinX, @MinY, @MaxX, @MaxY,
'                    @StartStation, @EndStation, @Longitud, @Area, @Side, @AlignmentHDI);"

'            Using command As New SQLiteCommand(sql, connection)
'                command.Parameters.AddWithValue("@TabName", closedPL.TabName)
'                command.Parameters.AddWithValue("@Name", closedPL.Name)
'                command.Parameters.AddWithValue("@GrupoPointName", closedPL.GrupoPointName)
'                command.Parameters.AddWithValue("@Layer", closedPL.Layer)
'                command.Parameters.AddWithValue("@MinX", closedPL.MinX)
'                command.Parameters.AddWithValue("@MinY", closedPL.MinY)
'                command.Parameters.AddWithValue("@MaxX", closedPL.MaxX)
'                command.Parameters.AddWithValue("@MaxY", closedPL.MaxY)
'                command.Parameters.AddWithValue("@StartStation", closedPL.StartStation)
'                command.Parameters.AddWithValue("@EndStation", closedPL.EndStation)
'                command.Parameters.AddWithValue("@Longitud", closedPL.Longitud)
'                command.Parameters.AddWithValue("@Area", closedPL.Area)
'                command.Parameters.AddWithValue("@Side", closedPL.Side)
'                command.Parameters.AddWithValue("@AlignmentHDI", closedPL.AlignmentHDI)

'                command.ExecuteNonQuery()
'            End Using
'        End Using
'    End Sub

'    ' Method to insert BorderPLines into the database
'    Public Sub InsertBorderData(borderPL As BorderPLines)
'        Using connection As SQLiteConnection = OpenConnection()
'            Dim sql As String = "
'                INSERT OR REPLACE INTO BorderPLines (
'                    TabName, Layer, MinX, MinY, MaxX, MaxY, StartStation, EndStation,
'                    Longitud, Area, Side, AlignmentHDI)
'                VALUES (@TabName, @Layer, @MinX, @MinY, @MaxX, @MaxY, @StartStation,
'                    @EndStation, @Longitud, @Area, @Side, @AlignmentHDI);"

'            Using command As New SQLiteCommand(sql, connection)
'                command.Parameters.AddWithValue("@TabName", borderPL.TabName)
'                command.Parameters.AddWithValue("@Layer", borderPL.Layer)
'                command.Parameters.AddWithValue("@MinX", borderPL.MinX)
'                command.Parameters.AddWithValue("@MinY", borderPL.MinY)
'                command.Parameters.AddWithValue("@MaxX", borderPL.MaxX)
'                command.Parameters.AddWithValue("@MaxY", borderPL.MaxY)
'                command.Parameters.AddWithValue("@StartStation", borderPL.StartStation)
'                command.Parameters.AddWithValue("@EndStation", borderPL.EndStation)
'                command.Parameters.AddWithValue("@Longitud", borderPL.Longitud)
'                command.Parameters.AddWithValue("@Area", borderPL.Area)
'                command.Parameters.AddWithValue("@Side", borderPL.Side)
'                command.Parameters.AddWithValue("@AlignmentHDI", borderPL.AlignmentHDI)

'                command.ExecuteNonQuery()
'            End Using
'        End Using
'    End Sub

'    ' Method to retrieve CCVDataClosedPL from the database by TabName
'    Public Function GetCCVDataByHandle(handle As String) As CCVDataClosedPL
'        Using connection As SQLiteConnection = OpenConnection()
'            Dim sql As String = "SELECT * FROM CCVDataClosedPL WHERE TabName = @TabName;"

'            Using command As New SQLiteCommand(sql, connection)
'                command.Parameters.AddWithValue("@TabName", handle)

'                Using reader As SQLiteDataReader = command.ExecuteReader()
'                    If reader.Read() Then
'                        Dim closedPL As New CCVDataClosedPL() With {
'                            .TabName = reader("TabName").ToString(),
'                            .Name = reader("Name").ToString(),
'                            .GrupoPointName = reader("GrupoPointName").ToString(),
'                            .Layer = reader("Layer").ToString(),
'                            .MinX = Convert.ToDouble(reader("MinX")),
'                            .MinY = Convert.ToDouble(reader("MinY")),
'                            .MaxX = Convert.ToDouble(reader("MaxX")),
'                            .MaxY = Convert.ToDouble(reader("MaxY")),
'                            .StartStation = Convert.ToDouble(reader("StartStation")),
'                            .EndStation = Convert.ToDouble(reader("EndStation")),
'                            .Longitud = Convert.ToDouble(reader("Longitud")),
'                            .Area = Convert.ToDouble(reader("Area")),
'                            .Side = reader("Side").ToString(),
'                            .AlignmentHDI = reader("AlignmentHDI").ToString()
'                        }
'                        Return closedPL
'                    End If
'                End Using
'            End Using
'        End Using
'        Return Nothing
'    End Function

'    ' Additional methods for updates, deletes, and more complex queries can be added here
'End Class
