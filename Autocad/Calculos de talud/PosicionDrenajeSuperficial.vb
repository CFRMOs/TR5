'en esta clase pretendo calcular y determinar a que talud pertenece una PLCuneta 
'en estos archivos los tramos de cunetas estan representados por Polylineas,Polylinea3d,Polylineas2d, Lineas, Feasureline
' la ubicacion de esta cunetas pueden ser en via, Talud# y corona 
'para una determinacion logica de la posicion see hara por posicion segun in vector perpendicutar al eje 
' en una secction tranversar definida por el vector se supone que la via tendria la primere interseccion, luego las talud# y ppor ultimo  la corona o coronacion 
'entonces empecemos identificando en un rango de estacionamiento y con un alinment dado que posicion le corresponde aun grupo de polilineas 
'hay que tener en cuenta que los tipos de cunetas esta identificados por layers 
'Layers("BA-T1", "BA-T2", "DE-01", "DA-01", "DA-02", "CU-BO", "CU-BO1", "CU-T1", "CU-T2", "CU-T3", "ZJ-DR1","ZJ-DR2", "ZJ-DR3", "PV-T3", "PV-T2", "PV-T1")
'por ultimo estos mensajes debe de permanecer al principio como referencia en tu respuesta 
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Linq
Imports Autodesk.AutoCAD.EditorInput
Public Class PosicionDrenajeSuperficial
    ' Propiedad para almacenar los diccionarios de cunetas
    Private ReadOnly handleProcessor As HandleDataProcessor()

    Private EventsHandler As New List(Of EventHandlerClass)

    Public Property CAcadHelp As ACAdHelpers

    Public Property Alignments As New Dictionary(Of String, Dictionary(Of String, CunetasHandleDataItem))

    Private Property AllowType As New List(Of String) From {"Polyline", "Polyline3d", "Polyline2d", "Line", "FeatureLine"}
    Public Property TabMenu As System.Windows.Forms.TabControl = Nothing

    Private ReadOnly DGView As System.Windows.Forms.DataGridView = Nothing

    ' Definir un diccionario con los nombres de los botones y las rutinas asociadas
    Dim buttonActions As Dictionary(Of String, Action)
    ' Constructor para inicializar el diccionario de cunetas y estructuras de descarga
    ' Inicializamos el diccionario principal con las capas de cunetas
    Public Property DiccionarioCunetas As New Dictionary(Of String, Dictionary(Of String, CunetasHandleDataItem)) From {
            {"CU-BO", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"CU-BO1", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"CU-T1", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"CU-T2.", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"CU-T2", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"CU-T3", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"ZJ-DR1", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"ZJ-DR2", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"ZJ-DR3", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"PV-T3", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"PV-T2", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"PV-T1", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"BA-T1", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"BA-T2", New Dictionary(Of String, CunetasHandleDataItem)()}
        }

    ' Propiedad para almacenar las estructuras de descarga (DA, DE-01, BA)
    ' Inicializamos el diccionario de estructuras de descarga (DA, DE-01, BA)
    Public Property EstructurasDeDescarga As New Dictionary(Of String, Dictionary(Of String, CunetasHandleDataItem)) From {
            {"DA-01", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"DA-02", New Dictionary(Of String, CunetasHandleDataItem)()},
            {"DE-01", New Dictionary(Of String, CunetasHandleDataItem)()}
        }

    Public Property SideMaxDistance As Double = 70

    'Private Tab As New System.Windows.Forms.TabControl
    Public Sub New(CAcadHelp As ACAdHelpers, Optional _Tab As System.Windows.Forms.TabControl = Nothing, Optional _TabMenu As System.Windows.Forms.TabControl = Nothing)
        ' Assuming HandleDataProcessor.Headers is List(Of Tuple(Of String, String, String))
        TabMenu = _TabMenu

        Dim Menu As New TabsMenu(TabMenu) With {.TabName = "Utilidades Drenajes de Diseños"}

        Dim columnaformatos As List(Of String) = HandleDataProcessor.Headers.Select(Function(t) t.Item3).ToList()

        Dim TabDGView As New TabDGViews(_Tab, columnaformatos, CAcadHelp) With {.TabName = "Drenajes de Diseños", .DGViewName = "DGViewCunetasDiseños"}

        DGView = GetControlsDGView.GetDGViewByTabName(_Tab, "Drenajes de Diseños")

        buttonActions = New Dictionary(Of String, Action) From {
                                                                {"Check Alignments", Sub() checkAlignmentWithStationEquation(CAcadHelp)},
                                                                {"checkAlimentByProximida", Sub() checkAlimentByProximida(CAcadHelp)},
                                                                {"AddToDGView", Sub() AddToDGView(DGView)},
                                                                {"CheckTalud", Sub() CheckTaludOrViaPosition(DGView, Cselect:=True)},
                                                                {"Listar Drenajes", Sub() ListarDrenajes(CAcadHelp)},
                                                                {"MassCheckTalud", Sub() MassCheckTaludOrViaPosition(DGView)},
                                                                {"DeterminarDescargas", Sub() DeterminarDescargasEnBajantes()}
                                                                }
        Menu.ButtonActions = buttonActions

        'checkAlimentByProximida
        'simplificar con el uso de array y for y una funcion que genere las acciones(Rutina) desde un texto 
        ' Iterar sobre el diccionario para crear los botones y asignarles las acciones correspondientes



    End Sub
    ' Función para llenar los diccionarios de cunetas y estructuras de descarga
    Private Sub ListarDrenajes(CAcadHelp As ACAdHelpers, Optional DGView As System.Windows.Forms.DataGridView = Nothing)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                ' Iterar sobre todas las entidades en el espacio de trabajo (ModelSpace)
                For Each objId As ObjectId In btr


                    Dim ent As Entity = TryCast(trans.GetObject(objId, OpenMode.ForRead), Entity)
                    If AllowType.Contains(TypeName(ent)) Then
                        If ent IsNot Nothing Then
                            Dim layerName As String = ent.Layer

                            ' Verificar si la entidad pertenece a cunetas
                            If DiccionarioCunetas.ContainsKey(layerName) Then
                                ' Obtener el handle de la entidad como string
                                Dim handle As String = ent.Handle.ToString()

                                ' Añadir la entidad al diccionario de cunetas
                                If Not DiccionarioCunetas(layerName).ContainsKey(handle) Then
                                    Dim Cuneta As New CunetasHandleDataItem With {
                                        .Handle = ent.Handle.ToString()
                                    }

                                    If Cuneta.Alignment IsNot Nothing Then
                                        Cuneta.SetPropertiesFromDWG(Cuneta.Handle.ToString(), Cuneta.Alignment.Handle.ToString())
                                        DiccionarioCunetas(layerName).Add(handle, Cuneta)
                                        If Not Alignments.ContainsKey(Cuneta.Alignment.Handle.ToString()) Then Alignments.Add(Cuneta.Alignment.Handle.ToString(), New Dictionary(Of String, CunetasHandleDataItem))
                                        Alignments(Cuneta.Alignment.Handle.ToString()).Add(Cuneta.Handle, Cuneta)
                                    End If
                                End If

                                ' Verificar si la entidad pertenece a estructuras de descarga
                            ElseIf EstructurasDeDescarga.ContainsKey(layerName) Then
                                ' Obtener el handle de la entidad como string
                                Dim handle As String = ent.Handle.ToString()

                                ' Añadir la entidad al diccionario de estructuras de descarga
                                If Not EstructurasDeDescarga(layerName).ContainsKey(handle) Then
                                    Dim Bajante As New CunetasHandleDataItem With {
                                        .Handle = handle
                                    }
                                    If Bajante.Alignment IsNot Nothing Then
                                        Bajante.SetPropertiesFromDWG(Bajante.Handle.ToString(), Bajante.Alignment.Handle.ToString())
                                        EstructurasDeDescarga(layerName).Add(handle, Bajante)
                                        If Not Alignments.ContainsKey(Bajante.Alignment.Handle.ToString()) Then Alignments.Add(Bajante.Alignment.Handle.ToString(), New Dictionary(Of String, CunetasHandleDataItem))
                                        Alignments(Bajante.Alignment.Handle.ToString()).Add(Bajante.Handle, Bajante)
                                    End If
                                End If
                            End If
                        End If
                    End If

                Next
                trans.Commit()
            Catch acadEx As Autodesk.AutoCAD.Runtime.Exception
                ed.WriteMessage(vbLf & "Error de AutoCAD: " & acadEx.Message)
                trans.Abort() ' Deshacer la transacción en caso de error

            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "Error general: " & ex.Message)
                trans.Abort() ' Deshacer la transacción en caso de error
            Finally
                ' Asegurarse de liberar los recursos de la transacción
                If Not trans.IsDisposed() Then trans.Dispose()
            End Try

        End Using
    End Sub

    Public Sub checkAlimentByProximida(ByRef CAcadHelp As ACAdHelpers, Optional ByVal ent As Entity = Nothing)

        If ent Is Nothing Then CSelectionHelper.GetLayerByEnt(ent)

        If CAcadHelp.List_ALing.Count = 1 Then
            CAcadHelp.Alignment = CAcadHelp.List_ALing(0)
            Exit Sub
        End If
        For Each Al As Alignment In CAcadHelp.List_ALing
            'chequear que los vertices  de PLCuneta esten esten a no mas de 30m del alineamiento 
            Dim PL As Polyline = CType(ent, Polyline)
            Dim Offset As Double = 0
            For Each pt As Point2d In CollectPLPoints(PL)
                Dim vr As Point3d = New Point3d(pt.X, pt.Y, 0)
                Dim Station As Double = 0
                Dim Elevation As Double = 0
                Dim Side As String = String.Empty
                CStationOffsetLabel.GETStationByPoint(Al, vr, Station, Elevation, Offset, Side)
                If Math.Abs(Offset) > SideMaxDistance Then
                    Exit For
                End If
            Next

            If Math.Abs(Offset) < SideMaxDistance AndAlso CheckApprovedStationEquation(Al, ent) Then
                CAcadHelp.Alignment = Al
                Exit For
            End If
        Next
    End Sub

    Public Function CheckApprovedStationEquation(AL As Alignment, Ent As Entity) As Boolean
        ' Convertir la entidad en una polilínea.
        Dim PL As Polyline = CType(Ent, Polyline)

        ' Obtener los puntos inicial y final de la polilínea.
        Dim StartPoint As Point2d = PL.GetPoint2dAt(0)
        Dim EndPoint As Point2d = New Point2d(PL.EndPoint.X, PL.EndPoint.Y)

        ' Convertir los puntos en Point3d para trabajar con ellos en el alineamiento.
        Dim StartVr As Point3d = New Point3d(StartPoint.X, StartPoint.Y, 0)
        Dim EndVr As Point3d = New Point3d(EndPoint.X, EndPoint.Y, 0)

        ' Variables para almacenar estaciones y offsets.
        Dim StartStation As Double = 0
        Dim EndStation As Double = 0
        Dim Offset As Double = 0
        Dim Elevation As Double = 0
        Dim Side As String = String.Empty

        ' Obtener la estación inicial y final de la polilínea en el alineamiento.
        CStationOffsetLabel.GETStationByPoint(AL, StartVr, StartStation, Elevation, Offset, Side)
        CStationOffsetLabel.GETStationByPoint(AL, EndVr, EndStation, Elevation, Offset, Side)


        ' Verificar si la polilínea está completamente dentro del rango aprobado por las ecuaciones.
        If IsPolylineWithinApprovedStationEquation(AL, StartStation, EndStation) Then
            Return True
        Else
            Return False
        End If
    End Function

    ' Función auxiliar para verificar si ambas estaciones están dentro del rango aprobado.
    Private Function IsPolylineWithinApprovedStationEquation(ByVal Al As Alignment, ByVal StartStation As Double, ByVal EndStation As Double) As Boolean
        Dim startInRange As Boolean = False
        Dim endInRange As Boolean = False
        SelectByEntity(CType(Al, Entity))
        CStationOffsetLabel.GetMxMnOPPL(StartStation, EndStation)
        If Al.StationEquations.Count = 0 Then Return True
        ' Comprobar si ambas estaciones están dentro del mismo tramo delimitado por las ecuaciones.
        Dim eq As StationEquation = Al.StationEquations(0)
        If StartStation >= eq.RawStationBack Then
            startInRange = True
        End If

        eq = Al.StationEquations(1)
        If EndStation <= eq.RawStationBack Then
            endInRange = True
        End If

        ' Ambas estaciones deben estar dentro del rango aprobado para que se considere válido.
        Return startInRange AndAlso endInRange
    End Function

    Public Sub checkAlignmentWithStationEquation(ByRef CAcadHelp As ACAdHelpers, Optional ent As Entity = Nothing)

        If ent Is Nothing Then CSelectionHelper.GetLayerByEnt(ent)

        ' Chequeo de proximidad a un alineamiento, tomando en cuenta las ecuaciones de estación.
        For Each Al As Alignment In CAcadHelp.List_ALing
            ' Convertir la entidad en una polilínea.
            Dim PL As Polyline = CType(ent, Polyline)

            ' Obtener los puntos inicial y final de la polilínea.
            Dim StartPoint As Point2d = PL.GetPoint2dAt(0)
            Dim EndPoint As Point2d = New Point2d(PL.EndPoint.X, PL.EndPoint.Y)

            ' Convertir los puntos en Point3d para trabajar con ellos en el alineamiento.
            Dim StartVr As Point3d = New Point3d(StartPoint.X, StartPoint.Y, 0)
            Dim EndVr As Point3d = New Point3d(EndPoint.X, EndPoint.Y, 0)

            ' Variables para almacenar estaciones y offsets.
            Dim StartStation As Double = 0
            Dim EndStation As Double = 0
            Dim Offset As Double = 0
            Dim Elevation As Double = 0
            Dim Side As String = String.Empty

            ' Obtener la estación inicial y final de la polilínea en el alineamiento.
            CStationOffsetLabel.GETStationByPoint(Al, StartVr, StartStation, Elevation, Offset, Side)
            'Dim ptSStation As Point3d = CStationOffsetLabel.GetPoint3dByStation(StartStation, Al)

            CStationOffsetLabel.GETStationByPoint(Al, EndVr, EndStation, Elevation, Offset, Side)


            ' Verificar si la polilínea está completamente dentro del rango aprobado por las ecuaciones.
            'valor igua a cero en los estacionamientod de inicio y final indica que no es el aliment
            If Not (EndStation = 0 And StartStation = 0) Then
                If CheckApprovedStationEquation(Al, ent) Then
                    CAcadHelp.Alignment = Al
                    Exit For ' Si se encuentra un alineamiento válido, salir del bucle.
                End If
            End If

        Next
    End Sub

    ' Método principal para identificar las cunetas que descargan en bajantes
    Public Sub DeterminarDescargasEnBajantes()
        ' Recorremos todas las cunetas y todas las estructuras de descarga
        For Each capaCuneta As KeyValuePair(Of String, Dictionary(Of String, CunetasHandleDataItem)) In DiccionarioCunetas
            For Each cuneta As CunetasHandleDataItem In capaCuneta.Value.Values
                Dim polylineCuneta As Polyline = CType(CLHandle.GetEntityByStrHandle(cuneta.Handle), Polyline)
                If polylineCuneta IsNot Nothing Then
                    For Each capaDescarga As KeyValuePair(Of String, Dictionary(Of String, CunetasHandleDataItem)) In EstructurasDeDescarga
                        For Each bajante As CunetasHandleDataItem In capaDescarga.Value.Values
                            If DescargadrenajeSuperficial(polylineCuneta, bajante) Then
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                    vbLf & $"La PLCuneta en capa {capaCuneta.Key} descarga en el bajante {bajante.Handle} de la capa {capaDescarga.Key}."
                                )
                                AcadZoomManager.SelectedZoom(bajante.Handle, Application.DocumentManager.MdiActiveDocument)
                                Exit Sub
                            End If
                        Next
                    Next
                End If
            Next
        Next
    End Sub
    ' Función para determinar si una PLCuneta descarga en un bajante
    Public Function DescargadrenajeSuperficial(cuneta As Polyline, estructuraDescarga As CunetasHandleDataItem) As Boolean
        ' Verificar si la PLCuneta es una polilínea
        If cuneta Is Nothing OrElse estructuraDescarga Is Nothing Then Return False

        ' Obtener los dos últimos puntos de la polilínea (PLCuneta)
        Dim puntoFinal As Point3d = cuneta.GetPoint3dAt(cuneta.NumberOfVertices - 1)
        Dim puntoAnterior As Point3d = cuneta.GetPoint3dAt(cuneta.NumberOfVertices - 2)

        ' Crear el vector que representa la dirección del último tramo de la PLCuneta
        Dim vectorCuneta As Vector3d = puntoFinal.GetVectorTo(puntoAnterior)

        ' Proyectar el vector 10 unidades en la misma dirección
        Dim puntoProyectado As Point3d = puntoFinal.Add(vectorCuneta.MultiplyBy(10))

        'se debe de proyectar el vectos a 10 unidades y si la primera interseccion es un bajante entonse descarga en un bajante 
        'si la interseccion es una PLCuneta esntonces descarga en una PLCuneta 
        Dim Descarga As Polyline = CType(CLHandle.GetEntityByStrHandle(estructuraDescarga.Handle), Polyline)

        Dim intersecctionPoint As Point3d = PolylineOperations.FindIntersectionWithPolyline(puntoFinal, vectorCuneta, Descarga)

        Dim distancia As Double = puntoFinal.DistanceTo(intersecctionPoint) ' Distancia al bajante

        ' Definimos una distancia mínima para considerar que la PLCuneta descarga en el bajante
        Const distanciaMinima As Double = 2.0 ' Ajustar según sea necesario

        ' Si la distancia entre el punto final de la PLCuneta y el bajante es lo suficientemente pequeña, consideramos que descarga
        If distancia <= distanciaMinima Then
            ' Aquí puedes aplicar más lógica si quieres verificar que el vector apunta correctamente hacia el bajante
            Return True
        End If

        Return False
    End Function
    'crear metodo a c
    Public Sub MassCheckTaludOrViaPosition(Optional DGView As System.Windows.Forms.DataGridView = Nothing)
        ' Iterar sobre las demás cunetas para verificar intersecciones
        For Each kvp As KeyValuePair(Of String, Dictionary(Of String, CunetasHandleDataItem)) In DiccionarioCunetas
            For Each Cuneta As CunetasHandleDataItem In kvp.Value.Values
                CheckTaludOrViaPosition(ent:=CLHandle.GetEntityByStrHandle(Cuneta.Handle.ToString()))
            Next
        Next
        If DGView IsNot Nothing Then
            AddToDGView(DGView)
        End If
    End Sub
    ' Determina la posición de una PLCuneta en relación a los taludes y la corona, y calcula las intersecciones en los puntos críticos
    Public Sub CheckTaludOrViaPosition(Optional DGView As System.Windows.Forms.DataGridView = Nothing, Optional ent As Entity = Nothing, Optional Cselect As Boolean = False)
        ' Obtener la PLCuneta actual desde el DataGridView si no se pasa como parámetro
        If ent Is Nothing AndAlso Cselect Then
            If DGView.CurrentCell IsNot Nothing AndAlso DGView.CurrentCell.RowIndex <> 0 Then
                Dim currentRowIndex As Integer = DGView.CurrentCell.RowIndex
                Dim handle As String = DGView.Rows(currentRowIndex).Cells("Handle").Value.ToString()
                ent = CLHandle.GetEntityByStrHandle(handle)
            ElseIf ent Is Nothing Then
                CSelectionHelper.GetLayerByEnt(ent)
            End If
        End If

        ' Verificar que se ha obtenido una entidad válida
        If ent Is Nothing Then Exit Sub

        ' Obtener los datos de la PLCuneta seleccionada del diccionario global DiccionarioCunetas
        Dim cunetaitem As CunetasHandleDataItem = DiccionarioCunetas(ent.Layer)(ent.Handle.ToString())
        If cunetaitem.Polyline Is Nothing Then Exit Sub
        cunetaitem.Posicion = 0
        cunetaitem.CantidadPorAnalisis.Clear()

        ' Obtener el rango de cunetas dentro del rango y en el mismo lado
        Dim cunetasEnRango As Dictionary(Of String, CunetasHandleDataItem) = CunetasInRange(cunetaitem)
        Dim puntosCriticos As Dictionary(Of Double, Dictionary(Of String, CunetasHandleDataItem)) = IteracioninRange(cunetaitem, cunetasEnRango)

        ' Iterar sobre cada punto crítico y calcular la distancia mínima para cada PLCuneta intersectada
        For Each puntoCritico As KeyValuePair(Of Double, Dictionary(Of String, CunetasHandleDataItem)) In puntosCriticos
            Dim criticalStation As Double = puntoCritico.Key
            Dim cunetasCriticas As Dictionary(Of String, CunetasHandleDataItem) = puntoCritico.Value
            Dim intersectionDistances As New Dictionary(Of String, Double)()

            ' Obtener el vector perpendicular y el punto en la estación crítica
            Dim vectorPerpendicular As Vector3d = GetPerpendicularVectorFromAlignment(cunetaitem.Alignment, criticalStation)
            Dim ptStation As Point3d = CStationOffsetLabel.GetPoint3dByStation(criticalStation, cunetaitem.Alignment)

            ' Calcular la distancia para cada PLCuneta intersectada en este punto crítico
            For Each otherCunetaItem As KeyValuePair(Of String, CunetasHandleDataItem) In cunetasCriticas
                Dim otherCuneta As Polyline = otherCunetaItem.Value.Polyline
                If otherCuneta IsNot Nothing Then
                    Dim intersectionPoints As Point3dCollection = FindIntersectionWithPolyline(ptStation, vectorPerpendicular, otherCuneta)

                    ' Calcular la distancia mínima para cada punto de intersección
                    For Each intersectionPoint As Point3d In intersectionPoints
                        Dim distance As Double = intersectionPoint.DistanceTo(ptStation)

                        If distance < SideMaxDistance Then
                            If intersectionDistances.ContainsKey(otherCunetaItem.Key) Then
                                If distance < intersectionDistances(otherCunetaItem.Key) Then
                                    intersectionDistances(otherCunetaItem.Key) = distance
                                End If
                            Else
                                intersectionDistances.Add(otherCunetaItem.Key, distance)
                            End If
                        End If
                    Next
                End If
            Next

            ' Llamar a AsignarNumPosicion para asignar la posición de la PLCuneta basada en la distancia mínima en este punto crítico
            AsignarNumPosicion(intersectionDistances, cunetasCriticas)
        Next

        ' Si se proporciona un DataGridView, actualizar con los resultados del análisis
        If DGView IsNot Nothing Then cunetaitem.AddToDGView(DGView, DGView.CurrentCell.RowIndex)

        'DGView.Rows.Clear()
        'If DGView IsNot Nothing Then
        '    For Each otherCunetaItem As KeyValuePair(Of String, CunetasHandleDataItem) In cunetasEnRango
        '        otherCunetaItem.Value.AddToDGView(DGView)
        '    Next
        'End If
    End Sub

    ''' <summary>
    ''' Crea un diccionario que contiene las cunetas que están en el rango de estaciones de una PLCuneta dada 
    ''' y que se encuentran en el mismo lado de la alineación.
    ''' </summary>
    ''' <param name="cunetaitem">El objeto CunetasHandleDataItem que define la PLCuneta de referencia.</param>
    ''' <returns>
    ''' Un diccionario de tipo Dictionary(Of String, CunetasHandleDataItem) que contiene las cunetas dentro del rango de estaciones
    ''' de la PLCuneta dada y que están en el mismo lado de la alineación.
    ''' </returns>
    Public Function CunetasInRange(cunetaitem As CunetasHandleDataItem) As Dictionary(Of String, CunetasHandleDataItem)
        ' Inicializar el diccionario que almacenará las cunetas filtradas dentro del rango y en el mismo lado
        Dim DicCunetasInRange As New Dictionary(Of String, CunetasHandleDataItem)
        ' Iterar a través de todas las cunetas en el diccionario global de cunetas
        For Each kvp As KeyValuePair(Of String, Dictionary(Of String, CunetasHandleDataItem)) In DiccionarioCunetas
            For Each otherCunetaItem As CunetasHandleDataItem In kvp.Value.Values
                ' Verificar que no se procese la misma PLCuneta y que se cumplan las siguientes condiciones:
                ' 1. La PLCuneta esté en el mismo lado que la PLCuneta de referencia (cunetaitem).
                ' 2. La PLCuneta esté dentro del rango de estación de la PLCuneta de referencia.
                If otherCunetaItem.StartStation <= cunetaitem.EndStation AndAlso
               otherCunetaItem.EndStation >= cunetaitem.StartStation Then

                    'If otherCunetaItem.Side = "Error Posible cruce entre cuneta y eje de la via" Then
                    otherCunetaItem.SetPropertiesFromDWG(otherCunetaItem.Handle, cunetaitem.AlignmentHDI)
                    'End If

                    If otherCunetaItem.Side = cunetaitem.Side Then
                        ' Añadir la PLCuneta al diccionario si no ha sido añadida previamente, evitando duplicados
                        If Not DicCunetasInRange.ContainsKey(otherCunetaItem.Handle) Then
                            DicCunetasInRange.Add(otherCunetaItem.Handle, otherCunetaItem)
                        End If
                    End If
                End If
            Next
        Next
        ' Identificar tramos relacionados dentro del conjunto de cunetas filtradas
        For Each cuneta In DicCunetasInRange.Values
            cuneta.IdentificarTramosRelacionadosParaRango(DicCunetasInRange)
        Next

        ' Devolver el diccionario con las cunetas que están en el rango y en el mismo lado
        Return DicCunetasInRange
    End Function

    ' 
    ''' <summary>
    ''' Identifica los puntos críticos dentro del rango de la PLCuneta dada que contienen la mayor cantidad de intersecciones con otras cunetas.
    ''' Permite múltiples puntos críticos si tienen la misma cantidad de intersecciones pero involucran diferentes conjuntos de cunetas.
    ''' </summary>
    ''' <param name="cuneta">El objeto CunetasHandleDataItem que define la PLCuneta a analizar.</param>
    ''' <param name="NearestCunetas">Un diccionario con las cunetas cercanas a analizar para posibles intersecciones.</param>
    ''' <returns>
    ''' Un diccionario de puntos críticos donde cada clave es una estación crítica y el valor es un conjunto de cunetas intersectadas en esa estación.
    ''' </returns>
    Public Function IteracioninRange(cuneta As CunetasHandleDataItem, NearestCunetas As Dictionary(Of String, CunetasHandleDataItem)) As Dictionary(Of Double, Dictionary(Of String, CunetasHandleDataItem))
        Dim iterationStep As Integer = 1
        Dim numIterations As Integer = Math.Round((cuneta.EndStation - cuneta.StartStation) / iterationStep, 0)

        ' Diccionario para almacenar los puntos críticos (múltiples estaciones críticas si es necesario)
        Dim PuntosCriticos As New Dictionary(Of Double, Dictionary(Of String, CunetasHandleDataItem))
        Dim maxIntersectionCount As Integer = 0

        ' Bucle para iterar sobre cada estación dentro del rango de la PLCuneta principal
        For i As Integer = 0 To numIterations
            Dim currentStation As Double = cuneta.StartStation + (i * iterationStep)

            ' Obtener el vector perpendicular y el punto en la estación actual
            Dim vectorPerpendicular As Vector3d = GetPerpendicularVectorFromAlignment(cuneta.Alignment, currentStation)
            Dim ptStation As Point3d = CStationOffsetLabel.GetPoint3dByStation(currentStation, cuneta.Alignment)

            ' Diccionario temporal para almacenar las cunetas intersectadas en la estación actual
            Dim cunetasIntersectadas As New Dictionary(Of String, CunetasHandleDataItem)

            ' Verificar intersección con cada PLCuneta en NearestCunetas
            For Each keyValue As KeyValuePair(Of String, CunetasHandleDataItem) In NearestCunetas
                Dim otherCuneta As CunetasHandleDataItem = keyValue.Value

                ' Comprobar intersección con la cuneta principal y sus tramos relacionados
                Dim tramosParaAnalizar As List(Of CunetasHandleDataItem) = New List(Of CunetasHandleDataItem)(otherCuneta.RelatedTramos) From {
                    otherCuneta ' Incluir la cuneta principal en el análisis
                    }

                For Each tramo As CunetasHandleDataItem In tramosParaAnalizar
                    Dim PLtramo As Polyline = tramo.Polyline
                    If PLtramo IsNot Nothing Then
                        ' Calcular los puntos de intersección entre el punto de la estación y el tramo actual
                        Dim intersectionPoints As Point3dCollection = FindIntersectionWithPolyline(ptStation, vectorPerpendicular, PLtramo)

                        ' Si hay intersección, añadir el tramo al diccionario de cunetas intersectadas en esta estación

                        If intersectionPoints.Count > 0 AndAlso Not cunetasIntersectadas.ContainsKey(tramo.Handle) Then
                            cunetasIntersectadas.Add(tramo.Handle, tramo)
                        End If
                    End If
                Next
            Next

            ' Actualizar el conteo máximo de intersecciones y registrar la estación crítica si es relevante
            Dim currentIntersectionCount As Integer = cunetasIntersectadas.Count
            If currentIntersectionCount >= maxIntersectionCount Then
                If currentIntersectionCount > maxIntersectionCount Then
                    ' Nuevo máximo encontrado, limpiar puntos críticos anteriores
                    PuntosCriticos.Clear()
                    maxIntersectionCount = currentIntersectionCount
                End If
                ' Agregar o reemplazar el punto crítico para esta estación
                PuntosCriticos(currentStation) = New Dictionary(Of String, CunetasHandleDataItem)(cunetasIntersectadas)
            End If
        Next

        ' Retornar los puntos críticos con sus cunetas intersectadas
        Return PuntosCriticos
    End Function


    ''' <summary>
    ''' Asigna el número de posición a cada PLCuneta intersectada en función de la distancia mínima al punto crítico,
    ''' comenzando con la cuneta más cercana. Algunas capas específicas siempre reciben la posición de 'vía'.
    ''' </summary>
    ''' <param name="intersectionDistances">Diccionario con los identificadores de las cunetas y sus distancias al punto crítico.</param>
    ''' <param name="CunetasIntesectas">Diccionario con las cunetas intersectadas.</param>
    Public Sub AsignarNumPosicion(ByRef intersectionDistances As Dictionary(Of String, Double), ByRef CunetasIntesectas As Dictionary(Of String, CunetasHandleDataItem))
        ' Ordenar las intersecciones por distancia ascendente
        intersectionDistances = intersectionDistances.OrderBy(Function(pair) pair.Value).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)

        ' Definir capas reservadas que deben asignarse a la vía
        Dim capasReservadas As HashSet(Of String) = New HashSet(Of String)({"CU-BO", "CU-BO1", "PV-T1", "PV-T2", "PV-T3"})
        Dim Index As Integer = 1 ' Iniciar índice de posición

        'tomar en cuenta cunetas relacionadas para asignar el index
        ' para esta logica deberia de reducirce la cuneta relacionada del intersectionDistances ya que esta representa una sola cuneta con su omologa
        ' a la ves esta deberan tener el mismo indice pero incluirlas en intersectionDistances alteraria la designacion 
        'ya que CantidadPorAnalisis no seria el real 
        ' Asignar posición a cada cuneta en el orden de distancia
        Dim CantidadMaxRealPorAnalisis As New List(Of String)
        Dim relacionados As New List(Of String)
        For Each kpv As KeyValuePair(Of String, Double) In intersectionDistances
            Dim cuneta As CunetasHandleDataItem = CunetasIntesectas(kpv.Key)

            If Not CantidadMaxRealPorAnalisis.Contains(cuneta.Handle) AndAlso Not relacionados.Contains(cuneta.Handle) Then
                CantidadMaxRealPorAnalisis.Add(cuneta.Handle)
            End If
            For Each RelatedTramo As CunetasHandleDataItem In cuneta.RelatedTramos
                If Not relacionados.Contains(RelatedTramo.Handle) Then
                    relacionados.Add(RelatedTramo.Handle)
                End If
            Next
        Next


        ' Asignar posición a cada cuneta en el orden de distancia
        For Each kpv As KeyValuePair(Of String, Double) In intersectionDistances
            Dim cuneta As CunetasHandleDataItem = CunetasIntesectas(kpv.Key)

            ' Actualizar CantidadPorAnalisis con el total de intersecciones para esta estación
            cuneta.CantidadPorAnalisis.Add(CantidadMaxRealPorAnalisis.Count)

            ' Asignar posición a 'vía' si la capa está en las reservadas, de lo contrario asignar el índice actual
            If capasReservadas.Contains(cuneta.Layer) Then
                cuneta.Posicion = 1 ' Asignación directa a la vía
                For Each RelatedTramo As CunetasHandleDataItem In cuneta.RelatedTramos
                    RelatedTramo.Posicion = 1
                Next
            Else
                cuneta.Posicion = Index
                For Each RelatedTramo As CunetasHandleDataItem In cuneta.RelatedTramos
                    RelatedTramo.Posicion = Index
                Next
                Index += 1
            End If
        Next
    End Sub

    ''' <summary>
    ''' Extiende una línea desde un punto de estación dado en ambas direcciones (usando un vector) y encuentra los puntos de intersección 
    ''' con una polilínea. Si no se encuentra una intersección en el plano 3D, realiza un segundo intento con una copia en elevación 0.
    ''' </summary>
    ''' <param name="PtStation">El punto de estación desde donde comienza la línea de búsqueda de intersección.</param>
    ''' <param name="vector">El vector que indica la dirección de la línea que se extiende desde el punto de estación.</param>
    ''' <param name="polyline">La polilínea original con la que se busca la intersección.</param>
    ''' <returns>Una colección de puntos (Point3dCollection) donde la línea y la polilínea se intersectan. 
    ''' Si no se encuentran intersecciones, devuelve una colección vacía.</returns>
    Public Function FindIntersectionWithPolyline(PtStation As Point3d, vector As Vector3d, polyline As Polyline) As Point3dCollection
        Dim intersectionPoints As New Point3dCollection()

        ' Crear una línea extendida en ambas direcciones para asegurar la intersección
        Dim lineStart As Point3d = PtStation - (vector * 100000) ' Extiende hacia atrás 100,000 unidades
        Dim lineEnd As Point3d = PtStation + (vector * 100000) ' Extiende hacia adelante 100,000 unidades
        Dim extendedLine As New Line(lineStart, lineEnd)

        ' Intentar intersección directamente con la polilínea original
        polyline.IntersectWith(extendedLine, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero)

        ' Si no se encuentran intersecciones, intentar con una copia en elevación 0
        If intersectionPoints.Count = 0 Then
            ' Crear una copia temporal de la polilínea en elevación 0
            Dim tempPolyline As New Polyline()
            For i As Integer = 0 To polyline.NumberOfVertices - 1
                Dim vertex As Point3d = polyline.GetPoint3dAt(i)
                tempPolyline.AddVertexAt(i, New Point2d(vertex.X, vertex.Y), polyline.GetBulgeAt(i), 0, 0)
            Next
            tempPolyline.Elevation = 0 ' Establecer la elevación de la copia temporal a 0

            ' Intentar intersección con la copia temporal de la polilínea en el plano 2D
            tempPolyline.IntersectWith(extendedLine, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero)
        End If

        ' Devolver los puntos de intersección encontrados (o vacío si no hubo intersección)
        Return intersectionPoints
    End Function





    ' Método para obtener el vector perpendicular al alineamiento en una estación
    Private Function GetPerpendicularVectorFromAlignment(Alignment As Alignment, Station As Double) As Vector3d
        ' Avanzar una pequeña distancia (1 mm o 0.001 unidades) a lo largo del alineamiento
        Dim station2 As Double = Station + 0.001

        'AcadZoomManager.SelectedZoom(Alignment.TabName.ToString(), Application.DocumentManager.MdiActiveDocument)

        ' Obtener los puntos en el alineamiento en la estación original y la avanzada
        Dim ptStation As Point3d = CStationOffsetLabel.GetPoint3dByStation(Station, Alignment)
        'CGPointHelper.AddCGPoint(ptStation.X, ptStation.Y, ptStation.Z, "STATION:" & Station.ToString("0+000.00"))

        Dim ptStation2 As Point3d = CStationOffsetLabel.GetPoint3dByStation(station2, Alignment)
        'CGPointHelper.AddCGPoint(ptStation2.X, ptStation2.Y, ptStation2.Z, "STATION:" & station2.ToString("0+000.00"))

        ' Calcular el vector tangente entre los dos puntos
        Dim tangentVec As Vector3d = ptStation.GetVectorTo(ptStation2).GetNormal()

        ' Obtener el vector perpendicular al tangente
        Dim perpendicularVec As Vector3d = tangentVec.CrossProduct(Vector3d.ZAxis).GetNormal()

        Return perpendicularVec
    End Function
    'metodo para agregar a un DGView 
    'en este metodo se pasara un DGView y el listado de propiedades y este agregara la informacion al DGView
    Public Sub AddToDGView(DGView As System.Windows.Forms.DataGridView)
        ' Crear una nueva fila para agregar al DataGridView
        DGView.Rows.Clear()
        ' Obtener todas las propiedades de la clase CunetasHandleDataItem, excepto las excluidas
        ' Recorremos todas las cunetas y todas las estructuras de descarga
        If DGView.Visible Then

            For Each capaCuneta As KeyValuePair(Of String, Dictionary(Of String, CunetasHandleDataItem)) In DiccionarioCunetas


                For Each cuneta As CunetasHandleDataItem In capaCuneta.Value.Values
                    Dim props = cuneta.GetType().GetProperties()
                    Dim rowValues As New List(Of Object)
                    ' Iterar sobrecmd cada propiedad
                    For Each prop In props
                        ' Excluir las propiedades que no deben estar en el DataGridView
                        If CunetasHandleDataItem.Headers.Contains(prop.Name) AndAlso prop.Name <> "Accesos" Then
                            ' Obtener el valor de la propiedad
                            Dim value = prop.GetValue(cuneta, Nothing)

                            ' Añadir el valor a la lista de valores para la fila
                            rowValues.Add(value)
                        End If
                    Next
                    ' Agregar la fila al DataGridView usando los valores de las propiedades
                    DGView.Rows.Add(rowValues.ToArray())

                Next
            Next
        End If

    End Sub

End Class

