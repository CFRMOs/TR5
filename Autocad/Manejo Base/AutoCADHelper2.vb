' todos los comentarios hechos por mi deberan de permanece
' en esta clase no se manejaran librerias com de autocad. esos proceso se maneja en clase dedicadas para ese objetivo 
' devuelme siempre la actualizacion de esta pagina competa 
' esta clase es temporal para ver tus sugerencias y no dañar loa que tengo en mi codigo 
' Clase temporal para propósitos de prueba
Imports System.Diagnostics
Imports System.Linq
Imports System.Windows.Documents
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Class AutoCADHelper2
    ' DEVOLVERA EL DICCIONARIO Dictionary(Of String, CunetasHandleDataItem) DONDE
    ' Object: FILA DE LA TABLA CORRESPONDIENTE AL HANDLE 
    ' String: HANDLE AS STRING
    Public Shared Function ProcesarHandlesDesdeDoc(docName As String, tableName As String, Optional WBFilePath As String = vbNullString) ', Optional ByRef SQLmase As DataBSQLManager = Nothing) As Dictionary(Of String, CunetasHandleDataItem)
        ' Si no se encuentra el código, salir
        Dim RefAcceso As String = AccesoNum(docName)
        If String.IsNullOrEmpty(RefAcceso) Then
            Console.WriteLine("El archivo no presenta referencia de accesos.")
            Return Nothing
        End If

        ' Obtener el número de acceso a partir de la referencia extraída
        Dim ACCESO As Double = CDbl(CodigoExtractor.ExtraerCodigo(RefAcceso, "\d+"))

        ' Verificar si se ha encontrado un código de acceso válido
        If ACCESO = 0 Then
            Throw New Exception("No se pudo extraer el código de acceso del nombre del archivo.")
        End If

        ' Obtener los datos de handle desde el archivo Excel
        Dim hndleData As Dictionary(Of String, CunetasHandleDataItem) = ExcelHelper.ObtenerDatosTabla2(tableName, ACCESO, WBFilePath)

        'crear proceso para base de datos 

        ' Devolver el diccionario de handles procesados
        Return hndleData
    End Function

    'en los nombres de los archivos se codifica y mediante a esta funccion obtenesmos el codigo representativo del acceso 
    Public Shared Function AccesoNum(Text As String) As String
        ' Buscar información del acceso en el nombre del archivo 
        Dim RefAcceso As String = String.Empty
        For Each Cod As String In {"ACC-\d+", "ACC\d+", "AC-\d+"}
            If String.IsNullOrEmpty(RefAcceso) Then
                RefAcceso = CodigoExtractor.ExtraerCodigo(Text, Cod)
            Else
                Exit For
            End If
        Next
        Return RefAcceso
    End Function

    ' Se verifica la existencia de los handles en hndleData separando lo existentes de lo no existentes
    Public Shared Sub VerificarHandles(hndleData As Dictionary(Of String, CunetasHandleDataItem),
                                       ByRef handlesEncontrados As Dictionary(Of String, CunetasHandleDataItem),
                                       ByRef handlesNoEncontrados As Dictionary(Of String, CunetasHandleDataItem))

        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                For Each kvp As KeyValuePair(Of String, CunetasHandleDataItem) In hndleData
                    Dim HandleDataItem As CunetasHandleDataItem = kvp.Value

                    Try
                        ' Intentar convertir el TabName a un valor hexadecimal
                        Dim handleObj As New Handle(Convert.ToInt64(HandleDataItem.Handle, 16))
                        Dim objId As ObjectId = db.GetObjectId(False, handleObj, 0)

                        ' Verificar si el ObjectId es válido
                        If objId.IsValid Then
                            If Not handlesEncontrados.ContainsKey(HandleDataItem.Handle) Then
                                handlesEncontrados.Add(HandleDataItem.Handle, HandleDataItem)
                                Dim CunetaEnt As Entity = TryCast(trans.GetObject(CLHandle.GetEntityIdByStrHandle(HandleDataItem.Handle), OpenMode.ForWrite), Entity)
                                If CunetaEnt.Layer <> HandleDataItem.Layer Then
                                    If HandleDataItem.Layer Is Nothing Then HandleDataItem.SyncPropExcel({"Layer"}, "Set")
                                    '    CunetaEnt.Layer = HandleDataItem.layer
                                    HandleDataItem.SyncPropExcel({"Layer"}, "Get")
                                End If
                                If Trim(HandleDataItem.FileName) <> Trim(Right(doc.Name, doc.Name.Length - InStrRev(doc.Name, "\"))) Then
                                    HandleDataItem.FileName = Trim(Right(doc.Name, doc.Name.Length - InStrRev(doc.Name, "\")))
                                    HandleDataItem.FilePath = doc.Name
                                    HandleDataItem.SyncPropExcel({"FileName", "FilePath"}, "Get")
                                End If
                            End If
                        Else
                            If Not handlesNoEncontrados.ContainsKey(HandleDataItem.Handle) Then
                                handlesNoEncontrados.Add(HandleDataItem.Handle, HandleDataItem)
                            End If
                        End If
                    Catch ex As Exception
                        ' Si falla la conversión o no se puede obtener el ObjectId, agregar a no encontrados
                        Console.WriteLine($"Error procesando el handle {HandleDataItem.Handle}: {ex.Message}")
                        If Not handlesNoEncontrados.ContainsKey(HandleDataItem.Handle) Then
                            handlesNoEncontrados.Add(HandleDataItem.Handle, HandleDataItem)
                        End If
                        'trans.Abort()
                    End Try
                Next

                ' Confirmar la transacción
                trans.Commit()

            Catch ex As Exception
                ' En caso de error, imprimir el mensaje y abortar la transacción
                Console.WriteLine("Error al procesar los handles: " & ex.Message)
                trans.Abort()
            Finally
                ' Liberar la transacción
                If Not trans.IsDisposed Then trans.Dispose()
            End Try
        End Using
    End Sub

    ' Se pretende obtener un listado relacionado de CunetasHandleDataItem con relación a la longitud dada de un CunetasHandleDataItem no existente (handlesNoEncontrados)
    ' Devuelve: Dictionary(Of String, (CunetasHandleDataItem, Dictionary(Of String, CunetasHandleDataItem))):
    ' - item1: handle del dato de la tabla que se ha comprobado su no existencia
    ' - item2: CunetasHandleDataItem con la información del handle
    ' - item3: Diccionario con las entidades del archivo DWG que cumplen con las condiciones
    Public Shared Sub LookByLength(ByRef handlesNoEncontrados As Dictionary(Of String, CunetasHandleDataItem),
                                   ByRef HandlesByLength As Dictionary(Of String, CunetasHandleDataItem),
                                   Optional actualizarfexcel As Boolean = False,
                                   Optional ByRef cListType As List(Of List(Of ObjectId)) = Nothing)


        Dim ListType As List(Of List(Of ObjectId))

        ' Listas predefinidas
        If cListType Is Nothing Then
            ListType = ConstructDiccionario(cListType)
        Else
            ListType = cListType
        End If

        Dim handlesToRemove As New List(Of String)

        ' Recorre los elementos de handlesNoEncontrados
        For Each kvp As KeyValuePair(Of String, CunetasHandleDataItem) In handlesNoEncontrados
            Try
                Dim HandleDataItem As CunetasHandleDataItem = kvp.Value

                ' Buscar la entidad con las condiciones
                With HandleDataItem

                    If HandleDataItem.AlignmentHDI <> "" Then Dim Align As Alignment = CType(CLHandle.GetEntityByStrHandle(HandleDataItem.AlignmentHDI), Alignment)

                    'recorer losalineamientos 
                    ' Buscar en los diferentes tipos de polilíneas
                    Dim AligmentsCertify As New AlignmentHelper()

                    For Each iTypes As List(Of ObjectId) In ListType ', listPolyline2d, listPolyline3d, listLines, listFeatureLines}


                        Dim AcEnt As Entity = FindPolylineByLength(.Longitud, 2, 2, .StartStation, .EndStation, iTypes, HandleDataItem.Side)

                        If AcEnt IsNot Nothing Then
                            Dim PL As Polyline = Nothing
                            If TypeOf AcEnt Is FeatureLine Then
                                PL = FeaturelineToPolyline.ConvertFeaturelineToPolyline(AcEnt, True)
                            ElseIf TypeOf AcEnt Is Polyline Then
                                PL = TryCast(AcEnt, Polyline)
                            End If
                            If String.IsNullOrEmpty(.AlignmentHDI) AndAlso PL IsNot Nothing Then
                                'como la linea no existe en la base de datos esta no optendra el alignment mendiante la clase handle
                                ' Añadir el handle a la lista de eliminación

                                handlesToRemove.Add(kvp.Value.Handle)

                                .Handle = PL.Handle.ToString()

                            End If
                            ' Actualizar propiedades si es necesario
                            If actualizarfexcel Then
                                If HandleDataItem.AlignmentHDI Is Nothing Then HandleDataItem.SyncPropExcel({"AlignmentHDI"}, "Set")

                                If HandleDataItem.Handle <> PL?.Handle.ToString() Then
                                    HandleDataItem.Handle = PL?.Handle.ToString()
                                End If
                            End If
                            ' Actualizar el diccionario HandlesByLength
                            HandleDataItem.SetPropertiesFromDWG(HandleDataItem.Handle, HandleDataItem.AlignmentHDI)
                            HandlesByLength.Add(HandleDataItem.Handle, HandleDataItem)
                            Exit For ' Salir del bucle interno si encontramos una coincidencia
                        End If
                    Next
                End With
            Catch ex As Exception
                Debug.WriteLine("error:" & ex.Message())
            End Try
        Next

        ' Después de recorrer el diccionario, remover los elementos encontrados
        For Each handle As String In handlesToRemove
            If handle Is Nothing Then handle = ""
            If handlesNoEncontrados IsNot Nothing AndAlso handlesNoEncontrados.ContainsKey(handle) Then handlesNoEncontrados.Remove(handle)
        Next
    End Sub

    Public Shared Function ConstructDiccionario(Optional ByRef cListType As List(Of List(Of ObjectId)) = Nothing) As List(Of List(Of ObjectId))

        ' Listas predefinidas
        If cListType IsNot Nothing AndAlso cListType.Count = 0 Then
            Dim listPolylines As List(Of ObjectId) = ConstructListType(Of Polyline)()
            cListType.Add(listPolylines)
            Dim listFeatureLines As List(Of ObjectId) = ConstructListType(Of FeatureLine)()
            cListType.Add(listFeatureLines)
            Dim listLines As List(Of ObjectId) = ConstructListType(Of Line)()
            cListType.Add(listLines)
            Dim listPolyline2d As List(Of ObjectId) = ConstructListType(Of Polyline2d)()
            cListType.Add(listPolyline2d)
            Dim listPolyline3d As List(Of ObjectId) = ConstructListType(Of Polyline3d)()
            cListType.Add(listPolyline3d)
        End If
        Return cListType
    End Function

    ' Método para identificar una polilínea que cumple con los criterios dados
    Public Shared Function FindPolylineByLength(longitud As Double, precision As Integer, Tolerancia As Double, EstST As Double, EstEnd As Double,
                                                ByRef listPolylines As List(Of ObjectId), side As String) As Entity ', ByRef listFeactureLines As List(Of ObjectId)) As Entity
        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        ' Obtener la base de datos asociada con el documento activo
        Dim db As Database = doc.Database
        ' Obtener el editor del documento (usado para mensajes de usuario y otras interacciones)
        Dim ed As Editor = doc.Editor

        ' Iniciar una transacción para leer y modificar la base de datos de AutoCAD
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Using doc.LockDocument
                Try
                    ' Abrir la tabla de bloques en modo lectura
                    Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    ' Abrir el registro del espacio del modelo en modo lectura
                    Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                    Dim objectIds As List(Of ObjectId) = CheckEntityByLenght(listPolylines, longitud, precision, Tolerancia)

                    Dim objectIdsByStation As List(Of ObjectId) = CheckEntityByStation(objectIds, longitud, EstST, EstEnd, precision, Tolerancia, side)

                    Dim entity As Entity = Nothing

                    If objectIdsByStation.Count <> 0 Then
                        entity = TryCast(trans.GetObject(objectIdsByStation(0), OpenMode.ForRead), Entity)
                    End If

                    If entity IsNot Nothing Then
                        Return entity
                    Else
                        ' Si no se encuentra ninguna polilínea que cumpla con los criterios, devolver Nothing
                        Return Nothing
                    End If
                Catch ex As Exception
                    ' Escribir el error en la consola de AutoCAD
                    ed.WriteMessage("Error en ConstructListType: " & ex.Message)
                    trans.Abort()
                    Return Nothing
                Finally
                    ' Asegurarse de que la transacción se elimine correctamente
                    If trans.IsDisposed = False Then trans.Dispose()
                End Try
            End Using
        End Using
    End Function

    Public Shared Function CheckEntityByStations(ByRef List As List(Of ObjectId), precision As Integer, Tolerancia As Double, EstST As Double, EstEnd As Double, alignment As Alignment) As Entity
        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        ' Obtener la base de datos asociada con el documento activo
        Dim db As Database = doc.Database
        ' Obtener el editor del documento (usado para mensajes de usuario y otras interacciones)
        Dim ed As Editor = doc.Editor

        ' Iniciar una transacción para leer y modificar la base de datos de AutoCAD
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Using doc.LockDocument
                Try
                    ' Abrir la tabla de bloques en modo lectura
                    Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                    ' Abrir el registro del espacio del modelo en modo lectura
                    Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                    Dim entity As Entity = Nothing
                    ' Recorrer cada objeto en el espacio del modelo
                    For Each objId As ObjectId In List

                        ' Obtener la entidad asociada con el ObjectId
                        entity = CType(trans.GetObject(objId, OpenMode.ForRead), Entity)

                        'AcadZoomManager.SelectedZoom(entity.TabName.ToString(), doc)
                        If TypeOf entity IsNot Polyline Then entity = ConcertTempPL(entity)

                        ' Verificar si la entidad es una polilínea

                        ' Convertir la entidad a una polilínea
                        Dim pline As Polyline = CType(entity, Polyline)
                        Dim side As String = String.Empty
                        ' Obtener la estación inicial y final de la polilínea usando el alineamiento proporcionado
                        Dim startStation As Double
                        Dim endStation As Double
                        Dim plineLength As Double
                        ' Verificar si la longitud de la polilínea está dentro del rango de precisión especificado
                        ' Considera un margen de error (+ o -) no mayor ni menor de 2, y acepta solo polilíneas abiertas

                        IdentifyPL.PLRelatedInf(pline, alignment, plineLength, startStation, endStation, side)
                        startStation = Math.Round(startStation, precision)
                        endStation = Math.Round(endStation, precision)

                        ' Verificar si las estaciones están dentro del rango de estaciones especificado
                        If Math.Abs(startStation - EstST) <= Tolerancia AndAlso
                                    Math.Abs(endStation - EstEnd) <= Tolerancia Then
                            ' Si la polilínea cumple con todos los criterios, devolverla
                            List.Remove(objId)
                            Return entity
                            trans.Commit()
                        Else
                            ed.WriteMessage("Errores de estacionamientos:")
                        End If
                    Next
                    trans.Commit()
                    Return Nothing

                Catch ex As Exception
                    ' Escribir el error en la consola de AutoCAD
                    ed.WriteMessage("Error en ConstructListType: " & ex.Message)
                    ' Abortar la transacción si hay una excepción
                    trans.Abort()
                    Return Nothing

                Finally
                    ' Asegurarse de que la transacción se elimine correctamente
                    If trans.IsDisposed = False Then trans.Dispose()
                End Try
            End Using
        End Using
    End Function

    Public Shared Function CheckEntityByStation(ByRef List As List(Of ObjectId), longitud As Double, EstST As Double, EstEnd As Double,
                                            precision As Integer, Tolerancia As Double, cSide As String) As List(Of ObjectId)
        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim Listresult As New List(Of ObjectId)
        Dim ListtoRemove As New List(Of ObjectId)
        Dim listTolerancias As New List(Of ToleranceRecord)

        ' Iniciar una transacción para leer y modificar la base de datos de AutoCAD
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Using doc.LockDocument()
                Try
                    ' Abrir la tabla de bloques en modo lectura
                    Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                    Dim AligmentsCertify As New AlignmentHelper()

                    ' Recorrer cada objeto en el espacio del modelo
                    For Each objId As ObjectId In List
                        ' Obtener la entidad asociada con el ObjectId
                        Dim entity As Entity = CType(trans.GetObject(objId, OpenMode.ForRead), Entity)

                        If TypeOf entity IsNot Polyline Then entity = ConcertTempPL(entity)

                        ' Convertir la entidad a una polilínea
                        Dim pline As Polyline = CType(entity, Polyline)
                        Dim side As String = String.Empty
                        Dim startStation As Double
                        Dim endStation As Double
                        Dim Len As Double

                        ' Obtener la información de la polilínea usando alineamientos
                        For Each alignment As Alignment In AligmentsCertify.List_ALing

                            IdentifyPL.PLRelatedInf(pline, alignment, Len, startStation, endStation, side)
                            startStation = Math.Round(startStation, precision)
                            endStation = Math.Round(endStation, precision)

                            Dim toleranceRecord As New ToleranceRecord(objId, longitud, Len, EstST, startStation, EstEnd, endStation)

                            ' Verificar si las estaciones están dentro del rango de estaciones especificado
                            If toleranceRecord.StartStationTolerance <= Tolerancia AndAlso
                               toleranceRecord.EndStationTolerance <= Tolerancia AndAlso
                               cSide = side Then
                                listTolerancias.Add(toleranceRecord)
                                ListtoRemove.Add(objId)
                                Listresult.Add(objId)
                            End If
                        Next
                    Next

                    For Each ObjID As ObjectId In ListtoRemove
                        List.Remove(ObjID)
                    Next

                    ' Encontrar el objeto con la menor tolerancia total
                    Dim minToleranceRecord = listTolerancias.OrderBy(Function(t) t.TotalTolerance()).FirstOrDefault()

                    ' Devolver solo el objeto con la mejor tolerancia si existe, o todos los que cumplieron las condiciones si no
                    If minToleranceRecord IsNot Nothing Then
                        Return New List(Of ObjectId)({minToleranceRecord.ObjectId})
                    End If


                    trans.Commit()

                    ' Devolver toda la lista de resultados si no hay un mínimo específico
                    Return Listresult

                Catch ex As Exception
                    ' Escribir el error en la consola de AutoCAD
                    ed.WriteMessage("Error en ConstructListType: " & ex.Message)
                    trans.Abort()
                    If Listresult.Count <> 0 Then
                        Return Listresult
                    Else
                        Return Nothing
                    End If

                Finally
                    ' Asegurarse de que la transacción se elimine correctamente
                    If trans.IsDisposed = False Then trans.Dispose()
                End Try
            End Using
        End Using
    End Function
    Public Shared Function CheckEntityByLenght(ByRef List As List(Of ObjectId), longitud As Double, precision As Integer, Tolerancia As Double) As List(Of ObjectId)
        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim Listresult As New List(Of ObjectId)
        Dim ListtoRemove As New List(Of ObjectId)
        ' Iniciar una transacción para leer y modificar la base de datos de AutoCAD
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Using doc.LockDocument()
                Try
                    ' Abrir la tabla de bloques en modo lectura
                    Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    ' Abrir el registro del espacio del modelo en modo lectura
                    Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                    ' Recorrer cada objeto en el espacio del modelo
                    For Each objId As ObjectId In List
                        ' Obtener la entidad asociada con el ObjectId
                        Dim entity As Entity = CType(trans.GetObject(objId, OpenMode.ForRead), Entity)
                        Dim Lenght As Double = GetEntityLength(entity)

                        ' Verificar si la longitud de la polilínea está dentro del rango de precisión especificado
                        ' Considera un margen de error (+ o -) no mayor ni menor de 2, y acepta solo polilíneas abiertas
                        If Math.Abs(Math.Round(Lenght, precision) - Math.Round(longitud, precision)) <= Tolerancia AndAlso IsEntityClosed(entity) = False Then
                            ListtoRemove.Add(objId)
                            Listresult.Add(objId)
                        End If
                    Next
                    'For Each ObjID As ObjectId In ListtoRemove
                    '    List.Remove(ObjID)
                    'Next
                    trans.Commit()
                    Return Listresult

                Catch ex As Exception
                    ' Escribir el error en la consola de AutoCAD
                    ed.WriteMessage("Error en ConstructListType: " & ex.Message)
                    ' Abortar la transacción si hay una excepción
                    trans.Abort()
                    If Listresult.Count <> 0 Then
                        Return Listresult
                    Else
                        Return Nothing
                    End If

                Finally
                    ' Asegurarse de que la transacción se elimine correctamente
                    If trans.IsDisposed = False Then trans.Dispose()
                End Try
            End Using
        End Using
    End Function
    Public Shared Function ConstructListType(Of T As Entity)() As List(Of ObjectId)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        ' Obtener la base de datos asociada con el documento activo
        Dim db As Database = doc.Database
        ' Obtener el editor del documento (usado para mensajes de usuario y otras interacciones)
        Dim ed As Editor = doc.Editor

        ' Lista para almacenar los ObjectId de las entidades del tipo T
        Dim entityList As New List(Of ObjectId)

        ' Iniciar una transacción para leer y modificar la base de datos de AutoCAD
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Abrir la tabla de bloques en modo lectura
                Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                ' Abrir el registro del espacio del modelo en modo lectura
                Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                ' Recorrer cada objeto en el espacio del modelo
                For Each objId As ObjectId In btr
                    ' Obtener la entidad asociada con el ObjectId
                    Dim entity As Entity = CType(trans.GetObject(objId, OpenMode.ForRead), Entity)
                    If IsEntityClosed(entity) = False Then
                        ' Verificar si la entidad es del tipo T (Polyline, Circle, etc.)
                        If TypeOf entity Is T Then
                            entityList.Add(objId)
                        End If
                    End If
                Next
                ' Confirmar la transacción y devolver la lista
                trans.Commit()
                Return entityList

            Catch ex As Exception
                ' Escribir el error en la consola de AutoCAD
                ed.WriteMessage("Error en ConstructListType: " & ex.Message)
                ' Abortar la transacción si hay una excepción
                trans.Abort()
                Return Nothing

            Finally
                ' Asegurarse de que la transacción se elimine correctamente
                If trans.IsDisposed = False Then trans.Dispose()
            End Try
        End Using
    End Function

    ' Función que verifica si una entidad está cerrada (retorna True si está cerrada, False si está abierta)
    ' Revisa el tipo de entidad y retorna su estado de cerrado si existe una propiedad que lo determine
    Public Shared Function IsEntityClosed(entity As Entity) As Boolean
        Select Case True
        ' Caso para polilíneas 2D: verifica la propiedad "Closed" de las polilíneas en 2D
            Case TypeOf entity Is Polyline
                Dim poly As Polyline = CType(entity, Polyline)
                Return poly.Closed

        ' Caso para FeatureLines: revisa si existe una propiedad "Closed" (si no, ajusta según tu contexto)
            Case TypeOf entity Is FeatureLine
                Dim feature As FeatureLine = CType(entity, FeatureLine)
                Return feature.Closed ' Asegúrate de que FeatureLine tenga una propiedad Closed o ajusta este caso

        ' Caso para líneas: Las líneas generalmente no tienen una propiedad "Closed", pueden considerarse abiertas
        ' Aquí, devuelve True si una línea estuviera cerrada, ajusta según la implementación
            Case TypeOf entity Is Line
                Dim Line As Line = CType(entity, Line)
                Return Line.Closed ' Ajusta si Line tiene o no una propiedad Closed en tu implementación

        ' Caso para polilíneas 2D (Polyline2d): verifica la propiedad "Closed"
            Case TypeOf entity Is Polyline2d
                Dim PL2D As Polyline2d = CType(entity, Polyline2d)
                Return PL2D.Closed

        ' Caso para polilíneas 3D (Polyline3d): verifica la propiedad "Closed"
            Case TypeOf entity Is Polyline3d
                Dim PL2D As Polyline3d = CType(entity, Polyline3d)
                Return PL2D.Closed

                ' Si el tipo de entidad no se identifica o no tiene propiedad "Closed", se considera como cerrada
            Case Else
                Return True
        End Select
    End Function

    ' Función que obtiene la longitud de una entidad compatible (Polyline, Polyline2d, Polyline3d, FeatureLine o Line)
    ' Devuelve la longitud como Double, o 0 si la entidad no tiene una propiedad de longitud
    Public Shared Function GetEntityLength(Entity As Entity) As Double
        Dim entityLength As Double = 0
        ' Selección de casos en función del tipo de entidad, cada uno devuelve la longitud específica
        Select Case True
        ' Caso para Polyline: retorna la propiedad Length directamente
            Case TypeOf Entity Is Polyline
                entityLength = CType(Entity, Polyline).Length

        ' Caso para Polyline2d: verifica y retorna la propiedad Length si existe en tu contexto
            Case TypeOf Entity Is Polyline2d
                entityLength = CType(Entity, Polyline2d).Length ' Ajusta según tu implementación

        ' Caso para Polyline3d: verifica y retorna la propiedad Length si existe en tu contexto
            Case TypeOf Entity Is Polyline3d
                entityLength = CType(Entity, Polyline3d).Length ' Ajusta según tu implementación

        ' Caso para FeatureLine: usa un método (p.ej., Length2D) que devuelva la longitud en 2D
        ' Cambia este método según la disponibilidad en tu contexto
            Case TypeOf Entity Is FeatureLine
                entityLength = CType(Entity, FeatureLine).Length2D() ' Reemplaza con el método correcto

        ' Caso para Line: verifica y retorna la propiedad Length
            Case TypeOf Entity Is Line
                entityLength = CType(Entity, Line).Length
        End Select
        ' Retorna la longitud calculada para la entidad, o 0 si no se obtuvo longitud
        Return entityLength
    End Function

    Public Shared Function ConcertTempPL(Entity As Entity) As Entity
        Dim polyline As Polyline

        Select Case True
        ' Convert FeatureLine to Polyline
            Case TypeOf Entity Is FeatureLine
                Dim FL As FeatureLine = CType(Entity, FeatureLine)
                polyline = ShFeaturelineToPolyline.ConvertFeaturelineToPolyline(FL)

        ' If entity is a Polyline, cast it directly
            Case TypeOf Entity Is Polyline
                polyline = CType(Entity, Polyline)

        ' Convert Polyline2d to Polyline
            Case TypeOf Entity Is Polyline2d
                Dim polyline2d As Polyline2d = CType(Entity, Polyline2d)
                polyline = CPoLy2dToPL(polyline2d, False)

        ' Convert Polyline3d to Polyline
            Case TypeOf Entity Is Polyline3d
                Dim polyline3d As Polyline3d = CType(Entity, Polyline3d)
                polyline = CPoLy3dToPL(polyline3d, False)

        ' Convert Line to Polyline
            Case TypeOf Entity Is Line
                Dim line As Line = CType(Entity, Line)
                polyline = ConvertLineToPolyline(line, False)

            Case Else
                ' If no supported type, return the original entity
                Return Entity
        End Select

        Return CType(polyline, Entity)

    End Function

    ' Método auxiliar para obtener información relacionada con una polilínea
    Shared Sub PLRelatedInf(PL As Polyline, Alignment As Alignment, ByRef Len As Double, ByRef startStation As Double, ByRef endStation As Double, ByRef Side As String)
        Dim Area As Double
        Dim Side1 As String = String.Empty
        Dim Side2 As String = String.Empty
        Dim StartPT As Point3d
        Dim EndPT As Point3d

        CStationOffsetLabel.ProcessPolyline(Alignment, PL, StartPT, EndPT, Len, Area, startStation, endStation, Side1, Side2)
        CStationOffsetLabel.GetMxMnOPPL(startStation, endStation)

        If Side1 = Side2 Then Side = Side1
    End Sub

    'prueba del metodo 
    Public Shared Sub ByLenghtAndStation()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Dim listPolylines As List(Of ObjectId) = ConstructListType(Of Polyline)()
        Dim objectIds As List(Of ObjectId) = CheckEntityByLenght(listPolylines, 57.46, 2, 2)
        Dim objectIdsByStation As List(Of ObjectId) = CheckEntityByStation(objectIds, 57.46, 2810.3, 2885.17, 2, 2, "Right")


        For Each ObjID As ObjectId In objectIdsByStation
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                ' Abrir la tabla de bloques en modo lectura
                Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                ' Abrir el registro del espacio del modelo en modo lectura
                Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                Dim entity As Entity = CType(trans.GetObject(ObjID, OpenMode.ForRead), Entity)
                'ahora si hay mas de un ojecto buscar el que tine la menor tolerancia 
                'para eso pasamo un listado de tolerancias previas 

                AcadZoomManager.SelectedZoom(entity.Handle.ToString(), Application.DocumentManager.MdiActiveDocument)
            End Using
        Next

    End Sub

End Class
Public Class ToleranceRecord
    Public Property ObjectId As ObjectId
    Public Property LengthTolerance As Double
    Public Property StartStationTolerance As Double
    Public Property EndStationTolerance As Double

    Public Sub New(id As ObjectId, cLenght As Double, Lenght As Double, cstartStation As Double, startStation As Double, cendStation As Double, endStation As Double)
        ObjectId = id
        ' Calcular tolerancias
        LengthTolerance = Math.Abs(cLenght - Lenght)
        StartStationTolerance = Math.Abs(cstartStation - startStation)
        EndStationTolerance = Math.Abs(cendStation - endStation)
    End Sub

    ' Método para obtener la tolerancia total
    Public Function TotalTolerance() As Double
        Return LengthTolerance + StartStationTolerance + EndStationTolerance
    End Function
End Class