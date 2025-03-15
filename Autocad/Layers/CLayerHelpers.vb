Imports System.Diagnostics
Imports System.Linq
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.LayerManager
Imports Autodesk.AutoCAD.Runtime
Imports ObjectId = Autodesk.AutoCAD.DatabaseServices.ObjectId
Imports ObjectIdCollection = Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection
Public Class CLayerHelpers
    Public Shared Sub GetextraLayer(ByRef nombresCapas As String(), LayerFNAme As String)
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        ' Verificar si hay transacciones abiertas
        If AcCurD.TransactionManager.NumberOfActiveTransactions > 0 Then
            AcEd.WriteMessage(vbCrLf & "Advertencia: Hay transacciones abiertas antes de iniciar una nueva.")
            Debug.Assert(AcCurD.TransactionManager.NumberOfActiveTransactions = 0, "Hay transacciones abiertas antes de iniciar una nueva.")
        End If

        ' Si el grupo de capas ya existe, obtener su referencia
        Dim LgExistente As LayerGroup = GetLayerGroupByName(LayerFNAme)

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            AcEd.WriteMessage(vbCrLf & "Transacción iniciada. Transacciones activas: " & AcCurD.TransactionManager.NumberOfActiveTransactions)
            Try
                Dim LYTable As LayerTable = trans.GetObject(AcCurD.LayerTableId, OpenMode.ForRead)

                ' Verificar si el grupo de capas existe
                If LgExistente IsNot Nothing Then
                    ' Bloquear el documento una vez antes del bucle
                    AcDoc.LockDocument()

                    ' Iterar sobre las capas existentes en el grupo
                    For Each layerId As ObjectId In LgExistente.LayerIds
                        Try
                            ' Obtener la capa
                            Dim LYObj As LayerTableRecord = TryCast(trans.GetObject(layerId, OpenMode.ForRead), LayerTableRecord)

                            ' Verificar si LYObj es Nothing
                            If LYObj Is Nothing Then
                                AcEd.WriteMessage(vbCrLf & "Warning: LYObj es Nothing para layerId: " & layerId.ToString())
                                Debug.Assert(LYObj IsNot Nothing, "LYObj es Nothing para layerId: " & layerId.ToString())
                                Continue For
                            End If

                            ' Verificar si la capa no está en la lista nombresCapas
                            If Not nombresCapas.Contains(LYObj.Name) Then
                                ' Agregar la capa a nombresCapas
                                ReDim Preserve nombresCapas(nombresCapas.Length)
                                nombresCapas(nombresCapas.Length - 1) = LYObj.Name
                            End If
                        Catch ex As Exception
                            AcEd.WriteMessage(vbCrLf & "Error procesando layerId " & layerId.ToString() & ": " & ex.Message)
                        End Try
                    Next
                End If
                trans.Commit()
            Catch ex As Exception
                ' Manejar cualquier error
                AcEd.WriteMessage(vbCrLf & "Error en transacción: " & ex.Message)
                trans.Abort()
            Finally
                If trans.IsDisposed = False Then
                    trans.Dispose()
                End If
            End Try
        End Using

        ' Verificar si la transacción se cerró correctamente
        AcEd.WriteMessage(vbCrLf & "Transacción finalizada. Transacciones activas: " & AcCurD.TransactionManager.NumberOfActiveTransactions)
        Debug.Assert(AcCurD.TransactionManager.NumberOfActiveTransactions = 0, "La transacción no se cerró correctamente.")
    End Sub

    Public Shared Sub CreateLayerGroup(nombresCapas As String(), LayerGroupName As String)
        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Dim LayFCol As LayerFilterCollection = AcCurD.LayerFilters.Root.NestedFilters

        If LayerGroupExists(LayerGroupName) Then
            GetextraLayer(nombresCapas, LayerGroupName)
            DeleteLayerFilterbyName(LayerGroupName)
        Else
            GetextraLayer(nombresCapas, LayerGroupName)
        End If
        'Crear un LAyer Filter 
        'Dim lf As LayerFilter = New LayerFilter()
        Dim Lg As New LayerGroup With {
            .Name = LayerGroupName
        }

        Dim ListObj As New ObjectIdCollection()


        For Each x As String In nombresCapas
            CreateLayerIfNotExists(x)
        Next

        Dim trans As Transaction = AcCurD.TransactionManager.StartTransaction()

        Using trans
            Try
                Dim LYTable As LayerTable = trans.GetObject(AcCurD.LayerTableId, OpenMode.ForWrite)
                ' Obtener las capas que ya están en el grupo
                Dim existingLayerIds As New List(Of ObjectId)()
                For Each layerId As ObjectId In Lg.LayerIds
                    existingLayerIds.Add(layerId)
                Next

                For Each LYId As ObjectId In LYTable
                    Dim LYObj As LayerTableRecord = trans.GetObject(LYId, OpenMode.ForWrite)

                    If nombresCapas.Contains(LYObj.Name) And Not existingLayerIds.Contains(LYId) Then
                        ListObj.Add(LYId)
                    End If
                Next

                'create a 
                Dim Lft As LayerFilterTree = AcCurD.LayerFilters
                Dim Lfc As LayerFilterCollection = Lft.Root.NestedFilters
                If Lg.LayerIds.Count = 0 Then
                    For Each id As ObjectId In ListObj
                        Lg.LayerIds.Add(id)
                    Next
                    Lfc.Add(Lg)
                    AcCurD.LayerFilters = Lft
                End If
                AcEd.WriteMessage("\n\{0}\group created containing {1} layers.\n", Lg.Name, ListObj.Count)
                trans.Commit()
            Catch ex As Exception
                ' Manejar cualquier error
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
            Finally
                If trans.IsDisposed = False Then
                    trans.Dispose()
                End If
            End Try
        End Using
    End Sub
    Public Shared Sub CreateLayerIfNotExistsRAcEnt(ByVal layerName As String)
        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Dim layerId As ObjectId
        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim lt As LayerTable = trans.GetObject(AcCurD.LayerTableId, OpenMode.ForWrite)

                ' Verificar si el layer ya existe
                If lt.Has(layerName) Then
                    layerId = lt(layerName)
                Else
                    ' Si el layer no existe, crearlo
                    Dim newLayer As New LayerTableRecord With {
                        .Name = layerName
                    }

                    ' Añadir el nuevo layer a la tabla de layers
                    lt.UpgradeOpen()
                    layerId = lt.Add(newLayer)
                    trans.AddNewlyCreatedDBObject(newLayer, True)

                    ' Commit the transaction
                    trans.Commit()

                    AcEd.WriteMessage(vbCrLf & "Layer '" & layerName & "' created.")
                End If
                AcDoc.LockDocument()

            Catch ex As Exception
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
            Finally
                If trans.IsDisposed = False Then
                    trans.Dispose()
                End If
            End Try
        End Using
    End Sub

    Public Shared Function CreateLayerIfNotExists(ByVal layerName As String) As ObjectId
        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Dim layerId As ObjectId

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim lt As LayerTable = trans.GetObject(AcCurD.LayerTableId, OpenMode.ForWrite)

                ' Verificar si el layer ya existe
                If lt.Has(layerName) Then
                    layerId = lt(layerName)
                Else
                    ' Si el layer no existe, crearlo
                    Dim newLayer As New LayerTableRecord With {
                        .Name = layerName
                    }

                    ' Añadir el nuevo layer a la tabla de layers
                    lt.UpgradeOpen()
                    layerId = lt.Add(newLayer)
                    trans.AddNewlyCreatedDBObject(newLayer, True)

                    ' Commit the transaction
                    trans.Commit()

                    AcEd.WriteMessage(vbCrLf & "Layer '" & layerName & "' created.")
                End If
            Catch ex As Exception
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
            Finally
                If trans.IsDisposed = False Then
                    trans.Dispose()
                End If
            End Try
        End Using

        Return layerId
    End Function

    Public Sub RemoveLayerGroup(ByVal groupName As String)
        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim filterTree As LayerFilterTree = AcCurD.LayerFilters
                Dim rootFilters As LayerFilterCollection = filterTree.Root.NestedFilters

                ' Buscar el grupo de capas por su nombre
                For Each filter As LayerFilter In rootFilters
                    If filter.Name = groupName AndAlso TypeOf filter Is LayerGroup Then
                        Dim groupToRemove As LayerGroup = DirectCast(filter, LayerGroup)
                        ' Eliminar las capas del grupo
                        For Each layerId As ObjectId In groupToRemove.LayerIds
                            groupToRemove.LayerIds.Remove(layerId)
                        Next

                        ' Eliminar el grupo de capas de la colección de filtros
                        rootFilters.Remove(groupToRemove)

                        AcEd.WriteMessage(vbCrLf & "Layer group '" & groupName & "' removed.")
                        trans.Commit()
                        trans.Dispose()
                        Exit For
                    End If
                Next

                ' Confirmar los cambios en la base de datos
                trans.Commit()
                trans.Dispose()
            Catch ex As Exception
                ' Manejar cualquier error
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
            Finally
                If trans.IsDisposed = False Then
                    trans.Dispose()
                End If
            End Try
        End Using
    End Sub
    ' Change layer of selected AutoCAD entity
    Public Shared Sub ChangeLayersAcEnt(hdString As String, LayerName As String) ', Alignment As Alignment)
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        'If Alignment Is Nothing Then Exit Sub

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim id As ObjectId = CLHandle.GetObjIdBStr(hdString)
                AcDoc.LockDocument()
                Dim ent As Entity = TryCast(trans.GetObject(id, OpenMode.ForWrite), Entity)
                If id.IsNull Then
                    AcDoc.Editor.WriteMessage(vbCrLf & "Error: El ObjectId es nulo.")
                    Exit Sub
                End If
                ent.Layer = LayerName
                'ent.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256)
                trans.Commit()
            Catch ex As Exception
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
            Finally
                If trans.IsDisposed = False Then
                    trans.Dispose()
                End If
            End Try
        End Using
    End Sub
    ' Change layer of selected AutoCAD entity
    Public Shared Function GetLayersAcEnt(hdString As String) As String ', Alignment As Alignment)
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Using AcDoc.LockDocument()
                Try
                    Dim id As ObjectId = CLHandle.GetObjIdBStr(hdString)
                    If id.IsNull Then
                        AcDoc.Editor.WriteMessage(vbCrLf & "Error: El ObjectId es nulo.")
                        Return String.Empty
                    End If
                    Dim ent As Entity = TryCast(trans.GetObject(id, OpenMode.ForRead), Entity)
                    trans.Commit()
                    Return ent.Layer
                Catch ex As Exception
                    AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                    trans.Abort()
                    Return String.Empty
                Finally
                    If Not trans.IsDisposed Then trans.Dispose()
                End Try
            End Using
        End Using
    End Function
    ' Función para establecer una capa como la capa activa
    Public Shared Sub SetCurrentLayer(layerName As String)
        ' Obtener el documento actual de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener el contenedor de capas (LayerTable)
                Dim layerTable As LayerTable = CType(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

                ' Verificar si la capa existe
                If layerTable.Has(layerName) Then
                    ' Obtener la capa y configurarla como activa
                    db.Clayer = layerTable(layerName)
                Else
                    ed.WriteMessage(vbLf & "La capa '" & layerName & "' no existe.")
                End If

                ' Confirmar los cambios
                trans.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
            End Try
        End Using
    End Sub
    ' Función para obtener el nombre de la capa activa
    Public Shared Sub GetCurrentLayer(ByRef layerName As String)
        ' Obtener el documento actual de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener el ID de la capa activa
                Dim activeLayerId As ObjectId = db.Clayer

                ' Obtener el registro de la capa activa (LayerTableRecord)
                Dim activeLayer As LayerTableRecord = CType(trans.GetObject(activeLayerId, OpenMode.ForRead), LayerTableRecord)

                ' Asignar el nombre de la capa activa al parámetro 'layerName'
                layerName = activeLayer.Name

                ' Confirmar los cambios
                trans.Commit()
            Catch ex As Exception
                ' Manejo de excepciones
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
            End Try
        End Using
    End Sub

End Class
