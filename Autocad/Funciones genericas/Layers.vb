Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.LayerManager
Imports Autodesk.AutoCAD.Runtime

Module Layers

    ' Esta función elimina un filtro de capa específico por su nombre.
    ' Si el filtro de capa no se encuentra o no se puede eliminar, se muestra un mensaje correspondiente.
    ' Después de eliminar el filtro, se actualiza la lista de filtros.
    Public Sub DeleteLayerFilterbyName(ByVal filterName As String)
        Dim doc As Document = GetDocumentManager().MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Try
            ' Obtener los filtros de capa existentes
            Dim lft As LayerFilterTree = db.LayerFilters
            Dim lfc As LayerFilterCollection = lft.Root.NestedFilters

            ' Encontrar el índice del filtro de capa basado en su nombre
            Dim filterIndex As Integer = -1
            For i As Integer = 0 To lfc.Count - 1
                If lfc(i).Name = filterName Then
                    filterIndex = i
                    Exit For
                End If
            Next

            ' Verificar si se encontró el filtro de capa
            If filterIndex = -1 Then
                ed.WriteMessage(vbCrLf & "Layer filter '" & filterName & "' not found.")
                Return
            End If

            ' Obtener el filtro seleccionado
            Dim lf As LayerFilter = lfc(filterIndex)

            ' Si es posible eliminarlo, hacerlo
            If Not lf.AllowDelete Then
                ed.WriteMessage(vbCrLf & "Layer filter '" & filterName & "' cannot be deleted.")
            Else
                lfc.Remove(lf)
                db.LayerFilters = lft

                ListLayerFilters() ' Actualizar la lista de filtros después de eliminar uno
            End If
        Catch ex As Exception
            ' Manejar cualquier error
            ed.WriteMessage(vbCrLf & "Error: " & ex.Message)
        Finally

        End Try
    End Sub

    ' Esta función verifica si un grupo de capas con el nombre especificado existe en el dibujo actual.
    ' Utiliza una transacción para recorrer los filtros de capa y comprobar su existencia.
    Public Function LayerGroupExists(ByVal groupName As String) As Boolean
        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim filterTree As LayerFilterTree = AcCurD.LayerFilters
                Dim rootFilters As LayerFilterCollection = filterTree.Root.NestedFilters

                For Each filter As LayerFilter In rootFilters
                    If filter.Name = groupName Then
                        trans.Commit()
                        Return True
                    End If
                Next
                trans.Commit()
                ' No se encontró un filtro con el mismo nombre
                Return False
            Catch ex As Exception
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
                Return False
            Finally
                trans.Dispose()
            End Try
        End Using
    End Function

    ' Esta función obtiene un grupo de capas por su nombre.
    ' Utiliza una transacción para recorrer los filtros de capa y devolver el grupo si se encuentra.
    Public Function GetLayerGroupByName(ByVal groupName As String) As LayerGroup
        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim filterTree As LayerFilterTree = AcCurD.LayerFilters
                Dim rootFilters As LayerFilterCollection = filterTree.Root.NestedFilters

                For Each filter As LayerFilter In rootFilters
                    If filter.Name = groupName AndAlso TypeOf filter Is LayerGroup Then
                        trans.Commit()
                        Return DirectCast(filter, LayerGroup)
                    End If
                Next
                trans.Commit()
                ' No se encontró un grupo de capas con el nombre especificado
                Return Nothing
            Catch ex As Exception
                ' Manejar cualquier error
                trans.Abort()
                Return Nothing
            Finally
                trans.Dispose()
            End Try
        End Using
    End Function

    ' Esta función crea un grupo de capas y agrega las capas especificadas.
    ' Si el grupo ya existe, obtiene las capas existentes en el grupo y las devuelve.
    Public Function CrearGrupoYCapas(LayerFNAme As String, nombresCapas As String()) As List(Of String)
        ' Obtener el documento activo de AutoCAD
        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database

        Dim LayerNames As New List(Of String)
        ' Obtener el editor de AutoCAD
        Dim acEd As Editor = AcDoc.Editor

        ' Verificar si el grupo de capas "Drenajes Longitudinales" ya existe
        'Dim CLG As New CLayerHelpers


        CLayerHelpers.CreateLayerGroup(nombresCapas,
                             LayerFNAme)

        Using trans As Transaction = AcDoc.TransactionManager.StartTransaction()
            ' Obtener la base de datos del dibujo

            Dim grupoExiste As Boolean = False
            Dim groupId As ObjectId = ObjectId.Null
            'Dim LayerFNAme As String = "Drenajes Longitudinales"
            Dim Lg As LayerGroup = GetLayerGroupByName(LayerFNAme)
            Try
                Dim LYTable As LayerTable = CType(trans.GetObject(AcCurD.LayerTableId, OpenMode.ForRead), LayerTable)
                ' Obtener las capas que ya están en el grupo
                For Each layerId As ObjectId In Lg.LayerIds
                    Dim LYObj As LayerTableRecord = trans.GetObject(layerId, OpenMode.ForRead)
                    LayerNames.Add(LYObj.Name)
                Next
                trans.Commit()
                Return LayerNames
            Catch ex As Exception
                ' Manejar cualquier error
                trans.Abort()
                acEd.WriteMessage("No se obtuvo la colección de capas """ & LayerFNAme & """ ya existe")
                Return Nothing
            Finally
                trans.Dispose()
            End Try
        End Using
    End Function

    ' Este comando lista todos los filtros de capas anidados en el dibujo actual y muestra si pueden ser eliminados o no.
    <CommandMethod("LLFS")>
    Public Sub ListLayerFilters()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Listar los filtros de capa anidados
        Dim lfc As LayerFilterCollection = db.LayerFilters.Root.NestedFilters

        For i As Integer = 0 To lfc.Count - 1
            Dim lf As LayerFilter = lfc(i)
            ed.WriteMessage(
                vbCrLf & "{0} - {1} (can{2} be deleted)",
                i + 1,
                lf.Name,
                If(lf.AllowDelete, "", "not")
            )
        Next
    End Sub
End Module
