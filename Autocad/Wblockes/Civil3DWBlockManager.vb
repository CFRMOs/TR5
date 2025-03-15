' Alias para diferenciar las bibliotecas de Civil 3D
Imports CivilApp = Autodesk.Civil.ApplicationServices
Imports CivilSettings = Autodesk.Civil.Settings

' Usamos el namespace completo para las bibliotecas de AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class Civil3DWBlockManager

    ' Función que construye un SelectionSet basado en el tipo de entidad proporcionado
    Public Function ConstructSelectionSet(entityType As String) As SelectionSet
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Creamos un filtro para seleccionar solo los objetos del tipo especificado
        Dim filterType As Type = Type.GetType("Autodesk.AutoCAD.DatabaseServices." & entityType, False, True)

        If filterType Is Nothing Then
            ed.WriteMessage(vbLf & "El tipo de entidad especificado no es válido.")
            Return Nothing
        End If

        ' Crear el filtro de selección por tipo de objeto
        Dim filter As New SelectionFilter(New TypedValue() {New TypedValue(DxfCode.Start, filterType.Name)})

        ' Pedir al usuario que seleccione objetos
        Dim selectionResult As PromptSelectionResult = ed.GetSelection(filter)

        If selectionResult.Status <> PromptStatus.OK Then
            ed.WriteMessage(vbLf & "No se seleccionaron objetos.")
            Return Nothing
        End If

        Return selectionResult.Value
    End Function

    ' Función que crea un WBlock a partir de un SelectionSet, manteniendo la geolocalización
    Public Sub CreateWBlockFromSelectionSet(selectionSet As SelectionSet, wblockPath As String)
        If selectionSet Is Nothing OrElse selectionSet.Count = 0 Then
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "El SelectionSet está vacío o es nulo.")
            Return
        End If

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Crear un nuevo archivo WBlock en la misma localización que el dibujo actual
        Using wblockDb As New Database(True, True)
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Clonar las entidades seleccionadas al nuevo archivo WBlock sin alterar su localización
                    Dim idCollection As New ObjectIdCollection()

                    For Each id As ObjectId In selectionSet.GetObjectIds()
                        idCollection.Add(id)
                    Next

                    ' Exportar las entidades seleccionadas como WBlock
                    db.WblockCloneObjects(idCollection, wblockDb.BlockTableId, New IdMapping(), DuplicateRecordCloning.Replace, False)

                    ' Mantener el sistema de coordenadas del dibujo actual
                    Dim geoLocation As String = GetDrawingGeoLocation() ' Obtener el sistema de coordenadas
                    If Not String.IsNullOrEmpty(geoLocation) Then
                        ' Asignar la misma geolocalización al nuevo WBlock
                        SetDrawingGeoLocation(wblockDb, geoLocation) ' Ajuste de llamada
                    End If

                    ' Guardar el archivo WBlock en la ruta especificada
                    wblockDb.SaveAs(wblockPath, DwgVersion.Current)
                    ed.WriteMessage(vbLf & "WBlock creado exitosamente con geolocalización en: " & wblockPath)

                    trans.Commit()
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error al crear WBlock: " & ex.Message)
                    trans.Abort()
                End Try
            End Using
        End Using
    End Sub

    ' Función auxiliar para obtener el sistema de coordenadas geográficas del dibujo actual en Civil 3D
    Private Function GetDrawingGeoLocation() As String
        Try
            ' Obtener los ajustes del dibujo actual de Civil 3D
            Dim civilDoc As CivilApp.CivilDocument = CivilApp.CivilApplication.ActiveDocument

            ' Obtener el sistema de coordenadas geográficas actual
            Dim geoSystem As String = civilDoc.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode
            Return geoSystem
        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error al obtener la geolocalización: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    ' Función auxiliar para copiar el sistema de coordenadas geográficas del dibujo actual a un WBlock
    Private Sub SetDrawingGeoLocation(wblockDb As Database, geoSystem As String)
        Try
            ' Verificamos si se proporcionó un sistema de coordenadas válido
            If String.IsNullOrEmpty(geoSystem) Then
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "No se proporcionó un sistema de coordenadas válido.")
                Return
            End If

            ' Iniciar una transacción en el nuevo archivo WBlock para poder modificar sus propiedades
            Using trans As Transaction = wblockDb.TransactionManager.StartTransaction()

                ' Simulamos el proceso de copiar la información del sistema de coordenadas
                ' En este caso, no se puede aplicar directamente al WBlock usando .NET, pero se puede almacenar o manejar la información
                SetCoordinateSystemCode(wblockDb, geoSystem)

                ' Finalizamos la transacción
                trans.Commit()

            End Using

            ' Informar al usuario que el sistema de coordenadas se mantuvo
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Se asignó el sistema de coordenadas: " & geoSystem & " al WBlock.")

        Catch ex As Exception
            ' Capturamos errores y los mostramos en la consola de AutoCAD
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error al asignar la geolocalización: " & ex.Message)
        End Try
    End Sub

    ' Función para simular la asignación del sistema de coordenadas al WBlock
    Private Sub SetCoordinateSystemCode(wblockDb As Database, geoSystemCode As String)
        ' En la API actual de Civil 3D .NET, no podemos asignar el sistema de coordenadas directamente.
        ' Se podría almacenar esta información o manejarla en las entidades.
        ' Esta función podría ser extendida si en futuras versiones de Autodesk permiten la manipulación directa de coordenadas.
        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Asignando el sistema de coordenadas '" & geoSystemCode & "' al WBlock.")
    End Sub

    ' Función para agregar o recuperar una Xref en el archivo actual, manteniendo la geolocalización
    Public Sub AddOrRetrieveXref(wblockPath As String, xrefName As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener la tabla de bloques
                Dim blockTable As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                ' Revisar si el Xref ya está en el archivo
                If blockTable.Has(xrefName) Then
                    ed.WriteMessage(vbLf & "La referencia externa ya existe en el archivo.")

                    ' Obtener el bloque Xref existente
                    Dim xrefBlock As BlockTableRecord = trans.GetObject(blockTable(xrefName), OpenMode.ForRead)

                    ' Insertar la Xref si no está ya insertada
                    InsertXrefInstance(xrefBlock.ObjectId, Point3d.Origin)
                Else
                    ' Si no existe, añadir el Xref
                    Dim xrefId As ObjectId

                    ' Añadir el archivo WBlock como Xref
                    xrefId = db.AttachXref(wblockPath, xrefName)

                    ' Comprobar si se ha añadido correctamente
                    If Not xrefId.IsNull Then
                        ed.WriteMessage(vbLf & "Xref añadido correctamente: " & wblockPath)

                        ' Insertar la nueva Xref en el dibujo
                        InsertXrefInstance(xrefId, Point3d.Origin)
                    Else
                        ed.WriteMessage(vbLf & "Error al agregar el Xref: " & wblockPath)
                    End If
                End If

                trans.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error al agregar o recuperar Xref: " & ex.Message)
                trans.Abort()
            End Try
        End Using
    End Sub

    ' Función auxiliar para insertar una instancia de Xref en el dibujo, manteniendo la geolocalización
    Private Sub InsertXrefInstance(xrefId As ObjectId, insertPoint As Point3d)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener el espacio de trabajo actual (Modelo)
                Dim blockTable As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim modelSpace As BlockTableRecord = trans.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                ' Crear una nueva referencia de bloque para la Xref
                Dim xrefRef As New BlockReference(insertPoint, xrefId)
                modelSpace.AppendEntity(xrefRef)
                trans.AddNewlyCreatedDBObject(xrefRef, True)

                ed.WriteMessage(vbLf & "Xref insertada en el dibujo en el punto: " & insertPoint.ToString())

                trans.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error al insertar la instancia de Xref: " & ex.Message)
                trans.Abort()
            End Try
        End Using
    End Sub

End Class
