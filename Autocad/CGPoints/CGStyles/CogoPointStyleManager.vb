Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Colors

Public Class CogoPointStyleManager

    Private ReadOnly civildoc As CivilDocument

    ' Constructor que recibe el documento de Civil 3D
    Public Sub New(ByRef civilDocument As CivilDocument)
        Me.civildoc = civilDocument
    End Sub

    ' Método para verificar si un estilo de punto ya existe en el documento
    Public Function GetPointStyleByName(styleName As String) As ObjectId
        Dim pointStyles As PointStyleCollection = civildoc.Styles.PointStyles

        ' Verificar si la colección de estilos está vacía o es null
        If pointStyles Is Nothing OrElse pointStyles.Count = 0 Then
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("No point styles available in the document.")
            Return ObjectId.Null
        End If

        ' Iniciar una transacción
        Using tr As Transaction = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction()
            Try
                For Each pointStyleId As ObjectId In pointStyles
                    ' Verificar si el ObjectId es válido
                    If pointStyleId = ObjectId.Null Then Continue For

                    ' Intentar obtener el PointStyle
                    Dim pointStyle As PointStyle = TryCast(tr.GetObject(pointStyleId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), PointStyle)

                    ' Verificar si se pudo obtener el estilo de punto
                    If pointStyle IsNot Nothing AndAlso pointStyle.Name = styleName Then
                        ' Si el estilo coincide, devolver el ObjectId
                        Return pointStyleId
                    End If
                Next
            Catch ex As Exception
                ' Manejar cualquier excepción que ocurra durante la transacción
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error retrieving PointStyle: " & ex.Message)
            End Try

            ' Si no se encuentra el estilo, devolver ObjectId.Null
            Return ObjectId.Null
        End Using
    End Function

    ' Método para crear un nuevo estilo de punto personalizado
    Public Function CreateCustomPointStyle(styleName As String, Optional markerType As CustomMarkerType = CustomMarkerType.CustomMarkerPlus, Optional markerSuperimposeType As CustomMarkerSuperimposeType = CustomMarkerSuperimposeType.Square, Optional markerColor As Color = Nothing, Optional labelColor As Color = Nothing) As ObjectId
        ' Obtener la base de datos del documento activo
        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database

        ' Crear el nuevo estilo de punto
        Dim pointStyleId As ObjectId = civildoc.Styles.PointStyles.Add(styleName)

        ' Si no se pasa un color de marcador, asignar uno por defecto (rojo)
        If markerColor Is Nothing Then
            markerColor = Color.FromColorIndex(ColorMethod.ByAci, 1) ' Rojo
        End If

        ' Si no se pasa un color de etiqueta, asignar uno por defecto (azul)
        If labelColor Is Nothing Then
            labelColor = Color.FromColorIndex(ColorMethod.ByAci, 5) ' Azul
        End If

        ' Obtener el estilo creado para modificar sus propiedades
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim pointStyle As PointStyle = CType(pointStyleId.GetObject(OpenMode.ForWrite), PointStyle)

                ' Definir el tipo de marcador
                pointStyle.MarkerType = PointMarkerDisplayType.UseCustomMarker

                ' Establecer el tipo de marcador personalizado
                pointStyle.CustomMarkerStyle = markerType
                pointStyle.CustomMarkerSuperimposeStyle = markerSuperimposeType

                ' Configurar el color del marcador
                pointStyle.GetDisplayStylePlan(PointDisplayStyleType.Marker).Color = markerColor

                ' Configurar el color de la etiqueta
                pointStyle.GetLabelDisplayStylePlan().Color = labelColor

                ' Guardar los cambios
                tr.Commit()
            Catch ex As Exception
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error al crear el estilo de punto: " & ex.Message)
            End Try
        End Using

        Return pointStyleId
    End Function

    ' Método para asignar un estilo a un CogoPoint
    Public Sub AssignStyleToCogoPoint(cogoPoint As CogoPoint, styleName As String)
        ' Verificar si el estilo ya existe
        Dim pointStyleId As ObjectId = GetPointStyleByName(styleName)

        ' Si el estilo no existe, crearlo
        If pointStyleId = ObjectId.Null Then
            ' Crear un estilo por defecto si no existe (esto puede personalizarse según sea necesario)
            pointStyleId = CreateCustomPointStyle(styleName)
        End If

        ' Asignar el estilo al CogoPoint
        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener el CogoPoint en modo escritura
                Dim cogoPointEnt As CogoPoint = CType(tr.GetObject(cogoPoint.ObjectId, OpenMode.ForWrite), CogoPoint)

                ' Asignar el estilo al CogoPoint
                cogoPointEnt.StyleId = pointStyleId

                ' Guardar los cambios
                tr.Commit()
            Catch ex As Exception
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error al asignar el estilo al CogoPoint: " & ex.Message)
            End Try
        End Using
    End Sub

End Class

Public Class CogoPointGroupStyleManager

    ' Asignar un estilo de punto a un grupo de CogoPoints
    Public Sub AssignPointStyleToGroup(groupName As String, pointStyleId As ObjectId, labelStyleId As ObjectId)
        ' Obtener el documento Civil 3D activo
        Dim civilDoc As CivilDocument = CivilApplication.ActiveDocument
        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener el grupo de puntos por nombre
                Dim pointGroupCollection As PointGroupCollection = civilDoc.PointGroups
                Dim groupId As ObjectId = GetPointGroupByName(groupName, pointGroupCollection)

                If groupId <> ObjectId.Null Then
                    ' Abrir el grupo de puntos para escritura
                    Dim pointGroup As PointGroup = CType(tr.GetObject(groupId, OpenMode.ForWrite), PointGroup)

                    ' Asignar el estilo de punto
                    If pointStyleId <> ObjectId.Null Then
                        pointGroup.PointStyleId = pointStyleId
                    End If

                    ' Asignar el estilo de etiquetas
                    If labelStyleId <> ObjectId.Null Then
                        'pointGroup.PointLabelStyleId = labelStyleId
                    End If

                    ' Guardar los cambios
                    tr.Commit()
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Estilo asignado correctamente al grupo de puntos: {groupName}" & vbLf)
                Else
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Grupo de puntos no encontrado: {groupName}" & vbLf)
                End If
            Catch ex As Exception
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Error: {ex.Message}" & vbLf)
            End Try
        End Using
    End Sub

    ' Eliminar el estilo de punto de un grupo de CogoPoints (restablecer a <none>)
    Public Sub RemovePointStyleFromGroup(groupName As String)
        ' Obtener el documento Civil 3D activo
        Dim civilDoc As CivilDocument = CivilApplication.ActiveDocument
        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener el grupo de puntos por nombre
                Dim pointGroupCollection As PointGroupCollection = civilDoc.PointGroups
                Dim groupId As ObjectId = GetPointGroupByName(groupName, pointGroupCollection)

                If groupId <> ObjectId.Null Then
                    ' Abrir el grupo de puntos para escritura
                    Dim pointGroup As PointGroup = CType(tr.GetObject(groupId, OpenMode.ForWrite), PointGroup)

                    ' Restablecer el estilo de punto a <none>
                    pointGroup.PointStyleId = ObjectId.Null

                    ' Restablecer el estilo de etiquetas a <none>
                    pointGroup.PointLabelStyleId = ObjectId.Null

                    ' Guardar los cambios
                    tr.Commit()
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Estilos eliminados del grupo de puntos: {groupName}" & vbLf)
                Else
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Grupo de puntos no encontrado: {groupName}" & vbLf)
                End If
            Catch ex As Exception
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"Error: {ex.Message}" & vbLf)
            End Try
        End Using
    End Sub

    ' Método auxiliar para obtener el ObjectId de un grupo de puntos por su nombre
    Private Function GetPointGroupByName(groupName As String, pointGroupCollection As PointGroupCollection) As ObjectId
        For Each groupId As ObjectId In pointGroupCollection
            Dim group As PointGroup = CType(groupId.GetObject(OpenMode.ForRead), PointGroup)
            If group.Name = groupName Then
                Return groupId
            End If
        Next
        Return ObjectId.Null
    End Function

End Class