Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Public Module CMDLayerTools
    <CommandMethod("CL")>
    Public Sub ChangeLayerOfEntitiess()
        Dim doc As Document = GetDocumentManager().MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Ask the user for the layer name, allowing
        ' spaces to be entered
        Dim pso As New PromptStringOptions(vbCrLf & "Enter name of layer to search for: ") With {
            .AllowSpaces = True
        }
        Dim pr As PromptResult = ed.GetString(pso)

        If pr.Status <> PromptStatus.OK Then
            Return
        End If

        Dim layerName As String = pr.StringResult

        ' We won't validate whether the layer exists -
        ' we'll just see what's returned by the selection.
        Dim tvs(0) As TypedValue
        tvs(0) = New TypedValue(CInt(DxfCode.LayerName), layerName)
        Dim sf As New SelectionFilter(tvs)
        Dim psr As PromptSelectionResult = ed.SelectAll

        Dim count As Integer = 0
        If psr.Status = PromptStatus.OK Then
            count = psr.Value.Count
        End If

        If psr.Status = PromptStatus.OK OrElse psr.Status = PromptStatus.Error Then
            ' Display the count of entities on that layer
            ed.WriteMessage(vbCrLf & "Found {0} entit{1} on layer ""{2}"".", count, If(count = 1, "y", "ies"), layerName)

            ' If there are some on this layer,
            ' prompt for the layer to move them to
            If count > 0 Then
                pso.Message = vbCrLf & "Enter new layer for these entities or return to leave them alone: "
                pr = ed.GetString(pso)

                If pr.Status <> PromptStatus.OK OrElse pr.StringResult = "" Then
                    Return
                End If

                Dim newLayerName As String = pr.StringResult

                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    ' This time we do check whether
                    ' the layer exists
                    Dim lt As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

                    If Not lt.Has(newLayerName) Then
                        ed.WriteMessage(vbCrLf & "Layer not found.")
                    Else
                        Dim changedCount As Integer = 0

                        ' We have the layer table open, so let's
                        ' get the layer ID and use that
                        Dim lid As ObjectId = lt(newLayerName)
                        For Each id As ObjectId In psr.Value.GetObjectIds()
                            Dim ent As Entity = CType(tr.GetObject(id, OpenMode.ForWrite), Entity)
                            ent.LayerId = lid
                            ' Could also have used:
                            ' ent.Layer = newLayerName;
                            ' but this way is more efficient and cleaner
                            changedCount += 1
                        Next

                        ed.WriteMessage(vbCrLf & "Changed {0} entit{1} from layer ""{2}"" to layer ""{3}"".", changedCount, If(changedCount = 1, "y", "ies"), layerName, newLayerName)
                    End If

                    tr.Commit()
                End Using
            End If
        End If
    End Sub

    Public Sub ChangeEntityLayerByHandle(Hdl As Handle, NewLayerName As String, BaseLayerName As String)
        Dim doc As Document = GetDocumentManager().MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        '' Convertir el handle de texto a un objeto TabName
        'Dim Hdl As New TabName(Convert.ToInt64(handle, 16))

        'Dim GestorHL As New HandleCommands

        Dim tvs(0) As TypedValue

        tvs(0) = New TypedValue(CInt(DxfCode.LayerName), BaseLayerName)

        Dim sf As New SelectionFilter(tvs)
        Dim psr As PromptSelectionResult = ed.SelectAll

        Dim count As Integer = 0
        If psr.Status = PromptStatus.OK Then
            count = psr.Value.Count
        End If

        If psr.Status = PromptStatus.OK OrElse psr.Status = PromptStatus.Error Then
            ' Display the count of entities on that layer

            ' If there are some on this layer,
            If count > 0 Then
                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    ' This time we do check whether
                    ' the layer exists
                    Dim lt As LayerTable = CType(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)

                    If Not lt.Has(NewLayerName) Then

                        ed.WriteMessage(vbCrLf & "Layer not found.")
                        tr.Abort()
                    Else
                        Dim changedCount As Integer = 0

                        ' We have the layer table open, so let's
                        ' get the layer ID and use that
                        Dim lid As ObjectId = lt(NewLayerName)
                        'Dim objhandleM As New HandleCommands
                        Dim exacobjID As ObjectId = CLHandle.GetEntityIdByHandle(Hdl)
                        For Each id As ObjectId In psr.Value.GetObjectIds()
                            If id = exacobjID Then
                                Try
                                    Dim ent As Entity = CType(tr.GetObject(id, OpenMode.ForWrite), Entity)
                                    ent.LayerId = lid
                                    changedCount += 1
                                    tr.Commit()
                                Catch ex As Exception
                                    ed.WriteMessage(vbCrLf & "Error: " & ex.Message)
                                    tr.Abort()
                                Finally
                                    tr.Dispose()
                                End Try
                                Exit For
                            End If
                        Next
                        'ed.WriteMessage(vbCrLf & "Changed {0} entit{1} from layer ""{2}"" to layer ""{3}"".", changedCount, If(changedCount = 1, "y", "ies"), BaseLayerName, NewLayerName)
                    End If
                End Using
            End If
        End If
    End Sub

End Module

