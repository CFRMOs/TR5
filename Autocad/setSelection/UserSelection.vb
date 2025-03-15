Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports C3DDBObject = Autodesk.Civil.DatabaseServices.DBObject
Imports C3DEntity = Autodesk.Civil.DatabaseServices.Entity
Module UserSelection
    Public Function GetPoint() As Point3d
        Dim doc As Document = GetDocumentManager().MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Dim pPtOpts As New PromptPointOptions(vbLf & "Enter a point:")
        Dim pPtRes As PromptPointResult = ed.GetPoint(pPtOpts)

        If pPtRes.Status = PromptStatus.OK Then
            Dim pt As Point3d = pPtRes.Value
            ed.WriteMessage(vbLf & "You picked: " & pt.ToString())
            Return pt
        Else
            ed.WriteMessage(vbLf & "Error or user cancelled")
        End If
    End Function
    Public Function SelectEntityByType(vtypeName As String, promptOptions As PromptEntityOptions, Optional ByRef PromptResult As PromptEntityResult = Nothing) As Entity

        Dim acDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Do
                    ' Solicitar al usuario que seleccione una entidad
                    PromptResult = acEd.GetEntity(promptOptions)

                    If PromptResult.Status = PromptStatus.OK Then
                        ' Obtener el objeto seleccionado
                        acDoc.Window.Focus()
                        Dim acEnt As Entity = acTrans.GetObject(PromptResult.ObjectId, OpenMode.ForRead)

                        'Verificar si el tipo de la entidad coincide
                        If TypeName(acEnt) = vtypeName Then
                            'Devolver la entidad si es del tipo deseado
                            acTrans.Commit()
                            Return acEnt
                        Else
                            'Notificar al usuario que la entidad seleccionada no es del tipo deseado
                            acEd.WriteMessage(vbCrLf & TypeName(acEnt))
                            acEd.WriteMessage(vbCrLf & "La entidad seleccionada no es del tipo " & vtypeName & ". Por favor, seleccione una entidad válida.")
                            acEnt = Nothing
                            Return Nothing
                        End If
                    ElseIf PromptResult.Status = PromptStatus.Cancel Then
                        ' Cancelar la transacción y devolver Nothing si se cancela la selección
                        acTrans.Dispose()
                        Return Nothing
                    End If
                Loop
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
            Finally
                'acTrans.Commit()
                If Not acTrans.IsDisposed Then
                    acTrans.Dispose() ' Dispose transaction if not already disposed
                End If
            End Try
            Return Nothing
        End Using
    End Function
    Public Function SelectEntity(promptOptions As PromptEntityOptions, Optional ByRef PromptResult As PromptEntityResult = Nothing) As Entity

        Dim acDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Do
                    ' Solicitar al usuario que seleccione una entidad
                    PromptResult = acEd.GetEntity(promptOptions)

                    If PromptResult.Status = PromptStatus.OK Then
                        ' Obtener el objeto seleccionado
                        Dim acEnt As Entity = acTrans.GetObject(PromptResult.ObjectId, OpenMode.ForRead)

                        'Verificar si el tipo de la entidad coincide
                        If Not acEnt = Nothing Then
                            'Devolver la entidad si es del tipo deseado
                            acTrans.Commit()
                            Return acEnt
                        Else
                            'Notificar al usuario que la entidad seleccionada no es del tipo deseado
                            acEd.WriteMessage(vbCrLf & TypeName(acEnt))
                            acEd.WriteMessage(vbCrLf & "No se selecciono una entidad. Por favor, seleccione una entidad válida.")
                            acEnt = Nothing
                            Return Nothing
                        End If
                    ElseIf PromptResult.Status = PromptStatus.Cancel Then
                        ' Cancelar la transacción y devolver Nothing si se cancela la selección
                        acTrans.Dispose()
                        Return Nothing
                    End If
                Loop
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
            Finally
                'acTrans.Commit()
                If Not acTrans.IsDisposed Then
                    acTrans.Dispose() ' Dispose transaction if not already disposed
                End If
            End Try

            Return Nothing
        End Using
    End Function
    Public Function SelectAcObjectByType(vtypeName As String, promptOptions As PromptEntityOptions) As C3DDBObject
        Dim acDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim acEnt As C3DEntity
        Dim Obj As C3DDBObject
        acEnt = SelectEntityByType(vtypeName, promptOptions)
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Obj = acTrans.GetObject(acEnt.Id, OpenMode.ForRead)

                acTrans.Commit()
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage("Error: " & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
        Return Obj
    End Function

End Module
