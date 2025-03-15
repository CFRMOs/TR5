Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Module TraceCenterLine
    Public Class Commands
        <CommandMethod("TraceCenterLine")>
        Public Sub TraceCenterLine()
            ' Obtener la aplicación y el editor activos de AutoCAD
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ' Pedir al usuario que seleccione el área con el contorno del hatch
            Dim promptResult As PromptEntityResult = ed.GetEntity("Seleccione el área con el contorno del hatch:")
            If promptResult.Status <> PromptStatus.OK Then
                Return
            End If

            ' Abrir la transacción para modificar la base de datos
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Abrir el contorno del hatch seleccionado
                    Dim ent As Entity = trans.GetObject(promptResult.ObjectId, OpenMode.ForRead)
                    If TypeOf ent Is Hatch Then
                        Dim hatch As Hatch = CType(ent, Hatch)

                        ' Obtener el centro del área del contorno del hatch
                        Dim area As Double = hatch.Area
                        Dim centroid As Point3d ' = hatch.GeometricExtents.MinPoint.GetMidPointTo(hatch.GeometricExtents.MaxPoint)

                        ' Calcular los puntos para dibujar la línea longitudinalmente por el centro del área
                        Dim startPoint As New Point3d(centroid.X, centroid.Y - area / 2, centroid.Z)
                        Dim endPoint As New Point3d(centroid.X, centroid.Y + area / 2, centroid.Z)

                        ' Dibujar la línea en el modelo de espacio
                        Dim line As New Line(startPoint, endPoint)
                        Dim btr As BlockTableRecord = CType(trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        btr.AppendEntity(line)
                        trans.AddNewlyCreatedDBObject(line, True)

                        ' Guardar los cambios
                        trans.Commit()
                        ed.WriteMessage("Línea trazada longitudinalmente por el centro del área del contorno del hatch.")
                    Else
                        ed.WriteMessage("El objeto seleccionado no es un hatch.")
                    End If
                Catch ex As Exception
                    ed.WriteMessage("Error al trazar la línea: " & ex.Message)
                    trans.Abort()
                End Try
            End Using
        End Sub
    End Class

End Module
