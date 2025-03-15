Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Module SolidCreation
    <CommandMethod("CmdCS")> Public Sub CreateSolid()
        ' Función para crear un sólido en AutoCAD
        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        If doc Is Nothing Then
            Return
        End If

        ' Empezar una transacción en la base de datos
        Using trans As Transaction = doc.Database.TransactionManager.StartTransaction()
            Try
                ' Abrir el espacio de modelos para escritura
                Dim blockTable As BlockTable = trans.GetObject(doc.Database.BlockTableId, OpenMode.ForRead)
                Dim modelSpace As BlockTableRecord = trans.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                ' Crear un sólido esférico
                Dim center As New Point3d(0, 0, 0)
                Dim radius As Double = 5.0
                Dim sphere As New Solid3d()
                sphere.CreateSphere(radius)
                sphere.TransformBy(Matrix3d.Displacement(center.GetAsVector()))

                ' Agregar el sólido al espacio de modelos
                modelSpace.AppendEntity(sphere)
                trans.AddNewlyCreatedDBObject(sphere, True)

                ' Completar la transacción
                trans.Commit()

                ' Notificar al usuario
                doc.Editor.WriteMessage("Sólido creado correctamente en AutoCAD.")

            Catch ex As Exception
                ' Manejar cualquier error
                doc.Editor.WriteMessage("Error al crear el sólido en AutoCAD: " & ex.Message)
                trans.Abort()
            End Try
        End Using
    End Sub
End Module
