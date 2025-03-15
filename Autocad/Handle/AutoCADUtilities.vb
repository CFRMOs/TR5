Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Public Class AutoCADUtilities

    ' Esta función oculta una entidad en AutoCAD utilizando su ObjectID
    Public Shared Sub EntidadVisibility(ByVal objectId As ObjectId, V As Boolean)
        ' Conexión a la aplicación AutoCAD activa
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        ' Iniciar una transacción
        Using tr As Transaction = AcDoc.Database.TransactionManager.StartTransaction()
            Try
                ' Obtener la referencia a la entidad usando su ObjectId
                Dim entidad As Entity = CType(tr.GetObject(objectId, OpenMode.ForWrite), Entity)

                ' Cambiar la visibilidad de la entidad a False (oculto)
                entidad.Visible = V

                ' Confirmar la transacción
                tr.Commit()
            Catch ex As Exception
                ' Manejar cualquier error que ocurra
                AcEd.WriteMessage("Error: " & ex.Message)
            End Try
        End Using
    End Sub
End Class
