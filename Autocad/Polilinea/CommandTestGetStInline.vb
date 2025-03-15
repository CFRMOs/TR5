Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class CommandTestGetStInline
    ' Método de prueba que será llamado como un comando de AutoCAD
    <CommandMethod("TestGetStInline")>
    Public Sub TestGetStInline()
        ' Obtener el documento actual y la base de datos
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim editor As Editor = doc.Editor

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Solicitar al usuario que seleccione un alineamiento
                Dim peo As New PromptEntityOptions(vbLf & "Seleccione un alineamiento:")
                peo.SetRejectMessage("Debe seleccionar un alineamiento." & vbLf)
                peo.AddAllowedClass(GetType(Alignment), False)
                Dim per As PromptEntityResult = editor.GetEntity(peo)
                If per.Status <> PromptStatus.OK Then
                    editor.WriteMessage(vbLf & "Comando cancelado.")
                    Return
                End If

                ' Obtener el objeto alineamiento
                Dim alingId As ObjectId = per.ObjectId
                Dim alignment As Alignment = TryCast(trans.GetObject(alingId, OpenMode.ForRead), Alignment)

                ' Solicitar al usuario que seleccione una polilínea
                peo.Message = vbLf & "Seleccione una polilínea:"
                peo.SetRejectMessage("Debe seleccionar una polilínea." & vbLf)
                peo.AddAllowedClass(GetType(Polyline), False)
                per = editor.GetEntity(peo)
                If per.Status <> PromptStatus.OK Then
                    editor.WriteMessage(vbLf & "Comando cancelado.")
                    Return
                End If

                ' Obtener el objeto polilínea
                Dim plId As ObjectId = per.ObjectId
                Dim polyline As Polyline = TryCast(trans.GetObject(plId, OpenMode.ForRead), Polyline)

                ' Solicitar la estación del usuario
                Dim pdo As New PromptDoubleOptions(vbLf & "Ingrese la estación:")
                Dim pdr As PromptDoubleResult = editor.GetDouble(pdo)
                If pdr.Status <> PromptStatus.OK Then
                    editor.WriteMessage(vbLf & "Comando cancelado.")
                    Return
                End If

                Dim station As Double = pdr.Value

                ' Llamar a la función GetStInline de la clase AlignmentLabelHelper
                Dim resultPoint As Point3d = AlignmentLabelHelper.GetStInline(station, polyline, alignment)
                CGPointHelper.AddCGPoint(resultPoint.X, resultPoint.Y, 0)
                ' Mostrar el resultado al usuario
                editor.WriteMessage(vbLf & "El punto más cercano en la polilínea perpendicular al alineamiento en la estación " & station.ToString() & " es: " & resultPoint.ToString())

                ' Confirmar la transacción
                trans.Commit()
            Catch ex As System.Exception
                editor.WriteMessage(vbLf & "Ocurrió un error: " & ex.Message)
                trans.Abort()
            End Try
        End Using
    End Sub
End Class
