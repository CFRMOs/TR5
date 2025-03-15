' Clase con un comando personalizado para poner a prueba la función
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class LabelTestCommands
    ' Comando personalizado para crear una etiqueta de segmento general
    <CommandMethod("CREAR_ETIQUETA_SEGMENTO")>
    Public Sub CrearEtiquetaSegmentoCommand()
        ' Obtener el documento activo y su editor
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        ' Iniciar transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Pedir al usuario que seleccione una polilínea
                Dim opts As New PromptEntityOptions(vbLf & "Seleccione una polilínea:")
                opts.SetRejectMessage("Debe seleccionar una polilínea.")
                opts.AddAllowedClass(GetType(Polyline), False)
                Dim res As PromptEntityResult = ed.GetEntity(opts)

                ' Validar si se seleccionó algo
                If res.Status <> PromptStatus.OK Then
                    ed.WriteMessage(vbLf & "No se seleccionó una polilínea.")
                    Return
                End If

                ' Obtener la polilínea seleccionada
                Dim pline As Polyline = trans.GetObject(res.ObjectId, OpenMode.ForRead)

                ' Pedir al usuario que indique un punto
                Dim ppr As PromptPointResult = ed.GetPoint(vbLf & "Indique un punto sobre la polilínea:")
                If ppr.Status <> PromptStatus.OK Then
                    ed.WriteMessage(vbLf & "No se seleccionó un punto.")
                    Return
                End If
                Dim point As Point3d = ppr.Value

                ' Pedir el nombre del estilo de etiqueta
                'Dim pstrOpts As New PromptStringOptions(vbLf & "Ingrese el nombre del estilo de etiqueta:")
                'Dim styleResult As PromptResult = ed.GetString(pstrOpts)
                'If styleResult.Status <> PromptStatus.OK Then
                '    ed.WriteMessage(vbLf & "No se ingresó un nombre de estilo.")
                '    Return
                'End If
                Dim styleName As String = "LP" 'styleResult.StringResult

                ' Llamar a la función para crear la etiqueta
                Dim point2d As New Point2d(point.X, point.Y)
                Dim labelCreator As New GeneralLabelCreator()
                labelCreator.CrearGeneralSegmentLabel(pline, point2d, styleName)

                ' Confirmar la transacción
                trans.Commit()
                ed.WriteMessage(vbLf & "Etiqueta creada correctamente.")
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
                trans.Abort() ' Deshacer la transacción en caso de error
            End Try
        End Using
    End Sub
End Class