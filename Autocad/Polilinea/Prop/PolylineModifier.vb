Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors

Public Class PolylineModifier

    ''' <summary>
    ''' Cambia las propiedades de una polilínea basada en su handle.
    ''' </summary>
    ''' <param name="polylineHandle">El handle de la polilínea.</param>
    ''' <param name="lineType">El tipo de línea a asignar.</param>
    ''' <param name="lineTypeScale">La escala del tipo de línea.</param>
    ''' <param name="globalWidth">El ancho global de la polilínea.</param>
    Public Sub ChangePolylineProperties(polylineHandle As String, lineType As String, lineTypeScale As Double, globalWidth As Double, Optional CColor As Color = Nothing)
        ' Obtener el documento y la base de datos activa
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Iniciar una transacción
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Convertir el handle a ObjectId
                Dim handle As Handle = New Handle(Convert.ToInt64(polylineHandle, 16))
                Dim polylineId As ObjectId = db.GetObjectId(False, handle, 0)

                ' Abrir la polilínea para su modificación
                Dim polyline As Polyline = TryCast(tr.GetObject(polylineId, OpenMode.ForWrite), Polyline)

                ' Validar si se obtuvo correctamente la polilínea
                If polyline IsNot Nothing Then


                    ' Cambiar la escala del tipo de línea
                    polyline.LinetypeScale = lineTypeScale

                    ' Cambiar el ancho global
                    polyline.ConstantWidth = globalWidth

                    'cambiar cofiguracion para que se muestre el line type entre vertices 
                    polyline.Plinegen = True
                    If CColor Is Nothing Then
                        polyline.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256)

                    Else
                        polyline.Color = CColor 'Color.FromColorIndex(ColorMethod.ByBlock, 0)

                    End If
                    ' Cambiar el tipo de línea
                    polyline.Linetype = lineType

                    ' Confirmar los cambios
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Se cambiaron las propiedades de la polilínea correctamente.")
                Else
                    ed.WriteMessage(vbLf & "No se encontró la polilínea con el handle proporcionado.")
                End If
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
            Finally
                ' Finalizar la transacción
                tr.Dispose()
            End Try
        End Using
    End Sub
End Class
