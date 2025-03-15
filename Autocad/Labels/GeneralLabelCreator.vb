' Importar todo AutoCAD y Civil 3D
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports CivDbServices = Autodesk.Civil.DatabaseServices
Imports CivStyles = Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.Civil.ApplicationServices

Public Class GeneralLabelCreator
    ' Función para crear un GeneralSegmentLabel basado en una polilínea y un punto
    Public Sub CrearGeneralSegmentLabel(ByVal pl As Polyline, ByVal point As Point2d, ByVal styleName As String)
        ' Obtener el documento y la base de datos de AutoCAD/Civil 3D
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim civilDoc As CivilDocument = CivilApplication.ActiveDocument

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Validar si el objeto Polyline no es nulo
                If pl Is Nothing OrElse pl.IsDisposed Then
                    ed.WriteMessage(vbLf & "Error: La polilínea no es válida o ha sido eliminada.")
                    Return
                End If

                ' Obtener el ObjectId de la polilínea
                Dim plId As ObjectId = pl.ObjectId

                ' Obtener el estilo de etiqueta por su nombre usando Civil 3D
                Dim labelLineStyleId As ObjectId = GetChildLabelStyleByName(styleName, "LP.Tag.1")

                ' Validar si se encontró el estilo de etiqueta
                If labelLineStyleId.IsNull Then
                    ed.WriteMessage(vbLf & "No se encontró el estilo de etiqueta especificado.")
                    Return
                End If

                ' Calcular el ratio para el punto proporcionado en la polilínea
                ' Para simplificar, asumimos que el ratio es un valor entre 0 y el número de segmentos de la polilínea
                Dim ratio As Double = GetRatioForPointOnPolyline(pl, point)

                ' Crear la etiqueta usando el método estático Create
                Dim labelId As ObjectId = CivDbServices.GeneralSegmentLabel.Create(plId, ratio)

                Dim GSL As GeneralSegmentLabel = trans.GetObject(labelId, OpenMode.ForWrite)
                'GSL.LabelLocation = New Point3d(point.X, point.Y, 0)
                'GSL.StyleName = styleName
                GSL.StyleId = labelLineStyleId
                ' Confirmar la transacción
                trans.Commit()
                ed.WriteMessage(vbLf & "Etiqueta de segmento general creada correctamente en la polilínea.")
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
                trans.Abort() ' Deshacer la transacción en caso de error
            Finally
                If Not trans.IsDisposed() Then trans.Dispose()
            End Try
        End Using
    End Sub
    ' Función para obtener el estilo hijo dentro de un estilo principal
    Private Function GetChildLabelStyleByName(ByVal parentStyleName As String, ByVal childStyleName As String) As ObjectId
        Dim labelStyleId As ObjectId = ObjectId.Null
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim curDoc As CivilDocument = CivilApplication.ActiveDocument

        ' Verificar que el nombre del estilo no esté vacío
        If String.IsNullOrEmpty(parentStyleName) Or String.IsNullOrEmpty(childStyleName) Then
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error: El nombre del estilo o del estilo hijo no puede estar vacío.")
            Return ObjectId.Null
        End If

        ' Iniciar una transacción para buscar el estilo principal y sus hijos
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener la colección de estilos de etiquetas de segmento general
                Dim labelStyleCollection As CivStyles.LabelStyleCollection = curDoc.Styles.LabelStyles.GeneralLineLabelStyles

                ' Verificar que la colección de estilos no esté vacía
                If labelStyleCollection.Count = 0 Then
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error: No hay estilos de etiquetas disponibles en el documento.")
                    Return ObjectId.Null
                End If

                ' Buscar el estilo principal por su nombre
                Dim parentLabelStyleId As ObjectId = ObjectId.Null
                For Each styleId As ObjectId In labelStyleCollection
                    Dim labelStyle As CivStyles.LabelStyle = trans.GetObject(styleId, OpenMode.ForRead)
                    If labelStyle.Name = parentStyleName Then
                        parentLabelStyleId = styleId
                        Exit For
                    End If
                Next

                ' Verificar si se encontró el estilo principal
                If parentLabelStyleId.IsNull Then
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & $"Error: No se encontró el estilo principal '{parentStyleName}'.")
                    Return ObjectId.Null
                End If

                ' Obtener el estilo principal para buscar estilos hijos
                Dim parentLabelStyle As CivStyles.LabelStyle = trans.GetObject(parentLabelStyleId, OpenMode.ForRead)

                ' Verificar si el estilo principal tiene hijos
                If parentLabelStyle.ChildrenCount > 0 Then
                    ' Iterar a través de los estilos hijos
                    For i As Integer = 0 To parentLabelStyle.ChildrenCount - 1
                        Dim childLabelStyleId As ObjectId = parentLabelStyle(i)
                        Dim childLabelStyle As CivStyles.LabelStyle = trans.GetObject(childLabelStyleId, OpenMode.ForRead)

                        ' Verificar si el nombre del estilo hijo coincide
                        If childLabelStyle.Name = childStyleName Then
                            labelStyleId = childLabelStyleId
                            Exit For
                        End If
                    Next
                End If

                ' Verificar si no se encontró el estilo hijo
                If labelStyleId.IsNull Then
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & $"Error: No se encontró el estilo hijo '{childStyleName}' dentro del estilo '{parentStyleName}'.")
                End If

                trans.Commit()
            Catch ex As Exception
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error al buscar el estilo hijo: " & ex.Message)
            End Try
        End Using

        Return labelStyleId
    End Function

    ' Función auxiliar para calcular el ratio basado en un punto en la polilínea
    Private Function GetRatioForPointOnPolyline(ByVal pl As Polyline, ByVal point As Point2d) As Double
        ' Inicializamos las variables para almacenar el ratio y la distancia acumulada
        Dim ratio As Double = 0.0
        Dim totalLength As Double = pl.Length
        Dim cumulativeLength As Double = 0.0

        ' Recorremos todos los segmentos de la polilínea
        For i As Integer = 0 To pl.NumberOfVertices - 2
            ' Obtener los dos puntos del segmento actual
            Dim startPoint As Point2d = pl.GetPoint2dAt(i)
            Dim endPoint As Point2d = pl.GetPoint2dAt(i + 1)

            ' Calcular la distancia del segmento actual
            Dim segmentLength As Double = startPoint.GetDistanceTo(endPoint)

            ' Verificar si el punto proporcionado está en este segmento (en su proyección)
            Dim segmentLine As LineSegment2d = New LineSegment2d(startPoint, endPoint)
            If segmentLine.IsOn(point) Then
                ' Si el punto está en este segmento, calculamos la distancia desde el punto inicial
                Dim partialLength As Double = startPoint.GetDistanceTo(point)
                ratio = (cumulativeLength + partialLength) / totalLength
                Exit For
            End If

            ' Acumular la longitud del segmento
            cumulativeLength += segmentLength
        Next

        ' Retornar el ratio calculado
        Return ratio
    End Function

End Class
