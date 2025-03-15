' Importamos los espacios de nombres necesarios
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices

Public Class GeneralSegmentLabelHelper
    ' Esta clase devuelve los componentes de un label 
    ' en particular, el contenido de los componentes de texto.

    Public Shared Function GetLabelText(label As GeneralSegmentLabel) As String
        ' Obtenemos el documento activo y el editor
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Verificamos si el label es válido
        If label Is Nothing Then
            ed.WriteMessage(vbLf & "El label proporcionado no es válido.")
            Return String.Empty
        End If

        ' Iniciar una transacción para acceder a la base de datos
        Using tr As Transaction = doc.TransactionManager.StartTransaction()
            ' Obtener el ObjectId del GeneralSegmentLabel
            ed.WriteMessage(vbLf & "Label ObjectId: " & label.ObjectId.ToString())

            ' Obtener la capa del GeneralSegmentLabel
            Dim layerName As String = label.Layer
            ed.WriteMessage(vbLf & "Label Layer: " & layerName)

            ' Obtener los identificadores de los componentes de texto del label
            Dim textComponentIds As ObjectIdCollection = label.GetTextComponentIds()

            ' Iterar sobre los componentes de texto para obtener su contenido
            Dim textComponentId As ObjectId = textComponentIds(0)
            Dim textEntity As DBObject = tr.GetObject(textComponentId, OpenMode.ForRead)
            Dim text As String = label.GetTextComponentOverride(textComponentId)

            ed.WriteMessage(vbLf & "Texto del componente: " & text)
            Return text
        End Using
    End Function

    ' Nueva función para explotar y obtener textos de DBText asociados al GeneralSegmentLabel
    Public Shared Function GetExplodedLabelText(label As GeneralSegmentLabel) As String
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim entityText As New System.Text.StringBuilder()

        If label Is Nothing Then
            ed.WriteMessage(vbLf & "El label proporcionado no es válido.")
            Return String.Empty
        End If

        Using tr As Transaction = doc.TransactionManager.StartTransaction()
            Try
                ' Obtener el ObjectId del label
                Dim labelId As ObjectId = label.ObjectId

                ' Obtener la entidad GeneralSegmentLabel completa
                Dim labelEntity As Entity = TryCast(tr.GetObject(labelId, OpenMode.ForRead), Entity)

                If labelEntity IsNot Nothing Then
                    ' Explota el GeneralSegmentLabel completo
                    Dim explodedObjects As List(Of DBObject) = FullExplode(labelEntity)

                    ' Procesar cada entidad explotada
                    For Each explodedObj As Entity In explodedObjects
                        ' Verificar si es un DBText y extraer el texto
                        If TypeOf explodedObj Is DBText Then
                            Dim text As DBText = CType(explodedObj, DBText)
                            entityText.AppendLine(text.TextString)
                        End If
                    Next
                End If
                tr.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error al explotar el label: " & ex.Message)
            End Try
        End Using
        ' Devolver el texto recopilado
        Return entityText.ToString().Trim()
    End Function

    ' Función que explota completamente una entidad y devuelve una lista de objetos DBObject
    Private Shared Function FullExplode(ent As Entity) As List(Of DBObject)
        Dim fullList As New List(Of DBObject)()

        Try
            Dim explodedObjects As New DBObjectCollection()
            ent.Explode(explodedObjects)

            ' Iterar sobre las entidades explotadas
            For Each explodedObj As DBObject In explodedObjects
                If TypeOf explodedObj Is BlockReference OrElse TypeOf explodedObj Is MText Then
                    ' Volver a explotar si es BlockReference o MText
                    fullList.AddRange(FullExplode(TryCast(explodedObj, Entity)))
                Else
                    fullList.Add(explodedObj)
                End If
            Next
        Catch ex As Exception
            ' Manejo de errores durante la explosión
        End Try

        Return fullList
    End Function
End Class
