'station offset labels civil 3d vb.net adding
'en vb.net  crear un funccion para crear un los lables offset de un punto dado con relacion a un alineamiento en civil 3D
'debe de agregar la ootion de seleccionar un punto que sera  en el cual crear el label

Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
'Imports Autodesk.Civil.LabelContentDisplayType


Public Class StationOffsetLabelCreator

    ' Función para crear una etiqueta de "Station Offset" basado en un punto dado
    Public Sub CrearStationOffsetLabel(alignmentName As String)
        ' Obtener el documento y la base de datos de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim civilDoc As CivilDocument = CivilApplication.ActiveDocument

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Pedir al usuario que seleccione un punto en el dibujo
            Dim ppr As PromptPointResult = ed.GetPoint("Seleccione un punto para la etiqueta de Station Offset: ")
            If ppr.Status <> PromptStatus.OK Then
                ed.WriteMessage("Punto no válido.")
                Return
            End If
            Dim puntoReferencia As Point3d = ppr.Value

            ' Obtener la alineación por nombre
            Dim alignmentId As ObjectId = GetAlignmentIdByName(civilDoc, alignmentName, trans)
            If alignmentId = ObjectId.Null Then
                ed.WriteMessage("Alineación no encontrada.")
                Return
            End If

            ' Abrir la alineación para lectura/escritura
            Dim alignment As Alignment = CType(trans.GetObject(alignmentId, OpenMode.ForWrite), Alignment)

            ' Calcular la estación y el offset del punto seleccionado
            Dim estacion As Double
            Dim offset As Double
            alignment.StationOffset(puntoReferencia.X, puntoReferencia.Y, estacion, offset)

            ' Obtener el estilo de la etiqueta de "Station Offset" (deberás ajustar según tu dibujo)
            Dim labelStyleId As ObjectId = GetStationOffsetLabelStyleId()

            If labelStyleId.IsNull Then
                'si no se encontro el label deberia de crearse 
                'se creara una funcion nueva para la creacion del labelStyle
                ed.WriteMessage("No se encontró el estilo de etiqueta de Station Offset.")
                Return
            End If

            ' Añadir la etiqueta de Station Offset a la alineación
            'alignment.AddLabel(Autodesk.Civil.LabelType.StationOffset, estacion, offset)

            ' Confirmar la transacción
            trans.Commit()
        End Using
    End Sub

    ' Función auxiliar para obtener el ObjectId de una alineación por nombre
    Private Function GetAlignmentIdByName(civilDoc As CivilDocument, alignmentName As String, trans As Transaction) As ObjectId
        For Each alignId As ObjectId In civilDoc.GetAlignmentIds()
            Dim alignment As Alignment = CType(trans.GetObject(alignId, OpenMode.ForRead), Alignment)
            If alignment.Name = alignmentName Then
                Return alignId
            End If
        Next
        Return ObjectId.Null
    End Function

    ' Función auxiliar para obtener el estilo de etiqueta de Station Offset
    ' Aquí puedes obtener o crear el estilo de etiqueta de Station Offset
    ' En este ejemplo se devuelve un ObjectId nulo. Aquí debes retornar un estilo válido.
    ' Deberás ajustar este método según cómo manejes los estilos de etiqueta en tu proyecto.
    Private Function GetStationOffsetLabelStyleId() As ObjectId
        ' Obtener el documento Civil 3D activo
        Dim civilDoc As CivilDocument = CivilApplication.ActiveDocument
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Definir el nombre del estilo que estamos buscando
        Dim stationOffsetLabelStyleName As String = "Station Offset" ' Cambia este nombre por el que estés utilizando

        ' Inicializar el ObjectId que se retornará
        Dim labelStyleId As ObjectId = ObjectId.Null

        Try
            ' Obtener la colección de estilos de etiquetas de alineación
            Dim styles As LabelStyleCollection = civilDoc.Styles.LabelStyles.AlignmentLabelStyles.StationOffsetLabelStyles

            ' Buscar el estilo de etiqueta por nombre
            For Each styleId As ObjectId In styles
                ' Abrir el estilo de etiqueta para leerlo
                Dim labelStyle As LabelStyle = CType(styleId.GetObject(OpenMode.ForRead), LabelStyle)
                ' Verificar si el nombre coincide
                If labelStyle.Name = stationOffsetLabelStyleName Then
                    labelStyleId = styleId
                    Exit For
                End If
            Next

            ' Verificar si se encontró el estilo
            If labelStyleId = ObjectId.Null Then
                labelStyleId = CrearStationOffsetLabelStyle()
                ed.WriteMessage(vbLf & "No se encontró el estilo de etiqueta de Station Offset: " & stationOffsetLabelStyleName)
            End If

        Catch ex As Exception
            ed.WriteMessage(vbLf & "Error al obtener el estilo de etiqueta de Station Offset: " & ex.Message)
        End Try

        ' Retornar el ObjectId del estilo encontrado o Null si no se encontró
        Return labelStyleId
    End Function
    ' Función auxiliar para crear un estilo de etiqueta de Station Offset
    Private Function CrearStationOffsetLabelStyle() As ObjectId
        Dim civilDoc As CivilDocument = CivilApplication.ActiveDocument
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Try
            ' Crear el estilo de etiqueta de Station Offset
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                ' Obtener la colección de estilos de etiquetas de Station Offset
                Dim labelStyles As LabelStyleCollection = civilDoc.Styles.LabelStyles.AlignmentLabelStyles.StationOffsetLabelStyles

                ' Crear un nuevo estilo de etiqueta utilizando el método Add de la colección
                Dim newLabelStyleId As ObjectId = labelStyles.Add("Station Offset")

                ' Obtener el nuevo estilo de etiqueta
                Dim newLabelStyle As LabelStyle = trans.GetObject(newLabelStyleId, OpenMode.ForWrite)

                ' Añadir el componente de texto al estilo
                Dim textComponentId As ObjectId = newLabelStyle.AddComponent("Text1", LabelStyleComponentType.Text)

                ' Obtener el componente de texto recién creado
                Dim textComponent As LabelStyleTextComponent = trans.GetObject(textComponentId, OpenMode.ForWrite)

                ' Configurar el contenido del texto directamente
                textComponent.Text.Contents.Value = "CU-" & 2 & "
                                                     STA=<[Station Value(Uft|FS|P2|RN|AP|Sn|TP|B2|EN|W0|OF)]>                                                           "

                ' Configurar propiedades de estilo de texto
                Dim textStyleId As ObjectId = civilDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles(0) ' Usa el primer estilo de texto, puedes personalizar esto
                'textComponent.StyleText = textStyleId  ' Establecer el estilo de texto ' Establecer el estilo de texto
                'textComponent.Text.Height = 0.1  ' Altura del texto
                ' Configurar propiedades de estilo de texto a través de StyleText
                'Dim styleText As StyleText = textComponent.StyleText
                'styleText.Height = 0.1  ' Establecer la altura del texto
                'styleText.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3) ' Cambia el color a verde
                ' Confirmar la transacción
                trans.AddNewlyCreatedDBObject(newLabelStyle, True)
                trans.Commit()

                ' Devolver el ObjectId del nuevo estilo
                Return newLabelStyle.ObjectId
            End Using

        Catch ex As Exception
            ed.WriteMessage(vbLf & "Error al crear el estilo de etiqueta de Station Offset: " & ex.Message)
            Return ObjectId.Null
        End Try
    End Function



End Class
Public Class CustomCommands

    ' Definir un comando AutoCAD llamado "TestStationOffsetLabel"
    <CommandMethod("TestStationOffsetLabel")>
    Public Sub TestStationOffsetLabelCommand()
        ' Obtener el documento de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Try
            ' Pedir al usuario que ingrese el nombre del alineamiento
            Dim promptOptions As New PromptStringOptions(vbLf & "Ingrese el nombre del alineamiento: ") With {
                .AllowSpaces = True
            }
            Dim result As PromptResult = ed.GetString(promptOptions)
            'EJE ACCESO 05
            ' Verificar si la entrada fue correcta
            If result.Status <> PromptStatus.OK Then
                ed.WriteMessage("Nombre de alineamiento no válido.")
                Return
            End If

            ' Obtener el nombre del alineamiento ingresado por el usuario
            Dim alignmentName As String = result.StringResult

            ' Instanciar la clase StationOffsetLabelCreator
            Dim labelCreator As New StationOffsetLabelCreator()

            ' Llamar a la función para crear la etiqueta de Station Offset
            labelCreator.CrearStationOffsetLabel(alignmentName)

        Catch ex As System.Exception
            ed.WriteMessage(vbLf & "Error: " & ex.Message)
        End Try
    End Sub

End Class