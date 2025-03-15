Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
Public Class CSelectionHelper

    Public Shared Function GetDataByEnt(Optional ByRef PromptResult As PromptEntityResult = Nothing, Optional ByRef acEntHandle As String = "") As List(Of (String, String))
        Dim AcEnt As Entity = Nothing
        Dim LData As New List(Of (String, String))()

        ' Get the selected layer from the entity
        Dim selectedLayer As String = ""
        If CLHandle.CheckIfExistHd(acEntHandle) Then
            AcEnt = CLHandle.GetEntityByStrHandle(acEntHandle)
            selectedLayer = AcEnt.Layer
        ElseIf acEntHandle = "" Then
            selectedLayer = CSelectionHelper.GetLayerByEnt(AcEnt, PromptResult)
            'acEntHandle = AcEnt.TabName.ToString()
        End If

        If AcEnt Is Nothing Then
            Return Nothing
        Else
            ' Add the selectedLayer as a tuple to the list
            LData.Add((selectedLayer.ToString(), ""))
            ' Add the selectedLayer as a tuple to the list
            LData.Add((AcEnt.Handle.ToString(), ""))
            Return LData
        End If

    End Function

    ' Función para obtener la capa de una entidad seleccionada
    Public Shared Function GetLayerByEnt(Optional ByRef acEnt As Entity = Nothing, Optional ByRef PromptResult As PromptEntityResult = Nothing) As String
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim TypeOfSelections As String() = {"FeatureLine", "Polyline", "Polyline2d", "Polyline3d", "Parcel", "TinSurface"}

        Do
            Dim promptALingOptions As New PromptEntityOptions("Seleccionar una entidad")
            acEnt = SelectEntity(promptALingOptions, PromptResult)
            If PromptResult.Status = PromptStatus.OK Then
                For Each TypeOfSelection As String In TypeOfSelections
                    If acEnt IsNot Nothing AndAlso TypeName(acEnt) = TypeOfSelection Then
                        Exit Do
                    End If
                Next
            ElseIf PromptResult.Status = PromptStatus.Cancel Then
                Return String.Empty
            End If
        Loop

        Dim selectedLayer As String = acEnt.Layer
        Return acEnt.Layer
    End Function
    ' Función para seleccionar una entidad
    Public Shared Function SelectEntity(Optional ByRef promptOptions As PromptEntityOptions = Nothing, Optional ByRef PromptResult As PromptEntityResult = Nothing) As Entity
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        PromptResult = ed.GetEntity(promptOptions)
        If PromptResult.Status = PromptStatus.OK Then
            Dim tr As Transaction = Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction()
            Using tr
                Dim entity As Entity = TryCast(tr.GetObject(PromptResult.ObjectId, OpenMode.ForRead), Entity)
                tr.Commit()
                Return entity
            End Using
        End If
        Return Nothing
    End Function
    ' Función para seleccionar una entidad y devolver el ObjectId
    Public Shared Function SelectEntityObjectid() As ObjectId 'ByRef promptOptions As PromptEntityOptions, ByRef PromptResult As PromptEntityResult) As ObjectId
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim PromptResult As PromptEntityResult
        Dim promptOptions As New PromptEntityOptions(vbLf & "Seleccione una entidad:")
        PromptResult = ed.GetEntity(PromptOptions)

        If PromptResult.Status = PromptStatus.OK Then
            Return PromptResult.ObjectId
        End If

        Return ObjectId.Null
    End Function

    ' Función para seleccionar una entidad y devolver el handel
    Public Shared Function SelectEntityHandle() As Handle
        ' Obtener el Editor del documento activo
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        ' Crear opciones de solicitud de entidad
        Dim promptOptions As New PromptEntityOptions(vbLf & "Seleccione una entidad:")

        ' Obtener el resultado de la selección de entidad
        Dim promptResult As PromptEntityResult = ed.GetEntity(promptOptions)

        ' Verificar si el usuario seleccionó una entidad correctamente
        If promptResult.Status = PromptStatus.OK Then
            ' Devolver el handle de la entidad seleccionada
            Return promptResult.ObjectId.Handle
        End If

        ' Si no se seleccionó una entidad, devolver TabName.Null
        Return Nothing
    End Function

    ' Función para seleccionar un alineamiento
    Public Shared Function SelectAlignment(ed As Editor, promptMessage As String) As Alignment
        Dim opts As New PromptEntityOptions(vbLf & promptMessage)
        opts.SetRejectMessage(vbLf & "Debe seleccionar un alineamiento.")
        opts.AddAllowedClass(GetType(Alignment), True)

        Dim res As PromptEntityResult = ed.GetEntity(opts)
        If res.Status = PromptStatus.OK Then
            Dim tr As Transaction = ed.Document.TransactionManager.StartTransaction()
            Using tr
                Dim alignment As Alignment = TryCast(tr.GetObject(res.ObjectId, OpenMode.ForRead), Alignment)
                tr.Commit()
                Return alignment
            End Using
        End If
        Return Nothing
    End Function
End Class