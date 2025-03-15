Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Module Seleccion
    Private CommandInProgress As Boolean = False

    ' Subscribirse a eventos de comandos al inicializar el módulo
    Public Sub Initialize(Optional acDoc As Document = Nothing)
        If acDoc Is Nothing Then acDoc = Application.DocumentManager.MdiActiveDocument
        AddHandler acDoc.CommandWillStart, AddressOf CommandStarted
        AddHandler acDoc.CommandEnded, AddressOf CommandEnded
        AddHandler acDoc.CommandCancelled, AddressOf CommandEnded
        AddHandler acDoc.CommandFailed, AddressOf CommandEnded
    End Sub

    ' Evento que se dispara cuando un comando empieza
    Private Sub CommandStarted(sender As Object, e As CommandEventArgs)
        CommandInProgress = True
    End Sub

    ' Evento que se dispara cuando un comando termina, falla o es cancelado
    Private Sub CommandEnded(sender As Object, e As CommandEventArgs)
        CommandInProgress = False
    End Sub

    Public Sub SelectByEntity(entidad As Entity)
        ' Obtener el documento y el editor activo de AutoCAD
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        ' Verificar si un comando está en progreso
        If CommandInProgress Then
            acEd.WriteMessage("Un comando está en ejecución. No se puede seleccionar la entidad.")
            Return
        End If

        ' Crear una transacción para acceder a la base de datos
        Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Obtener el ObjectId de la entidad
                Dim id As ObjectId = entidad.ObjectId

                ' Crear un array de ObjectIds con la entidad seleccionada
                Dim objectIds() As ObjectId = {id}

                ' Establecer la selección implícita
                acEd.SetImpliedSelection(objectIds)

                ' Confirmar la transacción
                trans.Commit()
            Catch ex As Exception
                acEd.WriteMessage("Error:" & ex.Message())
                trans.Abort()
            Finally
                If Not trans.IsDisposed Then trans.Dispose()
            End Try
        End Using
    End Sub
End Module
