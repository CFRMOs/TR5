Imports System.Diagnostics
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.EditorInput
Public Class AutoCADHandler
    Private WithEvents AcadApp As AcadApplication ' Cambia de acadApp a AcadApp

    'Private ReadOnly openDocument As AcadDocument

    Public Sub New()
        'Inicializa la aplicación de AutoCAD y obtiene el documento activo
        Try
            Debug.WriteLine("Iniciando AutoCAD...")
            AcadApp = CType(Application.AcadApplication, AcadApplication)
            AcadApp.Visible = True
            'openDocument = AcadApp.ActiveDocument
            Debug.WriteLine("AutoCAD iniciado y documento activo obtenido.")
        Catch ex As Exception
            Debug.WriteLine("Error al iniciar AutoCAD: " & ex.Message)
            Throw New Exception("No se pudo iniciar AutoCAD: " & ex.Message)
        End Try
    End Sub

    'Public Sub New(openDocument As AcadDocument)
    '    Me.openDocument = openDocument
    'End Sub

    Public Sub TestCopyEntityFromClosedDwg(dwgPath As String, handle As String)
        'Step 1: Clone the entity from the closed DWG file
        Dim clonedEntity As Entity = CloneEntityFromClosedDwg(dwgPath, handle)

        'Step 2 Add the cloned entity to the active document
        If clonedEntity IsNot Nothing Then
            AddEntityToActiveDocument(clonedEntity)
        Else
            Application.ShowAlertDialog("No se pudo clonar la entidad.")
        End If
    End Sub

    Private Function CloneEntityFromClosedDwg(dwgPath As String, handle As String) As Entity
        'This Function() clones an entity from a closed DWG file
        Using db As New Database(False, True)
            Try
                db.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, True, "")

                'Start the read transaction
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim objId As ObjectId = db.GetObjectId(False, New Handle(Convert.ToInt64(handle, 16)), 0)
                    Dim entity As Entity = TryCast(tr.GetObject(objId, OpenMode.ForRead), Entity)

                    If entity IsNot Nothing Then
                        Debug.WriteLine("Entidad encontrada. Clonando...")
                        Dim clonedEntity As Entity = entity.Clone()
                        tr.Commit()
                        Return clonedEntity
                    Else
                        Debug.WriteLine("No se encontró la entidad con el handle especificado.")
                        Application.ShowAlertDialog("No se pudo encontrar la entidad con el handle especificado.")
                        Return Nothing
                    End If
                End Using

            Catch ex As Exception
                Debug.WriteLine("Error al clonar la entidad: " & ex.Message)
                Application.ShowAlertDialog("Error al clonar la entidad: " & ex.Message)
                Return Nothing
            End Try
        End Using

        'Default return In Case something goes wrong unexpectedly
        Return Nothing
    End Function
    Private Sub AddEntityToActiveDocument(clonedEntity As Entity)
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor
        'This Function() adds a cloned entity to the ModelSpace of the active document
        Try
            'Obtener la base de datos del documento abierto

            'Start the write transaction
            Using destTr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                AcDoc.LockDocument()
                'Open the BlockTable for read
                Dim blockTable As BlockTable = destTr.GetObject(HostApplicationServices.WorkingDatabase.BlockTableId, OpenMode.ForRead)

                'Get the ModelSpace BlockTableRecord for write
                Dim modelSpace As BlockTableRecord = destTr.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                'Add the cloned entity to the ModelSpace
                modelSpace.AppendEntity(clonedEntity)
                destTr.AddNewlyCreatedDBObject(clonedEntity, True)

                destTr.Commit()

                Debug.WriteLine("Entidad clonada y añadida al documento abierto.")
                Application.ShowAlertDialog("Entidad copiada exitosamente.")
            End Using

        Catch ex As Exception
            Debug.WriteLine("Error al agregar la entidad al documento: " & ex.Message)
            Application.ShowAlertDialog("Error al agregar la entidad al documento: " & ex.Message)
        End Try
    End Sub
    Public Sub CopyEntityFromClosedDwg(dwgPath As String, handle As String, Optional ThisDrawing As Document = Nothing, Optional ByRef ResultedHandel As String = "")
        'Usar Database para abrir el archivo DWG cerrado
        Using db As New Database(False, True)
            Try
                db.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, True, "")

                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim objId As ObjectId = db.GetObjectId(False, New Handle(Convert.ToInt64(handle, 16)), 0)
                    Dim entity As Entity = tr.GetObject(objId, OpenMode.ForRead)

                    If entity IsNot Nothing Then
                        Debug.WriteLine("Entidad encontrada. Clonando...")
                        Dim AcDoc As Document
                        'Obtener la base de datos del documento abierto
                        If ThisDrawing IsNot Nothing Then
                            AcDoc = ThisDrawing
                        Else
                            AcDoc = Application.DocumentManager.MdiActiveDocument
                        End If
                        Dim AcCurD As Database = HostApplicationServices.WorkingDatabase

                        Using destTr As Transaction = AcCurD.TransactionManager.StartTransaction()

                            'Clonar la entidad y añadirla al espacio modelo del documento abierto
                            Dim copiedEntity As Entity = entity.Clone()
                            AcDoc.LockDocument()
                            Dim blockTable As BlockTable = destTr.GetObject(AcCurD.BlockTableId, OpenMode.ForRead)
                            Dim modelSpace As BlockTableRecord = destTr.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                            modelSpace.AppendEntity(copiedEntity)
                            destTr.AddNewlyCreatedDBObject(copiedEntity, True)
                            ResultedHandel = copiedEntity.Handle.ToString()
                            AcadZoomManager.ZoomToEntity(copiedEntity)
                            destTr.Commit()
                        End Using
                        Debug.WriteLine("Entidad clonada y añadida al documento abierto.")
                        'Application.ShowAlertDialog("Entidad copiada exitosamente.")
                    Else
                        Debug.WriteLine("No se encontró la entidad con el handle especificado.")
                        Application.ShowAlertDialog("No se pudo encontrar la entidad con el handle especificado.")
                    End If

                    tr.Commit()
                End Using

            Catch ex As Exception
                Debug.WriteLine("Error durante el proceso: " & ex.Message)
                Application.ShowAlertDialog("Error: " & ex.Message)
            End Try
        End Using
    End Sub
End Class

Public Class CopyFromClose
    <CommandMethod("CopyFromClosedDwg")>
    Public Sub CopyFromClosedDwgCommand()
        Try
            Debug.WriteLine("Iniciando prueba de AutoCADHandler.")
            Dim handler As New AutoCADHandler()
            Dim dwgPath As String = "D:\Desktop\Typsa - Las Placetas\MED-HLP-ACC4-DRE-012-R4.dwg" ' Reemplaza con la ruta real
            Dim handle As String = "18A8F" ' Reemplaza con el handle real

            handler.CopyEntityFromClosedDwg(dwgPath, handle)
            Debug.WriteLine("Prueba completada.")

        Catch ex As Exception
            Debug.WriteLine("Error en CopyFromClosedDwgCommand: " & ex.Message)
            Application.ShowAlertDialog("Error: " & ex.Message)
        End Try
    End Sub
End Class
