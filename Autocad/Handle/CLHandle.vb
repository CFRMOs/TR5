Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime


Public Class CLHandle
    Public Shared Function Chandle(hdString As String) As Handle
        ' Return an invalid TabName (i.e., TabName with value 0) if the string is null or empty
        hdString = hdString?.Replace("-", "")
        If String.IsNullOrEmpty(hdString) Then
            Return New Handle(0)

        End If

        Try
            ' Convert the string from hexadecimal to a long integer
            Dim Ln As Long = Convert.ToInt64(hdString, 16)
            ' Create and return a TabName based on the long value
            Return New Handle(Ln)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            ' Return an invalid TabName on error
            Return New Handle(0)
        End Try
    End Function


    ' Get ObjectId by handle string
    Public Shared Function GetObjIdBStr(hdString As String) As ObjectId
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Dim hd As Handle = Chandle(hdString)
        Dim id As ObjectId = AcCurD.GetObjectId(False, hd, 0)

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim ent As Entity = TryCast(trans.GetObject(id, OpenMode.ForRead), Entity)
                trans.Commit()
                Return ent.Id
            Catch ex As Exception
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
            Finally
                trans.Dispose()
            End Try
        End Using
    End Function

    Public Function GetObjectByHandle(hn As Handle) As DBObject

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcEd As Editor = doc.Editor
        Dim AcCurD As Database = doc.Database

        Dim id As ObjectId = AcCurD.GetObjectId(False, hn, 0)

        ' Finally let's open the object and erase it
        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim obj As DBObject = trans.GetObject(id, OpenMode.ForWrite)
                trans.Commit()
                Return obj
            Catch ex As Exception
                AcEd.WriteMessage(vbCrLf & "Error: " & ex.Message)
                trans.Abort()
                Return Nothing
            Finally
                trans.Dispose()
            End Try
        End Using
    End Function
    Public Shared Function GetEntityIdByStrHandle(hdString As String) As ObjectId
        Dim Handle As Handle = Chandle(hdString)
        Return GetEntityIdByHandle(Handle)
    End Function
    Public Shared Function GetEntityIdByHandle(hn As Handle) As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Try
            ' Attempt to get the ObjectId for the handle
            Dim obj As ObjectId = db.GetObjectId(False, hn, 0)
            Return obj
        Catch ex As Exception
            ' Write an error message to the command line
            ed.WriteMessage(vbCrLf & "Error: " & ex.Message)

            ' Return ObjectId.Null to indicate failure
            Return ObjectId.Null
        End Try
    End Function


    Public Shared Function GetEntityByStrHandle(ByRef hdString As String) As Entity
        Dim Handle As Handle = Chandle(hdString)
        Return GetEntityByHandle(Handle)
    End Function
    Public Shared Function CheckIfExistHd(hdString As String) As Boolean
        On Error Resume Next

        If hdString = "" Then Return False
        If GetEntityByStrHandle(hdString) IsNot Nothing Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function GetEntityByHandle(hn As Handle) As Entity
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        ' Get ObjectId by handle
        Dim id As ObjectId = GetEntityIdByHandle(hn)

        ' Return nothing if the ObjectId is null or invalid
        If id.IsNull Then
            Return Nothing
        End If

        ' Open a transaction to access the entity
        Using tr As Transaction = doc.TransactionManager.StartTransaction()
            Try
                ' Safely attempt to cast the object to an Entity
                Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)

                If ent IsNot Nothing Then
                    tr.Commit()
                    Return ent

                End If

            Catch ex As Exception
                ' TabName exceptions gracefully
                doc.Editor.WriteMessage("Error: " & ex.Message)
            Finally
                tr.Dispose()
            End Try
        End Using

        ' Return nothing if any issue occurs
        Return Nothing
    End Function

    Public Shared Function DeleteEntityByHandle(hn As Handle) As Boolean
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim id As ObjectId = GetEntityIdByHandle(hn)

        Using tr As Transaction = doc.TransactionManager.StartTransaction()
            Try
                Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForWrite), Entity)
                If ent IsNot Nothing Then
                    ent.Erase()
                    tr.Commit()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                ' Manejar la excepción según sea necesario
                ed.WriteMessage($"Error al eliminar la entidad: {ex.Message}")
                tr.Abort()
                Return False
            Finally
                tr.Dispose()
            End Try
        End Using
    End Function

    <CommandMethod("EH")>
    Public Sub EraseObjectFromHandle()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Try

            ' Ask for a string representing the handle
            Dim pr As PromptResult = ed.GetString(vbLf & "Enter handle of object to erase: ")
            If (pr.Status = PromptStatus.OK) Then

                ' Convert hexadecimal string to 64-bit integer

                Dim ln As Long = Convert.ToInt64(pr.StringResult, 16)

                ' Not create a TabName from the long integer

                Dim hd As New Handle(ln)

                ' And attempt to get an ObjectId for the TabName
                Dim id As ObjectId = db.GetObjectId(False, hd, 0)

                ' Finally let's open the object and erase it

                Dim tr As Transaction = doc.TransactionManager.StartTransaction()

                Dim obj As DBObject = tr.GetObject(id, OpenMode.ForWrite)


                obj.Erase()

                tr.Commit()

                tr.Dispose()

            End If

        Catch ex As System.Exception
            ed.WriteMessage("Exception: " + ex.ToString)
        End Try
    End Sub
    Public Shared Sub EraseObjectByHandle(handle As Handle)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Using tr As Transaction = doc.TransactionManager.StartTransaction()
            Try

                ' And attempt to get an ObjectId for the TabName
                Dim id As ObjectId = db.GetObjectId(False, handle, 0)

                ' Finally let's open the object and erase it
                Dim obj As DBObject = tr.GetObject(id, OpenMode.ForWrite)

                obj.Erase()
                tr.Commit()

            Catch ex As System.Exception
                ed.WriteMessage("Exception: " + ex.ToString)
            Finally
                tr.Dispose()
            End Try
        End Using
    End Sub
    Public Shared Function GetListofhandlebyLayer(Layer As String) As List(Of String)
        ' Obtener el editor activo de AutoCAD
        Dim acDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim LIstHandles As New List(Of String)
        ' Recorrer los objetos del modelo
        Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim ms As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead) ', BlockTableRecord)
                For Each objId As ObjectId In ms
                    Dim acObj As Object = trans.GetObject(objId, OpenMode.ForRead)
                    Dim acEntToAdd As Entity
                    acEntToAdd = CType(acObj, Entity)
                    If TypeOf acObj Is Polyline Or
                            TypeOf acObj Is FeatureLine Or
                            TypeOf acObj Is Line Or
                            TypeOf acObj Is Polyline3d Or
                            TypeOf acObj Is Parcel Or
                            TypeOf acObj Is TinSurface Then
                        ' Verificar si el objeto es una entidad y cumple las condiciones
                        If acEntToAdd.Layer.Equals(Layer) Then
                            LIstHandles.Add(acEntToAdd.Handle.ToString())
                        End If
                    End If
                Next
                trans.Commit()
                Return LIstHandles
            Catch ex As Exception
                acDoc.Editor.WriteMessage("Error: " & ex.Message)
                Return Nothing
            Finally
                If Not trans.IsDisposed Then trans.Dispose()
            End Try
        End Using
    End Function
End Class
'Public Function GetEntityByHandle(hn As TabName) As Entity

'    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'    Dim AcEd As Editor = doc.Editor
'    Dim AcCurD As Database = doc.Database

'    Dim id As ObjectId = AcCurD.GetObjectId(False, hn, 0)

'    ' Finally let's open the object and erase it

'    Dim tr As Transaction = doc.TransactionManager.StartTransaction()

'    Dim Ent As Entity = tr.GetObject(id, OpenMode.ForRead)

'    tr.Commit()

'    tr.Dispose()

'    Return Ent
'End Function