Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
Public Module AlignmentExtractor
    Public Function GetAllAlignments(ByRef List_ALingName As List(Of String)) As List(Of Alignment)
        Dim alignments As New List(Of Alignment)
        Dim ALingNames As New List(Of String)
        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        ' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Get the block table and the model space block table record
                Dim bt As BlockTable = CType(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim ms As BlockTableRecord = CType(acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                ' Iterate through each entity in the model space
                For Each objId As ObjectId In ms
                    Dim entity As Entity = TryCast(acTrans.GetObject(objId, OpenMode.ForRead), Entity)
                    If entity IsNot Nothing AndAlso TypeOf entity Is Alignment Then
                        Dim alignment As Alignment = CType(entity, Alignment)
                        alignments.Add(alignment)
                        ALingNames.Add(alignment.Name)
                    End If
                Next
                ' Commit the transaction
                List_ALingName = ALingNames
                acTrans.Commit()
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage(("Exception: " & ex.Message))
            Finally
                acTrans.Dispose()
            End Try
        End Using
        Return alignments

    End Function
    Public Function GetCLPLAtSt() As List(Of Polyline)
        ' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim polylines As New List(Of Polyline) ' Initialize the return list

        ' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Get the block table and the model space block table record
                Dim bt As BlockTable = CType(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim ms As BlockTableRecord = CType(acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                ' Iterate through each entity in the model space
                For Each objId As ObjectId In ms
                    ' Add your logic here to process each entity and add to the polylines list if needed
                Next
                ' Commit the transaction
                acTrans.Commit()
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage("Exception: " & ex.Message)
            Finally
                If Not acTrans.IsDisposed Then
                    acTrans.Dispose() ' Dispose transaction if not already disposed
                End If
            End Try
        End Using

        Return polylines ' Ensure a return value is always provided
    End Function

End Module
