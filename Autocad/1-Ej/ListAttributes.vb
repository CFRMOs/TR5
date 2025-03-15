Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Namespace MyApplication
    Public Class DumpAttributes
        <CommandMethod("LISTATT")>
        Public Sub ListAttributes()
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Dim tr As Transaction = db.TransactionManager.StartTransaction()

            ' Start the transaction
            Try
                ' Build a filter list so that only
                ' block references are selected
                Dim filList As TypedValue() = New TypedValue(0) {New TypedValue(CInt(DxfCode.Start), "INSERT")}
                Dim filter As New SelectionFilter(filList)
                Dim opts As New PromptSelectionOptions With {
                    .MessageForAdding = "Select block references: "
                }
                Dim res As PromptSelectionResult = ed.GetSelection(opts, filter)

                ' Do nothing if selection is unsuccessful
                If res.Status <> PromptStatus.OK Then
                    Return
                End If

                Dim selSet As SelectionSet = res.Value
                Dim idArray As ObjectId() = selSet.GetObjectIds()

                For Each blkId As ObjectId In idArray
                    Dim blkRef As BlockReference = CType(tr.GetObject(blkId, OpenMode.ForRead), BlockReference)
                    Dim btr As BlockTableRecord = CType(tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    ed.WriteMessage(ControlChars.Lf & "Block: " & btr.Name)
                    btr.Dispose()

                    Dim attCol As AttributeCollection = blkRef.AttributeCollection

                    For Each attId As ObjectId In attCol
                        Dim attRef As AttributeReference = CType(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                        Dim str As String = (ControlChars.Lf & "  Attribute Tag: " & attRef.Tag & ControlChars.Lf & "    Attribute String: " & attRef.TextString)
                        ed.WriteMessage(str)
                    Next
                Next

                tr.Commit()
            Catch ex As Exception
                tr.Abort()
                ed.WriteMessage(("Exception: " & ex.Message))
            Finally
                tr.Dispose()
            End Try
        End Sub
    End Class
End Namespace
