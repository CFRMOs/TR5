Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Public Module CMDDrawingRetrieval

    <CommandMethod("ListOpenDrawings")>
    Public Sub ListOpenDrawings()
        ' Get the document manager
        Dim docMgr As DocumentCollection = Application.DocumentManager
        Dim acDoc As Document = docMgr.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        ' Get the list of open documents
        Dim openDocs As DocumentCollection = docMgr

        ' Check if there are any open documents
        If openDocs.Count > 0 Then
            ' Display the list of open document names
            For Each doc As Document In openDocs
                Dim Pth As String = Left(doc.Name, InStrRev(doc.Name, "\"))

                For Each file As String In My.Computer.FileSystem.GetFiles(Pth)
                    acEd.WriteMessage(vbCrLf & "Open drawing: " & Replace(file, Pth, ""))
                Next
                acEd.WriteMessage(vbCrLf & "Open drawing: " & doc.Name)
            Next
        Else
        End If
    End Sub

End Module
