Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Public Module CMDTestEnt
    <CommandMethod("CMDTestEnt")> Public Sub TestEnt()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim acEnt As Entity

        Dim promptALingOptions As New PromptEntityOptions("Seleccionar una linea")
        Dim PromptResult As PromptEntityResult = Nothing

        acEnt = SelectEntity(promptALingOptions, PromptResult)
        If PromptResult.Status <> PromptStatus.OK Then Exit Sub
    End Sub

End Module
