Imports Autodesk.AutoCAD.EditorInput
Public Class GetLabelTextFromSLEnt
    Public Shared Sub GetLabelText()
        ' Obtenemos el documento activo y el editor
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim label As GeneralSegmentLabel = AlignmentLabelHelper.GetLabelForLineOrCurve(CSelectionHelper.SelectEntityObjectid())
        Dim textComp As String = GeneralSegmentLabelHelper.GetLabelText(label)
        Dim textExplo As String = GeneralSegmentLabelHelper.GetExplodedLabelText(label)
        Dim NumCuneta As String = CodigoExtractor.ExtraerCodigo(textExplo, "CU\d+-\d+")
    End Sub
End Class
