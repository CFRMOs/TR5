Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
'Imports ExcelInteropManager
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class COverlapCheck

    ' Define el comando para verificar solapamiento
    <CommandMethod("CheckPolylinesOverlap")>
    Public Sub CheckPolylinesOverlap()
        ' Obtener el documento y el editor de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Seleccionar la primera entidad para determinar la capa
        Dim entity1 As Entity = Nothing
        Dim layer1Name As String = CSelectionHelper.GetLayerByEnt(entity1)
        If String.IsNullOrEmpty(layer1Name) Then Exit Sub

        ' Seleccionar la segunda entidad para determinar la capa
        Dim entity2 As Entity = Nothing
        Dim layer2Name As String = CSelectionHelper.GetLayerByEnt(entity2)
        If String.IsNullOrEmpty(layer2Name) Then Exit Sub

        ' Seleccionar el alineamiento
        Dim alignment As Alignment = CSelectionHelper.SelectAlignment(ed, "Seleccione el alineamiento")
        If alignment Is Nothing Then Exit Sub

        ' Obtener las entidades que se solapan
        Dim resultList As List(Of (Entity, List(Of Entity))) = ListPLOverLaping(layer1Name, layer2Name, alignment)

        ' Bucle imprimiendo los resultados
        For Each result In resultList
            Dim AcEnt1 As Entity = result.Item1
            Dim overlappingEntities As List(Of Entity) = result.Item2

            ed.WriteMessage(vbLf & $"Entidad en {layer1Name} con handle {AcEnt1.Handle} se solapa con:")

            For Each AcEnt2 In overlappingEntities
                ed.WriteMessage(vbLf & $" - Entidad en {layer2Name} con handle {AcEnt2.Handle}")
            Next
        Next
    End Sub

    Public Shared Sub Overlap(DGView1 As System.Windows.Forms.DataGridView, DGView2 As System.Windows.Forms.DataGridView,
                              ByRef CAcadHelp As ACAdHelpers, columnData As List(Of (String, String, String)), columnFormats As List(Of String), ByRef VBListRelateTramos As ListadoRelacionados)
        CollectAndExportOverlappingInfo(DGView1, DGView2, CAcadHelp, columnData, columnFormats)
        DGView2.FindForm().Show()
        DGView1.FindForm().Show()
        VBListRelateTramos.ThisDrawing = CAcadHelp.ThisDrawing
    End Sub

    ' Collect and export overlapping information to Excel
    Public Shared Sub CollectAndExportOverlappingInfo(ByRef DGView1 As DataGridView, ByRef DGView2 As DataGridView, ByRef CAcadHelp As ACAdHelpers, columnData As List(Of (String, String, String)), columnFormats As List(Of String))
        Dim entity1 As Entity = Nothing
        Dim layer1Name As String = CSelectionHelper.GetLayerByEnt(entity1)
        If String.IsNullOrEmpty(layer1Name) Then Exit Sub

        Dim entity2 As Entity = Nothing
        Dim layer2Name As String = CSelectionHelper.GetLayerByEnt(entity2)
        If String.IsNullOrEmpty(layer2Name) Then Exit Sub

        CollectInfoOverlapping(DGView1, DGView2, CAcadHelp.ThisDrawing, CAcadHelp.Alignment, layer1Name, layer2Name, columnData, columnFormats)

        TranferSlapingInf.ExceTR(DGView1, columnData, columnFormats)

    End Sub




    ' Collect information of overlapping entities
    Public Shared Sub CollectInfoOverlapping(ByRef DGView1 As DataGridView, ByRef DGView2 As DataGridView,
                                             ByRef ThisDrawing As Document, ByRef c3d_Alignment As Alignment, layer1Name As String, layer2Name As String,
                                             columnData As List(Of (String, String, String)), columnFormats As List(Of String))
        If c3d_Alignment Is Nothing Then Exit Sub

        Dim resultList As List(Of (Entity, List(Of Entity))) = ListPLOverLaping(layer1Name, layer2Name, c3d_Alignment)

        DGView1.Rows.Clear()
        DGView2.Rows.Clear()

        DataGridViewHelper.Addcolumns(DGView1, columnData)
        DataGridViewHelper.Addcolumns(DGView2, columnData)

        For Each ListOverlap In resultList
            Dim mainEntity As Entity = ListOverlap.Item1
            Dim mainRowIndex As Integer = DataGridViewHelper.AddEntityInfo(ThisDrawing, c3d_Alignment, mainEntity, DGView1)

            For Each overlappingEntity In ListOverlap.Item2
                Dim secondaryRowIndex As Integer = DataGridViewHelper.AddEntityInfo(ThisDrawing, c3d_Alignment, overlappingEntity, DGView2)
                UpdateTramos(DGView2, secondaryRowIndex, mainEntity)
                UpdateTramos(DGView1, mainRowIndex, overlappingEntity)
            Next
        Next
    End Sub

    ' Update "Tramos" column
    Public Shared Sub UpdateTramos(DGView As DataGridView, rowIndex As Integer, AcEnt As Entity)
            Dim row As DataGridViewRow = DGView.Rows(rowIndex)
            Dim cell As DataGridViewCell = row.Cells("Tramos")
            Dim tramos As String = If(cell.Value, "")

            If InStr(tramos, AcEnt.Handle.ToString(), CompareMethod.Text) = 0 Then
                cell.Value = If(tramos = "", AcEnt.Handle.ToString(), tramos & "," & AcEnt.Handle.ToString())
            End If
        End Sub
        ' Función para listar entidades que se solapan
        Public Shared Function ListPLOverLaping(layer1 As String, layer2 As String, align As Alignment) As List(Of (Entity, List(Of Entity)))
            Dim resultList As New List(Of (Entity, List(Of Entity)))()

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

            Using AcTrans As Transaction = db.TransactionManager.StartTransaction()
                ' Obtener todas las entidades de layer1
                Dim bt As BlockTable = CType(AcTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = CType(AcTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                For Each objId As ObjectId In btr
                    Dim AcEnt1 As Entity = CType(AcTrans.GetObject(objId, OpenMode.ForRead), Entity)
                    If AcEnt1.Layer = layer1 AndAlso CStationOffsetLabel.IsSupportedEntity(AcEnt1) Then
                        Dim overlappingEntities As New List(Of Entity)()

                        ' Obtener información de la entidad
                        Dim startStation1, endStation1, len1, area1 As Double
                        Dim side1Start, side1End As String
                        Dim startPt1, endPt1 As Point3d
                        side1Start = String.Empty
                        side1End = String.Empty

                        CStationOffsetLabel.ProcessEntity(AcEnt1, startPt1, endPt1, startStation1, endStation1, side1Start, side1End, len1, area1, align)

                        ' Comparar con todas las entidades de layer2
                        For Each objId2 As ObjectId In btr
                            Dim AcEnt2 As Entity = CType(AcTrans.GetObject(objId2, OpenMode.ForRead), Entity)
                            If AcEnt2.Layer = layer2 AndAlso CStationOffsetLabel.IsSupportedEntity(AcEnt2) Then

                                ' Obtener información de la entidad
                                Dim startStation2, endStation2, len2, area2 As Double
                                Dim side2Start, side2End As String
                                Dim startPt2, endPt2 As Point3d
                                side2Start = String.Empty
                                side2End = String.Empty

                                CStationOffsetLabel.ProcessEntity(AcEnt2, startPt2, endPt2, startStation2, endStation2, side2Start, side2End, len2, area2, align)

                                ' Comprobar si se solapan
                                If AreStationsOverlapping(startStation1, endStation1, startStation2, endStation2, side1Start, side2Start) Then
                                    overlappingEntities.Add(AcEnt2)
                                End If
                            End If
                        Next

                        ' Añadir al resultado si hay solapamientos
                        If overlappingEntities.Count > 0 Then
                            resultList.Add((AcEnt1, overlappingEntities))
                        End If
                    End If
                Next
                AcTrans.Commit()
            End Using

            Return resultList
        End Function



        ' Función para comprobar si dos rangos de estaciones se solapan y están en el mismo lado
        Public Shared Function AreStationsOverlapping(start1 As Double, end1 As Double, start2 As Double, end2 As Double, side1 As String, side2 As String) As Boolean
            ' Las estaciones se solapan si están en el mismo lado y los rangos se cruzan
            Return side1 = side2 AndAlso start1 <= end2 AndAlso start2 <= end1
        End Function

        ' Función para obtener las estaciones y el lado de una polilínea usando STSL.GETStationByPoint
        Public Shared Sub GetStationByPoint(c3d_Alignment As Alignment, startPt As Point3d, endPt As Point3d, ByRef startStation As Double, ByRef endStation As Double, ByRef startSide As String, ByRef endSide As String)

            Dim startElev As Double
            Dim startOffset As Double
            CStationOffsetLabel.GETStationByPoint(c3d_Alignment, startPt, startStation, startElev, startOffset, startSide)

            Dim endElev As Double
            Dim endOffset As Double
            CStationOffsetLabel.GETStationByPoint(c3d_Alignment, endPt, endStation, endElev, endOffset, endSide)

            CStationOffsetLabel.GetMxMnOPPL(startStation, endStation)
        End Sub



    End Class

'' Collect overlapping information and export to Excel
'Private Sub IDSolapes_Click(sender As Object, e As EventArgs) Handles IDSolapes.Click

'    Dim DGView1 As DataGridView = Me.CunetasExistentesDGView
'    Dim DGView2 As DataGridView = VBListRelateTramos.DataGridView1
'    COverlapCheck.Overlap(DGView1, DGView2, CAcadHelp, columnData, columnFormats, VBListRelateTramos)
'End Sub
'' Button click event to add selected entity information 
'Private Sub RTramos_Click(sender As Object, e As EventArgs) Handles RTramos.Click
'    SetLData()
'    VBListRelateTramos.ThisDrawing = CAcadHelp.ThisDrawing
'    Dim DGView As DataGridView = CunetasExistentesDGView
'    Dim DGView1 As DataGridView = VBListRelateTramos.DataGridView1
'    Dim selected As DataGridViewCell = DGView.SelectedCells(0)
'    Dim Rows As DataGridViewRow = DGView.Rows(selected.RowIndex)
'    Dim cell As DataGridViewCell = Rows.Cells("Tramos")
'    Dim tramos As String = If(cell.Value IsNot Nothing, cell.Value.ToString(), "")
'    If InStr(tramos, hdString, CompareMethod.Text) = 0 Then
'        cell.Value = If(tramos = "", hdString, tramos & "," & hdString)
'    End If
'    DataGridViewHelper.Addcolumns(DGView1, columnData)
'    DGView1.Show()
'    DataGridViewHelper.StrAddEntityInfo(CAcadHelp.ThisDrawing, CAcadHelp.Alignment, hdString, VBListRelateTramos.DataGridView1)
'End Sub