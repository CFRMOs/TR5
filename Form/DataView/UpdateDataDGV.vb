Imports System.Linq
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.Geometry
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity

Public Class UpdateDataDGV
    'ultima version de actualizacion DGViewes con el uso de de la clase 
    Public Shared Sub UpDateDGVByHandleDataProcessor(DicHandles() As (Dictionary(Of String, CunetasHandleDataItem), DataGridView), ByRef CAcadHelp As ACAdHelpers, ByRef handleProcessor As HandleDataProcessor)
        'Dim DGView1 As DataGridView = Me.CunetasExistentesDGView
        'Dim DGView2 As DataGridView = Me.DGViewByLeght
        If handleProcessor.HandleData Is Nothing OrElse handleProcessor.HandleData.Count = 0 Then
            ' Si no hay datos, salir del sub
            Exit Sub
        End If
        'mostral datos en los DataGridView
        ' Bucle único para procesar ambos diccionarios
        For Each handleGroup As (Dictionary(Of String, CunetasHandleDataItem), DataGridView) In DicHandles
            Dim handleDic As Dictionary(Of String, CunetasHandleDataItem) = handleGroup.Item1
            Dim DGView As DataGridView = handleGroup.Item2

            DGView.Rows.Clear()

            For Each kvp As KeyValuePair(Of String, CunetasHandleDataItem) In handleDic
                Dim HandleDataItem As CunetasHandleDataItem = handleDic(kvp.Key)
                IdentifyPL.SetALingment(CAcadHelp, HandleDataItem.AlignmentHDI)
                HandleDataItem.AddToDGView(DGView, -1)
            Next
        Next
    End Sub

    Public Shared Function IsLYinGroup(selectedLayer As String, nombresCapas As String()) As Boolean
        Dim Adb As Boolean = True
        For i As Integer = 0 To nombresCapas.Length - 1
            If UCase(selectedLayer) = UCase(nombresCapas(i)) Then
                Adb = False
                Exit For
            End If
        Next
        Return Adb
    End Function
    Public Shared Sub SETLYGroup(selectedLayer As String, ByRef nombresCapas As String(), ByRef DCapas As Dictionary(Of String, String()), groupName As String)
        Dim Adb As Boolean = UpdateDataDGV.IsLYinGroup(selectedLayer, nombresCapas)

        If Adb AndAlso Not nombresCapas.Contains(selectedLayer) Then
            ReDim Preserve nombresCapas(UBound(nombresCapas) + 1)
            nombresCapas(UBound(nombresCapas)) = selectedLayer

            CrearGrupoYCapas(groupName, nombresCapas)
            DCapas(groupName) = nombresCapas
        End If
    End Sub

    ' Update DGView based on selected layer
    Public Shared Sub UpdateDataGridViewForSelectedLayer(DataGridView As DataGridView, Alignment As Alignment, selectedLayer As String, ThisDrawing As Document)
        Dim acEnt As Entity = Nothing
        If Alignment Is Nothing Then Exit Sub
        If DataGridView.Rows(0).Cells("TabName").Value IsNot Nothing Then
            Dim hdString As String = DataGridView.Rows(0).Cells("TabName").Value.ToString()
            acEnt = CLHandle.GetEntityByStrHandle(hdString)
        End If
        If acEnt IsNot Nothing Then
            Dim type As String = acEnt.GetType().ToString()
        End If
        DataGridView.Rows.Clear()
        DataGridView.FindForm.Update()
        'CAcadHelp.
        DateViewSet.SetViewCellEnt(ThisDrawing, Alignment, DataGridView.FindForm, selectedLayer, DataGridView)
        DataGridView.FindForm.Show()
        DataGridView.FindForm.Activate()
    End Sub
    Public Shared Sub CUPDateDGView(DGView As DataGridView, ByRef CAcadHelp As ACAdHelpers)
        UPDateDGView(DGView, CAcadHelp.ThisDrawing, CAcadHelp.Alignment)
    End Sub
    Public Shared Sub UPDateDGView(DGView As DataGridView, ThisDrawing As Document, Alignment As Alignment)
        Dim acEnt As Entity
        DGView.FindForm.Update()
        For Each rows As DataGridViewRow In DGView.Rows
            If rows.Cells(0).Value IsNot Nothing Then
                Dim hdString As String = rows.Cells(1).Value
                If hdString <> vbNullString Then
                    acEnt = CLHandle.GetEntityByStrHandle(hdString)
                    If acEnt IsNot Nothing Then
                        Dim selectedLayer As String = acEnt.Layer
                        UpdateDataDGV.UpdateSelectedEntityInfo(ThisDrawing, acEnt, rows, Alignment)
                    Else
                        DGView.Rows.Remove(rows)
                    End If
                End If
            End If
        Next
        DGView.FindForm.Show()
        DGView.FindForm.Activate()
    End Sub

    ' Update selected entity info
    Public Shared Function StrUpdateSelectedEntityInfo(ByVal ThisDrawing As Document, ByVal acEntHandle As String, selectedRow As DataGridViewRow, Alignment As Alignment) As Integer
        Return UpdateSelectedEntityInfo(ThisDrawing, CLHandle.GetEntityByStrHandle(acEntHandle), selectedRow, Alignment)
    End Function
    Public Shared Function UpdateSelectedEntityInfo(ByVal ThisDrawing As Document, ByVal acEnt As Entity, selectedRow As DataGridViewRow, Alignment As Alignment) As Integer
        Try
            selectedRow.Cells("Handle").Value = acEnt.Handle.ToString()
            Dim rowIndex As Integer = selectedRow.Index
            Dim StartPT, EndPT As Point3d
            Dim Len, Area, StPtStation, EDPtStation As Double
            Dim Side1 As String = String.Empty
            Dim Side2 As String = String.Empty
            CStationOffsetLabel.ProcessEntity(acEnt, StartPT, EndPT, StPtStation, EDPtStation, Side1, Side2, Len, Area, Alignment)

            PopulateRowWithEntityData(ThisDrawing, selectedRow, acEnt, StartPT, EndPT, Len, Area, StPtStation, EDPtStation, Side1, Alignment)

            Return rowIndex

        Catch ex As Exception
            MessageBox.Show($"Error updating entity info: {ex.Message}")
            Return -1
        End Try
    End Function

    ' Populate DGView row with entity data
    Public Shared Sub PopulateRowWithEntityData(ThisDrawing As Document, selectedRow As DataGridViewRow, acEnt As Entity, StartPT As Point3d, EndPT As Point3d, Len As Double, Area As Double, StPtStation As Double, EDPtStation As Double, Side As String, Alignment As Alignment)

        selectedRow.Cells("Layer").Value = acEnt.Layer
        selectedRow.Cells("MinX").Value = StartPT.X
        selectedRow.Cells("MinY").Value = StartPT.Y
        selectedRow.Cells("MaxX").Value = EndPT.X
        selectedRow.Cells("MaxY").Value = EndPT.Y
        selectedRow.Cells("StartStation").Value = StPtStation
        selectedRow.Cells("EndStation").Value = EDPtStation
        selectedRow.Cells("Longitud").Value = Len
        selectedRow.Cells("Area").Value = Area
        selectedRow.Cells("Side").Value = Side
        selectedRow.Cells("AlignmentHDI").Value = Alignment.Handle.ToString()

        Dim PlanoName As String = PlanoNombre(acEnt.Id)
        selectedRow.Cells("PLANO").Value = PlanoName


        ' Obtener el nombre del archivo desde la ruta completa
        Dim fileName As String = System.IO.Path.GetFileName(ThisDrawing.Name)
        selectedRow.Cells("FileName").Value = fileName
        selectedRow.Cells("FilePath").Value = ThisDrawing.Name

        Dim comentario As String = TryCast(GetDATA(acEnt, "Comentarios"), String)
        selectedRow.Cells("Comentarios").Value = comentario
        selectedRow.Cells("IDNum").Value = GetDATA(acEnt, "IDNum")
    End Sub

    'obtener loa comentarios de la columna de comentarios 
    Public Shared Function GetDATA(acEnt As Entity, Header As String)
        ''Obtener comentarios de Excel
        Dim ExRNGSe As New ExcelRangeSelector()

        Dim Headers As New List(Of String) From {Header}

        'Dim Data As List(Of Object) = ExRNGSe.SelectRowOnTbl("CunetasGeneral", Headers, acEnt.Handle.ToString())
        Dim comments As String = String.Empty
        Try
            comments = TryCast(ExRNGSe.HandleFromExcelTable(TBName:="CunetasGeneral", handle:=acEnt.Handle.ToString(), Header).Value, Object)
            Return comments

        Catch ex As Exception
            Return vbNullString
        End Try
    End Function
End Class
