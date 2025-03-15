Imports System.Windows.Forms
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
'Imports ExcelInteropManager2
Imports Alignment = Autodesk.Civil.DatabaseServices.Alignment
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
Imports StationOffsetLabel = Autodesk.Civil.DatabaseServices.StationOffsetLabel
Public Module GetlabelStation
    <CommandMethod("CMDGetlabelStation")>
    Public Sub CMDGetlabelStation()

        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCuDd As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor
        Dim c3d_Alignment As Alignment
        Dim SelMG As New CSelectionHelper
        Dim promptALingOptions As New PromptEntityOptions("Seleccionar un Alignment:")
        Dim acEntAlignment As Entity
        Dim promptLabelOptions As New PromptEntityOptions("Seleccionar un Label:")
        Dim acEntLabel As Entity
        acEntAlignment = SelectEntityByType("Alignment", promptALingOptions)

        acEntLabel = SelectEntityByType("StationOffsetLabel", promptLabelOptions)

        If acEntAlignment Is Nothing Or acEntLabel Is Nothing Then Exit Sub

        Using trans As Transaction = AcCuDd.TransactionManager.StartTransaction()

            Try
                'Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                'Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                Dim Slabel As StationOffsetLabel = CType(trans.GetObject(acEntLabel.Id, OpenMode.ForRead), StationOffsetLabel)
                'get the anchor point 
                Dim PT As Point3d = Slabel.LabelLocation

                Dim STStation_OffSet_Side As String = vbNullString, StPGL_Elev As Double

                Dim PtStation As Double, StStatOffSet As Double

                c3d_Alignment = CType(trans.GetObject(acEntAlignment.Id, OpenMode.ForRead), Alignment)

                CStationOffsetLabel.GETStationByPoint(c3d_Alignment, PT, PtStation, StPGL_Elev, StStatOffSet, STStation_OffSet_Side)

                Dim valorFormateado As String = PtStation.ToString("0+000.00")

                Clipboard.SetText(valorFormateado)

                AcEd.WriteMessage(vbCrLf & valorFormateado)

                trans.Commit()
            Catch ex As Exception
                trans.Abort()
                AcEd.WriteMessage(vbCrLf & ex.Message)
            Finally
                trans.Dispose()
            End Try
        End Using

    End Sub
    <CommandMethod("CMDPickStation")>
    Public Sub CMDPickStation()

        Dim AcDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim AcCuDd As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor
        Dim promptBlockOptions As New PromptEntityOptions("Seleccionar un BlockReference:")
        Dim promptALingOptions As New PromptEntityOptions("Seleccionar un Alignment:")
        Dim xl As New ExcelAppTR
        Dim HostAppS As HostApplicationServices = HostApplicationServices.Current

        Dim FPath As String = HostAppS.FindFile(AcDoc.Name, AcCuDd, FindFileHint.Default)
        'AcDoc.

        Dim c3d_Alignment As Alignment = CType(SelectEntityByType("Alignment", promptALingOptions), Alignment)
        If c3d_Alignment Is Nothing Or TypeName(c3d_Alignment) <> "Alignment" Then Exit Sub

        Do
            Using trans As Transaction = AcCuDd.TransactionManager.StartTransaction()
                Dim PromptEntityR As PromptEntityResult = Nothing

                Dim BloRef As BlockReference = CType(SelectEntityByType("BlockReference", promptBlockOptions, PromptEntityR), BlockReference)

                If BloRef Is Nothing Or TypeName(BloRef) <> "BlockReference" Then Exit Sub

                If PromptEntityR.Status = PromptStatus.OK Then
                    Try
                        Dim PT As Point3d = GetPoint()

                        Dim STStation_OffSet_Side As String = vbNullString
                        Dim StPGL_Elev As Double
                        Dim PtStation As Double
                        Dim StStatOffSet As Double

                        CStationOffsetLabel.GETStationByPoint(c3d_Alignment, PT, PtStation, StPGL_Elev, StStatOffSet, STStation_OffSet_Side)

                        'Dim valorFormateado As String = PtStation.ToString("0+000.00")

                        'getting access to acadobject(entity)'s property  
                        Dim btr As BlockTableRecord = CType(trans.GetObject(BloRef.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                        Dim AR As Array = {BloRef.Handle.ToString(), PtStation, STStation_OffSet_Side, btr.Name, c3d_Alignment.Handle.ToString()}
                        Dim ARTitulos As Array = {"TabName", "Estacion", "Lado", "Tipo", "HandleALRF"}
                        Dim ARFT As Array = {"@", "0+000.00", "@", "@", "@"}

                        'xl.SetTransferData(AR, ARTitulos, ARFT)

                        trans.Commit()
                    Catch ex As Exception
                        trans.Abort()
                        AcEd.WriteMessage(vbCrLf & ex.Message)
                    Finally
                        trans.Dispose()
                    End Try
                Else
                    AcEd.WriteMessage(vbLf & "Error or user cancelled")
                End If
            End Using
        Loop
    End Sub

    '' Process polyline entities
    'Public Sub ProcessPolyline(Alignment As Alignment, ByVal PL As Polyline, ByRef StartPT As Point3d, ByRef EndPT As Point3d, ByRef Len As Double, ByRef Area As Double, ByRef StPtStation As Double, ByRef EDPtStation As Double, ByRef STStation_OffSet_Side As String, ByRef EDStation_OffSet_Side As String)
    '    If PL.Closed Then
    '        GetMxMnBorder(PL, Alignment, StPtStation, EDPtStation, STStation_OffSet_Side)
    '    Else
    '        StartPT = PL.StartPoint
    '        EndPT = PL.EndPoint
    '        GetexStation(Alignment, StPtStation, EDPtStation, StartPT, EndPT, STStation_OffSet_Side, EDStation_OffSet_Side)
    '    End If
    '    Len = PL.Length
    '    Area = PL.Area
    'End Sub

End Module
