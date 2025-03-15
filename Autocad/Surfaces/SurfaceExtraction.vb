Imports System.Linq
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.SurfaceExtractionSettingsType
Imports C3DAligment = Autodesk.Civil.DatabaseServices.Alignment
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
Imports Polyline3d = Autodesk.AutoCAD.DatabaseServices.Polyline3d
Imports TinSurface = Autodesk.Civil.DatabaseServices.TinSurface
Public Class GPOCommands
    <CommandMethod("GSu", CommandFlags.UsePickSet)>
    Public Sub GatherSurfaceData()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        Dim promptALingOptions As New PromptEntityOptions("Seleccionar un Alignment:")
        Dim c3d_Alignment As C3DAligment = CType(SelectEntityByType("Alignment", promptALingOptions), C3DAligment)
        If c3d_Alignment Is Nothing Or TypeName(c3d_Alignment) <> "Alignment" Then Exit Sub

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                Dim Result As New List(Of Object)
                For Each objId As ObjectId In acBlkTblRec
                    Dim acObj As Object = acTrans.GetObject(objId, OpenMode.ForWrite)
                    If TypeOf acObj Is TinSurface Then
                        Dim surface As TinSurface = CType(acObj, TinSurface)
                        Dim PolyCol As DBObjectCollection = EntExtractBorder(objId)
                        'con los vertices recolectados calcular el rango de estacionamiento de la superficie respecto a un alignment
                        Dim MXstation As Double
                        Dim MNstation As Double

                        acEd.WriteMessage(vbCrLf & "Area: " & surface.Name & " TabName" & surface.Handle.ToString())

                        For Each PolyL As Polyline3d In PolyCol
                            GetDataSurface(PolyL, c3d_Alignment, MXstation, MNstation)
                            Dim PL As Polyline = CPoLy3dToPL(PolyL)
                            acEd.WriteMessage(vbCrLf & "Area: " & PL.Area.ToString("0.00"))
                            acEd.WriteMessage(vbCrLf & "Estacion inicial: " & MXstation.ToString("0+000.00"))
                            acEd.WriteMessage(vbCrLf & "Estacion final:  " & MNstation.ToString("0+000.00"))
                        Next
                    End If
                Next
                acTrans.Commit()
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage(("Exception: " & ex.Message))
            Finally
                acTrans.Dispose()
            End Try
        End Using

    End Sub
    <CommandMethod("GPO", CommandFlags.UsePickSet)>
    Public Sub GatherPolylines()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim acEd As Editor = doc.Editor

        ' Solicitar al usuario que seleccione las entidades
        Dim pso As New PromptSelectionOptions With {
            .MessageForAdding = vbLf & "Select Surface to enclose: ",
            .AllowDuplicates = False,
            .AllowSubSelections = True,
            .RejectObjectsFromNonCurrentSpace = True,
            .RejectObjectsOnLockedLayers = False
        }

        Dim psr As PromptSelectionResult = acEd.GetSelection(pso)
        If psr.Status <> PromptStatus.OK Then
            Return
        End If
        Dim promptALingOptions As New PromptEntityOptions("Seleccionar un Alignment:")
        Dim c3d_Alignment As C3DAligment = CType(SelectEntityByType("Alignment", promptALingOptions), Alignment)
        If c3d_Alignment Is Nothing Or TypeName(c3d_Alignment) <> "Alignment" Then Exit Sub

        ' Recolectar vértices de las polilíneas seleccionadas
        For Each id As ObjectId In psr.Value.GetObjectIds()
            Dim PolyCol As DBObjectCollection = EntExtractBorder(id)
            'con los vertices recolectados calcular el rango de estacionamiento de la superficie respecto a un alignment
            For Each PolyL As Polyline3d In PolyCol
                Dim MXstation As Double
                Dim MNstation As Double
                GetDataSurface(PolyL, c3d_Alignment, MXstation, MNstation)
            Next
        Next
    End Sub
    Public Sub GetDataSurface(PolyL As Polyline3d, c3d_Alignment As C3DAligment, ByRef MXstation As Double, ByRef Mnstation As Double)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim acEd As Editor = doc.Editor

        Dim iResult As New List(Of Object)

        Dim RefPtStation As Double
        Dim PGL_Elev As Double
        Dim StatOffSet As Double
        Dim mStation_OffSet_Side As String = vbNullString

        Dim acPts3d As Point3dCollection = CollectPL3DPoints(PolyL)

        Dim StationList As New List(Of Double)

        For Each Pt3d As Point3d In acPts3d
            CStationOffsetLabel.GETStationByPoint(c3d_Alignment, Pt3d, RefPtStation, PGL_Elev, StatOffSet, mStation_OffSet_Side)
            StationList.Add(RefPtStation)
        Next

        'Calculo de Maximos y minimos de los estacionamientos definidos por los puntos de los bordes y el alineamiento 
        MXstation = StationList.Max()
        Mnstation = StationList.Min()

    End Sub
    Function EntExtractBorder(id As ObjectId) As DBObjectCollection
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim AcEd As Editor = doc.Editor
        Dim acTrans As Transaction = db.TransactionManager.StartTransaction()
        Dim PolylineCol As New DBObjectCollection()
        Using acTrans
            Try
                'For Each id As ObjectId In psr.Value.GetObjectIds()
                Dim ent As Entity = TryCast(acTrans.GetObject(id, OpenMode.ForRead), Entity)
                Dim Sur As TinSurface = CType(ent, TinSurface)
                Dim ExBorder As ObjectIdCollection = Sur.ExtractBorder(Plan)
                For Each idBorder As ObjectId In ExBorder
                    Dim PL3DBorder As Polyline3d = CType(acTrans.GetObject(idBorder, OpenMode.ForRead), Polyline3d)
                    PolylineCol.Add(PL3DBorder)
                Next
                acTrans.Commit()
                Return PolylineCol
            Catch ex As Exception
                acTrans.Abort()
                AcEd.WriteMessage("Error :" & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
        Return Nothing
    End Function
End Class

