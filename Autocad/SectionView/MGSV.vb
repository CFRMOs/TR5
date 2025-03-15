Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports C3DAligment = Autodesk.Civil.DatabaseServices.Alignment
Imports C3DSectionView = Autodesk.Civil.DatabaseServices.SectionView
Imports DBObject = Autodesk.AutoCAD.DatabaseServices.DBObject
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
Imports Section = Autodesk.Civil.DatabaseServices.Section
Public Module CmdSectionView
    <CommandMethod("CmdSectionView")>
    Public Sub CmdSectionView()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        Do
            Dim promptPointOptions As New PromptPointOptions("Seleccionar un punto en el dibujo:")
            Dim promptPointResult As PromptPointResult = acEd.GetPoint(promptPointOptions)
            Dim point As Point3d = promptPointResult.Value

            If promptPointResult.Status = PromptStatus.OK Then

                Dim SV As SectionView = GetSVieBywPoint(point)
                AddPointToAling(SV, New Point3d(point.X, point.Y, 0))

            ElseIf promptPointResult.Status = PromptStatus.Cancel Then
                ' Cancelar la transacción y devolver Nothing si se cancela la selección
                Exit Do
            End If
        Loop

    End Sub
    <CommandMethod("Cmdsv")>
    Public Sub Cmdsv()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                Dim Result As New List(Of Object)
                Dim objs As New DBObjectCollection()

                Dim promptPLOptions As New PromptEntityOptions("Seleccionar un SectionView:")
                Dim PromptPLResult As PromptEntityResult = Nothing
                Dim ent As Entity = SelectEntityByType("SectionView", promptPLOptions, PromptPLResult)
                ent.Explode(objs)
                For Each Obj As DBObject In objs
                    If TypeName(Obj) = "BlockReference" Then
                        'ed.WriteMessage(vbCrLf & TypeName(Obj))
                        Dim objs2 As New DBObjectCollection()
                        ent.Explode(objs2)
                        For Each Obj2 As DBObject In objs2
                            If TypeName(Obj) = "Line" Then
                                acTrans.Commit()

                            End If
                        Next

                    End If
                Next
            Catch ex As Exception
                acEd.WriteMessage(("Exception: " & ex.Message))
                acTrans.Abort()
                'Return Nothing
            Finally
                If Not acTrans.IsDisposed Then
                    acTrans.Dispose() ' Dispose transaction if not already disposed
                End If
            End Try
        End Using


    End Sub
    <CommandMethod("CmdSVPL")>
    Public Sub CmdSVPL()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        Do
            Dim promptPLOptions As New PromptEntityOptions("Seleccionar un Polyline:")
            Dim PromptPLResult As PromptEntityResult = Nothing
            Dim PL As Polyline = CType(SelectEntityByType("Polyline", promptPLOptions, PromptPLResult), Polyline)

            If PL Is Nothing Or TypeName(PL) <> "Polyline" Then Exit Sub

            If PromptPLResult.Status = PromptStatus.OK Then
                Dim acPts2d As Point2dCollection = CollectPLPoints(PL)

                Dim Point As New Point3d(acPts2d(0).X, acPts2d(0).Y, 0)

                Dim SV As SectionView = GetSVieBywPoint(Point)

                For Each pt As Point2d In acPts2d
                    AddPointToAling(SV, New Point3d(pt.X, pt.Y, 0))
                Next
            ElseIf PromptPLResult.Status = PromptStatus.Cancel Then
                ' Cancelar la transacción y devolver Nothing si se cancela la selección
                Exit Do
            End If
        Loop
    End Sub
    Function GetSVieBywPoint(Point As Point3d) As SectionView
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each id As ObjectId In CivilApplication.ActiveDocument.GetAlignmentIds()
                    Dim algn As C3DAligment = TryCast(acTrans.GetObject(id, OpenMode.ForRead), C3DAligment)
                    For Each SLGrId As ObjectId In algn.GetSampleLineGroupIds()
                        Dim SlGr As SampleLineGroup = TryCast(acTrans.GetObject(SLGrId, OpenMode.ForRead), SampleLineGroup)
                        For Each SLid As ObjectId In SlGr.GetSampleLineIds()
                            Dim SL As SampleLine = TryCast(acTrans.GetObject(SLid, OpenMode.ForRead), SampleLine)
                            Dim Stid As ObjectId = SL.GetSectionViewIds(0)
                            Dim SV As SectionView = TryCast(acTrans.GetObject(Stid, OpenMode.ForRead), SectionView)
                            Dim ext As Extents3d = SV.GeometricExtents()
                            Dim Mxpt As Point3d = ext.MaxPoint
                            Dim Mnpt As Point3d = ext.MinPoint
                            If Mxpt.X > Point.X And Point.X > Mnpt.X And Mxpt.Y > Point.Y And Point.Y > Mnpt.Y Then
                                acTrans.Commit()
                                Return SV
                            End If
                        Next
                    Next
                Next
            Catch ex As Exception
                acEd.WriteMessage("Exception: " & ex.Message)
                acTrans.Abort()
            Finally
                If Not acTrans.IsDisposed Then
                    acTrans.Dispose() ' Dispose transaction if not already disposed
                End If
            End Try
        End Using
        Return Nothing ' Ensure a return value is always provided
    End Function
    Public Sub AddPointToAling(c3d_SectionView As SectionView, Point As Point3d)
        Dim elevation As Double = 0, offset As Double = 0
        Dim Station As Double = 0
        Dim Eat As Double
        Dim Nor As Double

        GetDataPointbyIDSetctionView(Point, c3d_SectionView.Id, elevation, offset, Station, Eat, Nor)

        CGPointHelper.AddCGPoint(Eat, Nor, elevation, "This is section Point")
    End Sub
    Public Function GetDataPointbyIDSetctionView(point As Point3d, idSV As ObjectId, ByRef elevation As Double, ByRef offset As Double, ByRef Station As Double, ByRef Eat As Double, ByRef Nor As Double) As Boolean
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        ' Obtener la mínima elevación del SectionView
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim sv As C3DSectionView = CType(acTrans.GetObject(idSV, OpenMode.ForWrite), C3DSectionView)
                Dim Sl As SampleLine = CType(acTrans.GetObject(sv.SampleLineId, OpenMode.ForRead), SampleLine)

                sv.IsElevationRangeAutomatic = False
                Dim minimumElevation As Double = sv.ElevationMin

                ' Obtener la información del SectionView
                Dim insertionPoint As Point3d = sv.Location

                sv.FindOffsetAndElevationAtXY(point.X, point.Y, offset, elevation)

                Dim STId As ObjectId = Sl.GetSectionIds(0)

                Dim ST As Section = CType(acTrans.GetObject(STId, OpenMode.ForRead), Section)
                Station = ST.Station()

                Dim AlignId As ObjectId = ParentAlignmentId(Sl)
                Dim ALgn As Alignment = CType(acTrans.GetObject(AlignId, OpenMode.ForRead), Alignment)

                ALgn.PointLocation(Station, offset, Eat, Nor)

                acTrans.Commit()
                Return True
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage(("Exception: " & ex.Message))
                Return False
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function
    <System.Runtime.CompilerServices.Extension()>
    Public Function ParentAlignmentId(ByVal sl As SampleLine) As ObjectId
        Dim algnId As ObjectId = ObjectId.Null
        For Each id As ObjectId In CivilApplication.ActiveDocument.GetAlignmentIds()
            Dim algn As Alignment = DirectCast(id.GetObject(OpenMode.ForRead), Alignment)
            For Each slgId As ObjectId In algn.GetSampleLineGroupIds()
                If slgId = sl.GroupId Then
                    algnId = id
                    Exit For
                End If
            Next
            If algnId <> ObjectId.Null Then
                Exit For
            End If
        Next
        Return algnId
    End Function
End Module
