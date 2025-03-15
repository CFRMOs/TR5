Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Public Module CollectPLinePoints
    Public Function CollectPL3DPoints(acPoly3d As Polyline3d) As Point3dCollection
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = doc.Database
        Dim AcEd As Editor = doc.Editor

        Dim acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
        Using acTrans
            Try
                '' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                               OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                  OpenMode.ForRead)

                '' Get the coordinates of the 3D polyline
                Dim acPts3d As New Point3dCollection()
                For Each acObjIdVert As ObjectId In acPoly3d
                    Dim acPolVer3d As PolylineVertex3d
                    acPolVer3d = acTrans.GetObject(acObjIdVert,
                                     OpenMode.ForRead)

                    acPts3d.Add(acPolVer3d.Position)
                Next
                acTrans.Commit()
                Return acPts3d
            Catch ex As Exception
                acTrans.Abort()
                AcEd.WriteMessage("Error :" & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function
    Public Function CollectPLPoints(acPoly As Polyline) As Point2dCollection
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = doc.Database
        Dim AcEd As Editor = doc.Editor
        '' Get the coordinates of the 2D polyline
        Dim acPts2d As New Point2dCollection()
        'Dim Ac2dpl As Polyline2d = ConvertToPolyline2d(acPoly.Id)

        Dim acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
        Using acTrans
            Try
                '' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                               OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                doc.LockDocument()
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                  OpenMode.ForWrite)

                For i As Integer = 0 To acPoly.NumberOfVertices - 1
                    Dim pt As Point2d = acPoly.GetPoint2dAt(i)
                    acPts2d.Add(pt)
                Next

                acTrans.Commit()
                Return acPts2d
            Catch ex As Exception
                acTrans.Abort()
                AcEd.WriteMessage("Error :" & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function
    Public Function CollectPL2dPoints(acPoly As Polyline2d) As Point3dCollection
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = doc.Database
        Dim AcEd As Editor = doc.Editor
        '' Get the coordinates of the 2D polyline
        Dim acPts2d As New Point3dCollection()
        'Dim Ac2dpl As Polyline2d = ConvertToPolyline2d(acPoly.Id)

        Dim acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
        Using acTrans
            Try
                '' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                               OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                  OpenMode.ForWrite)
                For Each VrtxID As ObjectId In acPoly
                    Dim VrTx As Vertex2d = acTrans.GetObject(VrtxID, OpenMode.ForRead)
                    Dim pt As Point3d = VrTx.Position
                    acPts2d.Add(pt)
                Next
                acTrans.Commit()
                Return acPts2d
            Catch ex As Exception
                acTrans.Abort()
                AcEd.WriteMessage("Error :" & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function
    Public Function ConvertToPolyline2d(ByVal plineId As ObjectId) As Polyline2d
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = doc.Database
        Dim AcEd As Editor = doc.Editor

        Using acTrans As Transaction = acCurDb.TransactionManager.StartOpenCloseTransaction()
            Try
                Dim Pline As Polyline = acTrans.GetObject(plineId, OpenMode.ForWrite)

                acTrans.AddNewlyCreatedDBObject(Pline, False)

                Dim poly2 As Polyline2d = Pline.ConvertTo(True)

                acTrans.AddNewlyCreatedDBObject(poly2, True)

                acTrans.Commit()

                Return poly2
            Catch ex As Exception
                acTrans.Abort()
                AcEd.WriteMessage("Error :" & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function
End Module
