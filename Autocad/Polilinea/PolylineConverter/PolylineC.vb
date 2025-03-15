Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil

Public Module PolylineC

    Public Function CPoLy3dToPL(acPoly3d As Polyline3d, Optional AP As Boolean = True) As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = doc.Database
        Dim acEd As Editor = doc.Editor
        Dim acPts3d As Point3dCollection = CollectPL3DPoints(acPoly3d)
        Dim acPts2d As New Point2dCollection()

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                '' Open the Block table for read
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim PL As New Polyline()
                Dim i As Long = 0

                For Each pt As Point3d In acPts3d
                    PL.AddVertexAt(i, New Point2d(pt.X, pt.Y), 0, 0, 0)
                    i += 1
                Next

                If AP Then
                    acBlkTblRec.AppendEntity(PL)
                    acTrans.AddNewlyCreatedDBObject(PL, True)
                End If
                acTrans.Commit()
                Return PL
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage("Error :" & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function

    Public Function CPoLy2dToPL(acPoly2d As Polyline2d, Optional AP As Boolean = True) As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = doc.Database
        Dim acEd As Editor = doc.Editor
        Dim acPts3d As Point3dCollection = CollectPL2dPoints(acPoly2d)
        Dim acPts2d As New Point2dCollection()

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                '' Open the Block table for read
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim PL As New Polyline()
                Dim i As Long = 0
                For Each pt As Point3d In acPts3d
                    PL.AddVertexAt(i, New Point2d(pt.X, pt.Y), 0, 0, 0)
                    i += 1
                Next
                If AP Then
                    acBlkTblRec.AppendEntity(PL)
                    acTrans.AddNewlyCreatedDBObject(PL, True)
                End If
                acTrans.Commit()


                Return PL
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage("Error :" & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function

    Public Function CrearPLWR(acPoly As Polyline) As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = doc.Database
        Dim acEd As Editor = doc.Editor
        Dim acPts2d As Point2dCollection = CollectPLPoints(acPoly)

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                '' Open the Block table for read
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim PL As New Polyline()
                Dim i As Long = 0
                For Each pt As Point2d In acPts2d
                    PL.AddVertexAt(i, New Point2d(pt.X, pt.Y), 0, 0, 0)
                    i += 1
                Next

                ' Add the new Polyline to the Block Table Record
                acBlkTblRec.AppendEntity(PL)
                acTrans.AddNewlyCreatedDBObject(PL, True)
                acTrans.Commit()

                Return PL
            Catch ex As Exception
                acTrans.Abort()
                acEd.WriteMessage("Error :" & ex.Message)
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function

    Public Function CreatePolyline(points As List(Of Point3d)) As Polyline
        ' Obtener el documento actual y la base de datos
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        ' Crear una nueva polilínea
        Dim acPoly As New Polyline()

        ' Comenzar una transacción
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Abrir el bloque de espacio modelo para escritura
                Dim blkTable As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim blkTableRecord As BlockTableRecord = acTrans.GetObject(blkTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                ' Añadir los puntos a la polilínea
                For i As Integer = 0 To points.Count - 1
                    Dim pt As Point3d = points(i)
                    acPoly.AddVertexAt(i, New Point2d(pt.X, pt.Y), 0, 0, 0)
                Next

                ' Añadir la polilínea al espacio modelo y a la transacción
                blkTableRecord.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)

                ' Guardar los cambios y finalizar la transacción
                acTrans.Commit()

                ' Devolver la polilínea creada
                Return acPoly
            Catch ex As System.Exception
                ' Manejar cualquier error
                acDoc.Editor.WriteMessage("Error: " & ex.Message)
                acTrans.Abort()
                Return Nothing
            Finally
                acTrans.Dispose()
            End Try
        End Using
    End Function
    Public Function ConvertLineToPolyline(acLine As Line, Optional AP As Boolean = True) As Polyline
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database

        ' Crear una nueva polilínea a partir de la línea
        Dim acPoly As New Polyline()

        ' Iniciar una transacción
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                ' Añadir los puntos de la línea a la polilínea
                acPoly.AddVertexAt(0, New Point2d(acLine.StartPoint.X, acLine.StartPoint.Y), 0, 0, 0)
                acPoly.AddVertexAt(1, New Point2d(acLine.EndPoint.X, acLine.EndPoint.Y), 0, 0, 0)

                ' Añadir la polilínea al espacio modelo
                Dim blkTable As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim blkTableRecord As BlockTableRecord = acTrans.GetObject(blkTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                If AP Then
                    blkTableRecord.AppendEntity(acPoly)
                    acTrans.AddNewlyCreatedDBObject(acPoly, True)
                End If

                ' Guardar los cambios
                acTrans.Commit()

                ' Devolver la polilínea creada
                Return acPoly
            Catch ex As System.Exception
                acEd.WriteMessage("Error: " & ex.Message)
                acTrans.Abort()
                Return Nothing
            End Try
        End Using
    End Function

    Public Function ConvertFeaturelineToPolyline(featureline As FeatureLine) As Polyline
        Dim polyline As New Polyline()

        ' Obtener los puntos de la Featureline
        Dim points As Point3dCollection = featureline.GetPoints(FeatureLinePointType.AllPoints)

        ' Añadir los puntos a la Polyline
        For i As Integer = 0 To points.Count - 1
            Dim pt As Point3d = points(i)
            polyline.AddVertexAt(i, New Point2d(pt.X, pt.Y), 0, 0, 0)
        Next

        Return polyline
    End Function

End Module
