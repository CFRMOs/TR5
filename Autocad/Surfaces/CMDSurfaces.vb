Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Public Class CMDSurfaces
    <CommandMethod("CMDSurfaces")>
    Public Sub PrintOutSurfaces()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                For Each objId As ObjectId In acBlkTblRec
                    Dim acObj As Object = acTrans.GetObject(objId, OpenMode.ForWrite)

                    If TypeOf acObj Is TinSurface Then
                        Dim surface As TinSurface = CType(acObj, TinSurface)

                        ' Obtener el límite de la superficie como una polilínea
                        'Dim boundaryPolyline As Polyline = GetSurfaceBoundary(surface)

                        ' Imprimir el nombre de la superficie y el área
                        acEd.WriteMessage($"Nombre de la superficie: {surface.Name}" & vbLf)
                        acEd.WriteMessage($"Área de la superficie: {surface.Area}" & vbLf)
                        'acEd.WriteMessage($"Área de la superficie:" & boundaryPolyline.Length & vbLf)
                        ' Agregar el código necesario para manipular la polilínea 'boundaryPolyline' según tus necesidades

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

End Class
