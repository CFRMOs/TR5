Imports System.Windows.Controls
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil

Public Class FeaturelineToPolyline
    <CommandMethod("ConvertFeaturelineToPolyline")>
    Public Shared Sub CMDConvertFeaturelineToPolyline(Optional ByRef HDString As String = "")
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Pedir al usuario que seleccione una Featureline
        Dim selOpts As New PromptEntityOptions(vbLf & "Seleccione una Featureline: ")
        selOpts.SetRejectMessage(vbLf & "Debe seleccionar una Featureline.")
        selOpts.AddAllowedClass(GetType(FeatureLine), False)
        Dim selRes As PromptEntityResult = ed.GetEntity(selOpts)

        If selRes.Status <> PromptStatus.OK Then
            ed.WriteMessage(vbLf & "Operación cancelada.")
            Return
        End If

        Using trans As Transaction = doc.Database.TransactionManager.StartTransaction()
            Try
                ' Obtener la Featureline seleccionada
                Dim featureline As FeatureLine = CType(trans.GetObject(selRes.ObjectId, OpenMode.ForRead), FeatureLine)

                ' Convertir la Featureline a Polyline
                Dim polyline As Polyline = ConvertFeaturelineToPolyline(featureline, True)

                polyline.Layer = featureline.Layer


                HDString = polyline.Handle.ToString()
                ' Confirmar la transacción
                trans.Commit()

            Catch ex As Exception
                ed.WriteMessage("Error :" & ex.Message)
                trans.Abort()
            Finally
                If Not trans.IsDisposed() Then trans.Dispose()
            End Try


        End Using
    End Sub

    Public Shared Function ConvertFeaturelineToPolyline(featureline As FeatureLine, Optional addtoModel As Boolean = False) As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Dim polyline As New Polyline()

        ' Obtener los puntos de la Featureline
        Dim points As Point3dCollection = featureline.GetPoints(FeatureLinePointType.AllPoints)
        Using trans As Transaction = doc.Database.TransactionManager.StartTransaction()
            Using doc.LockDocument()
                Try
                    ' Añadir los puntos a la Polyline
                    For i As Integer = 0 To points.Count - 1
                        Dim pt As Point3d = points(i)
                        Dim bulge As Double = 0 ' Valor inicial de bulge, ajustar si es necesario
                        If i <> points.Count - 1 Then
                            bulge = featureline.GetBulge(i)
                        End If

                        polyline.AddVertexAt(i, New Point2d(pt.X, pt.Y), bulge, 0, 0)
                    Next
                    ' Agregar la Polyline al espacio modelo
                    If polyline IsNot Nothing AndAlso addtoModel Then
                        Dim blkTbl As BlockTable = CType(trans.GetObject(doc.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim blkTblRec As BlockTableRecord = CType(trans.GetObject(blkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                        blkTblRec.AppendEntity(polyline)
                        trans.AddNewlyCreatedDBObject(polyline, True)
                    End If
                    trans.Commit()
                    Return polyline
                Catch ex As Exception
                    ed.WriteMessage("Error: " & ex.Message)
                    trans.Abort()
                    Return Nothing
                Finally
                    If Not trans.IsDisposed() Then trans.Dispose()
                End Try
            End Using
        End Using
    End Function
End Class
Public Class ShFeaturelineToPolyline
    Public Shared Function ConvertFeaturelineToPolyline(featureline As FeatureLine) As Polyline
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
End Class
