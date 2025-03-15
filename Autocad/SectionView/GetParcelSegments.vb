Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
'<Assembly: CommandClass(GetType(GetParcelSegmentsVB.NET.MyCommands))>

'Namespace GetParcelSegmentsVB.NET
Public Class GetParcelSegments
    <CommandMethod("GetParcelSegments")>
    Public Sub GetParcelSegments()
        ' Obtener el documento activo de AutoCAD
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor

        ' Iniciar una transacción
        Using acTrans As Transaction = acDb.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = CType(acTrans.GetObject(acDb.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = CType(acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                ' Obtener el civil document
                ' Solicitar la selección de un sitio
                Dim sitePrompt As New PromptEntityOptions("Seleccione un sitio: ")
                sitePrompt.SetRejectMessage("Debe seleccionar un sitio.")
                sitePrompt.AddAllowedClass(GetType(Site), True)

                Dim siteResult As PromptEntityResult = ed.GetEntity(sitePrompt)
                If siteResult.Status <> PromptStatus.OK Then Return

                Dim siteId As ObjectId = siteResult.ObjectId
                Dim site As Parcel = TryCast(acTrans.GetObject(siteId, OpenMode.ForRead), Parcel)
                Dim pl As Polyline = site.BaseCurve()
                For Each objId As ObjectId In btr
                    Dim ent As Entity = CType(acTrans.GetObject(objId, OpenMode.ForRead), Entity)
                    If TypeName(ent) = "ParcelSegment" Then
                        Dim ParcelSeg As ParcelSegment = TryCast(acTrans.GetObject(ent.Id, OpenMode.ForRead), ParcelSegment)
                        'Dim parcelId As bound = ParcelSegment.BlockId()
                        Dim intersectPoints As New Point3dCollection()
                        'ParcelSegment.IntersectWith(site, Intersect.OnBothOperands, intersectPoints, IntPtr.Zero, IntPtr.Zero)
                    End If

                Next

                ' Confirmar la transacción
                acTrans.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
            End Try
        End Using
    End Sub
End Class
'End Namespace
