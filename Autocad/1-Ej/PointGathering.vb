Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports DbNurbSurface = Autodesk.AutoCAD.DatabaseServices.NurbSurface
Imports DbSurface = Autodesk.AutoCAD.DatabaseServices.Surface

Namespace PointGathering
    Public Class Commands
        <CommandMethod("GP", CommandFlags.UsePickSet)>
        Public Sub GatherPoints()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            ' Ask user to select entities
            Dim pso As New PromptSelectionOptions With {
                .MessageForAdding = vbLf & "Select entities to enclose: ",
                .AllowDuplicates = False,
                .AllowSubSelections = True,
                .RejectObjectsFromNonCurrentSpace = True,
                .RejectObjectsOnLockedLayers = False
            }

            Dim psr As PromptSelectionResult = ed.GetSelection(pso)
            If psr.Status <> PromptStatus.OK Then
                Return
            End If

            ' Collect points on the component entities
            Dim pts As New Point3dCollection()

            Dim acTrans As Transaction = db.TransactionManager.StartTransaction()
            Using acTrans
                Dim btr As BlockTableRecord = CType(acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

                For Each so As SelectedObject In psr.Value
                    Dim ent As Entity = CType(acTrans.GetObject(so.ObjectId, OpenMode.ForRead), Entity)

                    ' Collect the points for each selected entity
                    Dim entPts As New Point3dCollection()
                    CollectPoints(acTrans, ent, entPts)

                    ' Add a physical DBPoint at each Point3d
                    For Each pt As Point3d In entPts
                        Dim dbp As New DBPoint(pt)
                        btr.AppendEntity(dbp)
                        acTrans.AddNewlyCreatedDBObject(dbp, True)
                    Next
                Next
                acTrans.Commit()
            End Using
        End Sub

        Private Sub CollectPoints(ByVal acTrans As Transaction, ByVal ent As Entity, ByVal pts As Point3dCollection)
            ' We'll start by checking a block reference for attributes, getting their bounds and adding them to the point list.
            Dim br As BlockReference = TryCast(ent, BlockReference)
            If br IsNot Nothing Then
                For Each arId As ObjectId In br.AttributeCollection
                    Dim obj As DBObject = acTrans.GetObject(arId, OpenMode.ForRead)
                    If TypeOf obj Is AttributeReference Then
                        Dim ar As AttributeReference = DirectCast(obj, AttributeReference)
                        ExtractBounds(ar, pts)
                    End If
                Next
            End If

            ' For surfaces we'll collect points across its surface.
            Dim sur As DbSurface = TryCast(ent, DbSurface)
            If sur IsNot Nothing Then
                Dim nurbs As DbNurbSurface() = sur.ConvertToNurbSurface()
                For Each nurb As DbNurbSurface In nurbs
                    Dim ustart As Double = nurb.UKnots.StartParameter
                    Dim uend As Double = nurb.UKnots.EndParameter
                    Dim uinc As Double = (uend - ustart) / nurb.UKnots.Count
                    Dim vstart As Double = nurb.VKnots.StartParameter
                    Dim vend As Double = nurb.VKnots.EndParameter
                    Dim vinc As Double = (vend - vstart) / nurb.VKnots.Count

                    For u As Double = ustart To uend Step uinc
                        For v As Double = vstart To vend Step vinc
                            pts.Add(nurb.Evaluate(u, v))
                        Next
                    Next
                Next
            End If

            ' For 3D solids we'll fire a number of rays from the centroid in random directions in order to get a sampling of points on the outside.
            Dim sol As Solid3d = TryCast(ent, Solid3d)
            If sol IsNot Nothing Then
                For i As Integer = 0 To 499
                    Dim mp As Solid3dMassProperties = sol.MassProperties

                    Using pl As New Plane()
                        pl.Set(mp.Centroid, RandomVector3d())
                        Using reg As Region = sol.GetSection(pl)
                            Using ray As New Ray()
                                ray.BasePoint = mp.Centroid
                                ray.UnitDir = RandomVectorOnPlane(pl)

                                reg.IntersectWith(ray, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero)
                            End Using
                        End Using
                    End Using
                Next
            End If

            ' Now we start the terminal cases - for basic objects - before we recurse for more complex objects.
            Dim cur As Curve = TryCast(ent, Curve)
            If cur IsNot Nothing AndAlso Not (TypeOf cur Is Polyline OrElse TypeOf cur Is Polyline2d OrElse TypeOf cur Is Polyline3d) Then
                Dim segs As Integer = If(TypeOf ent Is Line, 2, 20)
                Dim param As Double = cur.EndParam - cur.StartParam

                For i As Integer = 0 To segs - 1
                    Try
                        Dim pt As Point3d = cur.GetPointAtParameter(cur.StartParam + (i * param / (segs - 1)))
                        pts.Add(pt)
                    Catch
                    End Try
                Next
            ElseIf TypeOf ent Is DBPoint Then
                pts.Add(DirectCast(ent, DBPoint).Position)
            ElseIf TypeOf ent Is DBText Then
                ExtractBounds(DirectCast(ent, DBText), pts)
            ElseIf TypeOf ent Is MText Then
                Dim txt As MText = DirectCast(ent, MText)
                Dim pts2 As Point3dCollection = txt.GetBoundingPoints()
                For Each pt As Point3d In pts2
                    pts.Add(pt)
                Next
            ElseIf TypeOf ent Is Face Then
                Dim f As Face = DirectCast(ent, Face)
                Try
                    For i As Short = 0 To 3
                        pts.Add(f.GetVertexAt(i))
                    Next
                Catch
                End Try
            ElseIf TypeOf ent Is Solid Then
                Dim s As Solid = DirectCast(ent, Solid)
                Try
                    For i As Short = 0 To 3
                        pts.Add(s.GetPointAt(i))
                    Next
                Catch
                End Try
            Else
                Dim oc As New DBObjectCollection()
                Try
                    ent.Explode(oc)
                    If oc.Count > 0 Then
                        For Each obj As DBObject In oc
                            Dim ent2 As Entity = TryCast(obj, Entity)
                            If ent2 IsNot Nothing AndAlso ent2.Visible Then
                                CollectPoints(acTrans, ent2, pts)
                            End If
                            obj.Dispose()
                        Next
                    End If
                Catch
                End Try
            End If
        End Sub

        Private Function RandomVectorOnPlane(ByVal pl As Plane) As Vector3d
            Dim ran As New Random()

            Dim absx As Double = ran.NextDouble()
            Dim absy As Double = ran.NextDouble()

            Dim x As Double = If(ran.NextDouble() < 0.5, -absx, absx)
            Dim y As Double = If(ran.NextDouble() < 0.5, -absy, absy)

            Dim v2 As New Vector2d(x, y)
            Return New Vector3d(pl, v2)
        End Function

        Private Function RandomVector3d() As Vector3d
            Dim ran As New Random()

            Dim absx As Double = ran.NextDouble()
            Dim absy As Double = ran.NextDouble()
            Dim absz As Double = ran.NextDouble()

            Dim x As Double = If(ran.NextDouble() < 0.5, -absx, absx)
            Dim y As Double = If(ran.NextDouble() < 0.5, -absy, absy)
            Dim z As Double = If(ran.NextDouble() < 0.5, -absz, absz)

            Return New Vector3d(x, y, z)
        End Function

        Private Sub ExtractBounds(ByVal txt As DBText, ByVal pts As Point3dCollection)
            If txt.Bounds.HasValue AndAlso txt.Visible Then
                Dim txt2 As New DBText With {
                    .Normal = Vector3d.ZAxis,
                    .Position = Point3d.Origin,
                    .TextString = txt.TextString,
                    .TextStyleId = txt.TextStyleId,
                    .LineWeight = txt.LineWeight
                }
                txt2.Thickness = txt2.Thickness
                txt2.HorizontalMode = txt.HorizontalMode
                txt2.VerticalMode = txt.VerticalMode
                txt2.WidthFactor = txt.WidthFactor
                txt2.Height = txt.Height
                txt2.IsMirroredInX = txt2.IsMirroredInX
                txt2.IsMirroredInY = txt2.IsMirroredInY
                txt2.Oblique = txt.Oblique

                If txt2.Bounds.HasValue Then
                    Dim maxPt As Point3d = txt2.Bounds.Value.MaxPoint

                    Dim bounds() As Point2d = {
                        Point2d.Origin,
                        New Point2d(0.0, maxPt.Y),
                        New Point2d(maxPt.X, maxPt.Y),
                        New Point2d(maxPt.X, 0.0)
                    }

                    Dim pl As New Plane(txt.Position, txt.Normal)

                    For Each pt As Point2d In bounds
                        pts.Add(pl.EvaluatePoint(pt.RotateBy(txt.Rotation, Point2d.Origin)))
                    Next
                End If
            End If
        End Sub
    End Class
End Namespace
