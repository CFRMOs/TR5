'Imports Autodesk.AutoCAD.DatabaseServices
'Imports Autodesk.AutoCAD.EditorInput
'Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.Civil.DatabaseServices

'Public Class SectionViewToPolyline

'    <CommandMethod("SectionViewToPolyline")>
'    Public Sub SectionViewToPolyline()
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim ed As Editor = doc.Editor
'        Dim db As Database = doc.Database

'        Try
'            ' Seleccionar el SectionView
'            Dim sectionViewPrompt As PromptEntityOptions = New PromptEntityOptions("Seleccione un SectionView: ")
'            sectionViewPrompt.SetRejectMessage("La entidad seleccionada no es un SectionView.")
'            sectionViewPrompt.AddAllowedClass(GetType(SectionView), True)

'            Dim sectionViewResult As PromptEntityResult = ed.GetEntity(sectionViewPrompt)
'            If sectionViewResult.Status <> PromptStatus.OK Then
'                Return
'            End If

'            Using transaction As Transaction = db.TransactionManager.StartTransaction()
'                Dim sectionViewId As ObjectId = sectionViewResult.ObjectId
'                Dim sectionView As SectionView = TryCast(transaction.GetObject(sectionViewId, OpenMode.ForRead), SectionView)

'                If sectionView IsNot Nothing Then
'                    ' Obtener la polylinea del terreno o Finish Grade
'                    Dim profile As Profile = GetProfileFromSectionView(sectionView, "Finish Grade")
'                    If profile IsNot Nothing Then
'                        Dim polyline As Polyline = ConvertProfileToPolyline(profile)
'                        If polyline IsNot Nothing Then
'                            ' Añadir la polylinea al dibujo
'                            Dim bt As BlockTable = TryCast(transaction.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
'                            Dim btr As BlockTableRecord = TryCast(transaction.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
'                            btr.AppendEntity(polyline)
'                            transaction.AddNewlyCreatedDBObject(polyline, True)
'                        End If
'                    End If
'                End If

'                transaction.Commit()
'            End Using

'        Catch ex As Exception
'            ed.WriteMessage("Error: " & ex.Message)
'        End Try
'    End Sub

'    Private Function GetProfileFromSectionView(sectionView As SectionView, profileName As String) As Profile
'        'sectionView.seg
'        For Each profileId As ObjectId In sectionView.GetProfileIds()
'            Dim profile As Profile = TryCast(profileId.GetObject(OpenMode.ForRead), Profile)
'            If profile IsNot Nothing AndAlso profile.Name = profileName Then
'                Return profile
'            End If
'        Next
'        Return Nothing
'    End Function

'    Private Function ConvertProfileToPolyline(profile As Profile) As Polyline
'        '    Dim polyline As New Polyline()
'        '    Dim pointIndex As Integer = 0

'        '    For Each pv As ProfileViewEntity In profile.ProfileViewEntities
'        '        Dim pnt As Autodesk.AutoCAD.Geometry.Point3d = pv.Location
'        '        polyline.AddVertexAt(pointIndex, New Autodesk.AutoCAD.Geometry.Point2d(pnt.X, pnt.Y), 0, 0, 0)
'        '        pointIndex += 1
'        '    Next

'        '    Return polyline
'    End Function

'End Class

