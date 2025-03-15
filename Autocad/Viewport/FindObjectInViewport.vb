Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Module AutoCADCommands

    ' Comando principal que encuentra un objeto en un viewport y activa el layout correspondiente
    <CommandMethod("FindObjectInViewport")>
    Public Sub FindObjectInViewport()
        ' Obtén el documento activo y su base de datos
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Try
            ' Solicita al usuario que seleccione un objeto
            Dim promptSelOpts As New PromptSelectionOptions With {
                .MessageForAdding = "Seleccione un objeto: "
            }
            Dim promptSelRes As PromptSelectionResult = ed.GetSelection(promptSelOpts)

            If promptSelRes.Status <> PromptStatus.OK Then
                ed.WriteMessage("No se seleccionó ningún objeto.")
                Return
            End If

            ' Solicita al usuario que seleccione un Alignment
            Dim promptEntityOpts As New PromptEntityOptions("Seleccione un Alignment: ")
            promptEntityOpts.SetRejectMessage("El objeto seleccionado no es un Alignment.")
            promptEntityOpts.AddAllowedClass(GetType(Alignment), False)
            Dim promptEntityRes As PromptEntityResult = ed.GetEntity(promptEntityOpts)

            If promptEntityRes.Status <> PromptStatus.OK Then
                ed.WriteMessage("No se seleccionó ningún Alignment.")
                Return
            End If

            Dim selectedObjId As ObjectId = promptSelRes.Value(0).ObjectId
            Dim layout As Layout = Nothing
            Dim planoName As String = PlanoNombre(selectedObjId, layout)

            If layout IsNot Nothing Then
                Dim List_ALing As List(Of Alignment) = GetAllAlignments(New List(Of String)())
                If List_ALing.Count > 0 Then
                    Dim RngLisST As (Double, Double) = GetAlignmentStationRange(GetViewportsInLayout(layout)(0), List_ALing(0))
                    ActivateLayout(layout.LayoutName)
                Else
                    ed.WriteMessage("No se encontraron alignments.")
                End If
            Else
                ed.WriteMessage("No se encontró ningún layout.")
            End If
        Catch ex As Exception
            ed.WriteMessage("Error en FindObjectInViewport: " & ex.Message)
        End Try
    End Sub

    ' Función que obtiene el rango de estacionamiento de un viewport y un alignment
    Public Function GetAlignmentStationRange(vp As Viewport, align As Alignment) As (Double, Double)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Try
            ' Obtener el límite del viewport en el sistema de coordenadas del mundo
            Dim vpExtents As Extents3d = GetViewportWorldExtents(vp)
            Dim minPoint As Point3d = vpExtents.MinPoint
            Dim maxPoint As Point3d = vpExtents.MaxPoint

            ' Convertir los puntos al sistema de coordenadas del alignment
            Dim startStation As Double = 0
            Dim endStation As Double = 0
            align.StationOffset(minPoint.X, minPoint.Y, 0, startStation)
            align.StationOffset(maxPoint.X, maxPoint.Y, 0, endStation)

            ' Asegurar que el rango esté en orden ascendente
            If startStation > endStation Then
                Dim temp As Double = startStation
                startStation = endStation
                endStation = temp
            End If

            Return (startStation, endStation)
        Catch ex As Exception
            ed.WriteMessage("Error en GetAlignmentStationRange: " & ex.Message)
            Return (0, 0)
        End Try
    End Function

    ' Función auxiliar que obtiene los límites del viewport en el sistema de coordenadas del mundo
    Private Function GetViewportWorldExtents(vp As Viewport) As Extents3d
        Try
            Dim vpTransform As Matrix3d = GetViewportTransform(vp)
            Dim center As Point3d = New Point3d(vp.ViewCenter.X, vp.ViewCenter.Y, 0).TransformBy(vpTransform)
            Dim halfWidth As Double = vp.Width / 2
            Dim halfHeight As Double = vp.Height / 2

            Dim minPoint As New Point3d(center.X - halfWidth, center.Y - halfHeight, 0)
            Dim maxPoint As New Point3d(center.X + halfWidth, center.Y + halfHeight, 0)

            Return New Extents3d(minPoint, maxPoint)
        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error en GetViewportWorldExtents: " & ex.Message)
            Return New Extents3d()
        End Try
    End Function

    ' Función que obtiene el nombre del plano y el layout correspondiente a un objeto seleccionado
    Public Function PlanoNombre(selectedObjId As ObjectId, Optional ByRef layout As Layout = Nothing) As String
        Try
            layout = GetLayout(selectedObjId)
            If layout Is Nothing Then Return String.Empty
            Dim planoName As String = ""
            'MED-HLP - ACC - 4 - DRE - 3
            ' Buscar el patrón en el layout

            For Each x As String In {"AGO-HLP-AC-jj-DRE-", "AGO-HLP-AC-jj-GEO-", "SMA-HLP-AC-jj-DRE-", "CID-HLP-AC-jj-DRE-", "MED-HLP-ACC-jj-DRE-"}
                For Each J As String In {"\d{1}", "\d{2}"}
                    If planoName = "" Then
                        BuscarPatronEnLayout(Application.DocumentManager.MdiActiveDocument.Database, layout, Replace(x, "jj", J), planoName)
                        If InStr(planoName, "NOTAS:", CompareMethod.Text) <> 0 Then planoName = ""

                        If InStr(planoName, "Model", CompareMethod.Text) <> 0 Then
                            If Format(layout.LayoutName, "000") = "000" Then
                                planoName = Replace(planoName, "Model", layout.LayoutName)
                            Else
                                planoName = Replace(planoName, "Model", Format(layout.LayoutName, "000"))
                            End If
                        ElseIf Not String.IsNullOrEmpty(planoName) Then
                            For Each xx As String In {"-\d{1}", "-\d{2}"}
                                Dim Patron As String = GetPatron(xx, planoName)
                                If Not String.IsNullOrEmpty(Patron) Then GoTo FPlano
                            Next
                        End If
                        If planoName <> "" Then GoTo FPlano
                    End If
                Next
            Next
FPlano:
            Return planoName
        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error en PlanoNombre: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    ' Función que obtiene el patrón de una cadena de texto
    Function GetPatron(Patron As String, text As String) As String
        Try
            Dim regex As New Regex(Patron, RegexOptions.IgnoreCase)
            Dim match As Match = regex.Match(text)

            If match.Success Then
                Return match.Value
            Else
                Return String.Empty
            End If
        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error en GetPatron: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    ' Función que obtiene el layout en el que se encuentra un objeto seleccionado
    Function GetLayout(selectedObjId As ObjectId) As Layout
        ' Obtén el documento activo y su base de datos
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim layoutDict As DBDictionary = trans.GetObject(db.LayoutDictionaryId, OpenMode.ForRead)
                For Each entry As DBDictionaryEntry In layoutDict
                    Dim layout As Layout = trans.GetObject(entry.Value, OpenMode.ForRead)
                    Dim viewports As List(Of Viewport) = GetViewportsInLayout(layout)

                    For Each vp As Viewport In viewports
                        If IsObjectInViewport(selectedObjId, vp, trans) Then
                            trans.Commit()
                            Return layout
                        End If
                    Next
                Next
                trans.Commit()
            Catch ex As Exception
                ed.WriteMessage("Error en GetLayout: " & ex.Message)
                trans.Abort()
                Return Nothing
            Finally
                If trans.IsDisposed = False Then trans.Dispose()
            End Try
        End Using
        Return Nothing
    End Function

    ' Función que obtiene todos los viewports en un layout específico
    Function GetViewportsInLayout(layout As Layout) As List(Of Viewport)
        ' Obtén el documento activo y su base de datos
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim viewports As New List(Of Viewport)
                Dim btr As BlockTableRecord = trans.GetObject(layout.BlockTableRecordId, OpenMode.ForRead)
                For Each objId As ObjectId In btr
                    Dim ent As Entity = TryCast(trans.GetObject(objId, OpenMode.ForRead), Entity)
                    If TypeOf ent Is Viewport Then
                        viewports.Add(CType(ent, Viewport))
                    End If
                Next
                trans.Commit()
                Return viewports
            Catch ex As Exception
                ed.WriteMessage("Error en GetViewportsInLayout: " & ex.Message)
                trans.Abort()
                Return Nothing
            End Try
        End Using
    End Function

    ' Función que verifica si un objeto está en un viewport específico
    Private Function IsObjectInViewport(objId As ObjectId, vp As Viewport, trans As Transaction) As Boolean
        Try
            Dim ent As Entity = TryCast(trans.GetObject(objId, OpenMode.ForRead), Entity)
            If ent Is Nothing Then
                Return False
            End If

            Dim vpTransform As Matrix3d = GetViewportTransform(vp)
            Dim ext As Extents3d = ent.GeometricExtents
            ext.TransformBy(vpTransform)

            Dim vpExtents As New Extents2d(vp.ViewCenter.X - vp.ViewHeight / 2, vp.ViewCenter.Y - vp.ViewHeight / 2, vp.ViewCenter.X + vp.ViewHeight / 2, vp.ViewCenter.Y + vp.ViewHeight / 2)
            Dim objExtents As New Extents2d(ext.MinPoint.X, ext.MinPoint.Y, ext.MaxPoint.X, ext.MaxPoint.Y)
            Return vpExtents.MinPoint.X < objExtents.MaxPoint.X AndAlso vpExtents.MaxPoint.X > objExtents.MinPoint.X AndAlso vpExtents.MinPoint.Y < objExtents.MaxPoint.Y AndAlso vpExtents.MaxPoint.Y > objExtents.MinPoint.Y
        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error en IsObjectInViewport: " & ex.Message)
            Return False
        End Try
    End Function

    ' Función que obtiene la transformación de un viewport
    Private Function GetViewportTransform(vp As Viewport) As Matrix3d
        Try
            Dim viewDir As Vector3d = vp.ViewDirection
            Dim viewCenter As Point2d = vp.ViewCenter
            Dim viewTarget As Point3d = vp.ViewTarget
            Return Matrix3d.WorldToPlane(viewDir) * Matrix3d.Displacement(New Point3d(viewCenter.X, viewCenter.Y, 0) - viewTarget)
        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error en GetViewportTransform: " & ex.Message)
            Return Matrix3d.Identity
        End Try
    End Function

    ' Función que busca un patrón en un layout y actualiza el nombre del plano
    Private Sub BuscarPatronEnLayout(db As Database, layout As Layout, patron As String, ByRef Optional planoName As String = "")
        Dim regex As New Regex(patron, RegexOptions.IgnoreCase)

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim btr As BlockTableRecord = trans.GetObject(layout.BlockTableRecordId, OpenMode.ForRead)
                For Each objId As ObjectId In btr
                    Dim ent As Entity = TryCast(trans.GetObject(objId, OpenMode.ForRead), Entity)
                    If ent IsNot Nothing Then
                        BuscarPatronEnEntidad(ent, regex, trans, planoName)
                        If planoName <> "" Then Exit Sub
                    End If
                Next
                trans.Commit()
            Catch ex As Exception
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error en BuscarPatronEnLayout: " & ex.Message)
                trans.Abort()
            End Try
        End Using
    End Sub

    ' Función que busca un patrón en una entidad específica
    Private Sub BuscarPatronEnEntidad(ent As Entity, regex As Regex, trans As Transaction, ByRef Optional planoName As String = "")
        Try
            If TypeOf ent Is DBText Then
                Dim dbText As DBText = CType(ent, DBText)
                If regex.IsMatch(dbText.TextString) Then
                    planoName = ActualizarCampo(dbText)
                End If
            ElseIf TypeOf ent Is MText Then
                Dim mText As MText = CType(ent, MText)
                If regex.IsMatch(mText.Contents) Then
                    planoName = ActualizarCampo(mText)
                End If
            ElseIf TypeOf ent Is BlockReference Then
                Dim blockRef As BlockReference = CType(ent, BlockReference)
                ' Recorre los atributos del BlockReference
                For Each attId As ObjectId In blockRef.AttributeCollection
                    Dim attRef As AttributeReference = CType(trans.GetObject(attId, OpenMode.ForRead), AttributeReference)
                    If regex.IsMatch(attRef.TextString) Then
                        planoName = ActualizarCampo(attRef)
                        Exit Sub
                    End If
                Next
                ' Recorre las entidades dentro del BlockTableRecord del BlockReference
                Dim btr As BlockTableRecord = CType(trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                For Each id As ObjectId In btr
                    Dim nestedEnt As Entity = CType(trans.GetObject(id, OpenMode.ForRead), Entity)
                    If nestedEnt IsNot Nothing Then
                        BuscarPatronEnEntidad(nestedEnt, regex, trans, planoName)
                        If planoName <> "" Then Exit Sub
                    End If
                Next
            End If
        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error en BuscarPatronEnEntidad: " & ex.Message)
        End Try
    End Sub

    ' Función que actualiza el campo de una entidad y devuelve su texto
    Private Function ActualizarCampo(ent As Entity) As String
        Try
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                If ent.ExtensionDictionary.IsValid Then
                    Dim extDict As DBDictionary = CType(trans.GetObject(ent.ExtensionDictionary, OpenMode.ForRead), DBDictionary)
                    ' Se puede agregar lógica adicional aquí si es necesario
                End If
                trans.Commit()
            End Using

            If TypeOf ent Is DBText Then
                Return CType(ent, DBText).TextString
            ElseIf TypeOf ent Is MText Then
                Return CType(ent, MText).Contents
            ElseIf TypeOf ent Is AttributeReference Then
                Return CType(ent, AttributeReference).TextString
            End If

            Return String.Empty
        Catch ex As Exception
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error en ActualizarCampo: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    'Función que activa un layout específico
    Private Sub ActivateLayout(layoutName As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Try
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Dim layoutMgr As LayoutManager = LayoutManager.Current
                layoutMgr.CurrentLayout = layoutName
                trans.Commit()
            End Using
            ed.WriteMessage(vbCrLf & "Layout " & layoutName & " activado.")
        Catch ex As Exception
            ed.WriteMessage("Error en ActivateLayout: " & ex.Message)
        End Try
    End Sub

End Module
