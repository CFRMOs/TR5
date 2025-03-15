Imports System.Windows
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
'AcadZoomManager.SelectedZoomEnt
Public Class AcadZoomManager
    Public Shared Sub ZoomToEntity(entidad As Entity)
        ' Obtener el editor activo de AutoCAD
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor
        Dim view As New ViewTableRecord()

        Dim min As Point3d = entidad.GeometricExtents.MinPoint
        Dim max As Point3d = entidad.GeometricExtents.MaxPoint
        ZoomToExtends(min, max)
    End Sub
    Public Shared Sub ZoomToExtends(min As Point3d, max As Point3d)
        ' Obtener el editor activo de AutoCAD
        Dim acDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor
        Dim view As New ViewTableRecord()
        ChangeToModelSpace()
        Dim min2d As New Point2d(min.X, min.Y)
        Dim max2d As New Point2d(max.X, max.Y)

        view.CenterPoint = min2d + ((max2d - min2d) / 2.0)

        view.Height = (max2d.Y - min2d.Y) * 2

        view.Width = (max2d.X - min2d.X) * 2

        acEd.SetCurrentView(view)
    End Sub
    Public Shared Sub ChangeToModelSpace()
        ' Obtener el documento actual y el editor
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Iniciar una transacción
        Using trans As Transaction = doc.TransactionManager.StartTransaction()
            ' Obtener el BlockTable y el BlockTableRecord actual
            Dim bt As BlockTable = trans.GetObject(doc.Database.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = trans.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForRead)

            ' Verificar si el espacio actual es un layout
            If btr.IsLayout And LayoutManager.Current.CurrentLayout <> "Model" Then
                ' Cambiar al espacio de modelo
                ' Obtener el LayoutManager
                Dim layoutManager As LayoutManager = LayoutManager.Current
                ' Cambiar a ModelSpace
                layoutManager.CurrentLayout = "Model"
                ed.WriteMessage(vbLf & "Se ha cambiado al espacio de modelo.")
            Else
                ed.WriteMessage(vbLf & "Ya estás en el espacio de modelo.")
            End If

            ' Confirmar la transacción
            trans.Commit()
        End Using
    End Sub

    Public Shared Sub SelectedZoom(hdString As String, ByRef ThisDrawing As Document)
        If hdString = vbNullString Then Exit Sub
        If ThisDrawing IsNot Nothing Then
            Application.DocumentManager.MdiActiveDocument = ThisDrawing ' Reactivar el documento original
        Else
            MessageBox.Show("El documento original no está disponible.")
            Return
        End If

        Dim entidad As Entity = CLHandle.GetEntityByStrHandle(hdString)
        ' Reactivar el documento original antes de continuar

        ' Verificar si se encontró la entidad y realizar la acción deseada (por ejemplo, hacer zoom a la entidad)
        SelectedZoomEnt(entidad)
    End Sub
    ' Verificar si se encontró la entidad y realizar la acción deseada (por ejemplo, hacer zoom a la entidad)
    Public Shared Sub SelectedZoomEnt(entidad As Entity)
        If entidad IsNot Nothing Then
            ' Realizar zoom a la entidad en AutoCAD
            ZoomToEntity(entidad)
            SelectByEntity(entidad)
        Else
            MessageBox.Show("No se encontró ninguna entidad con el handle especificado.")
        End If
    End Sub
End Class
