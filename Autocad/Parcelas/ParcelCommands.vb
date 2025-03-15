Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class ParcelCommands
    <CommandMethod("CreateParcelPolyline")>
    Public Sub CreateParcelPolyline()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        Try
            ' Obtener la parcela seleccionada
            Dim parcel As Parcel = SelectParcel(ed)

            If parcel Is Nothing Then
                ed.WriteMessage(vbLf & "The selected object is not a Parcel.")
                Return
            End If

            ' Crear la polilínea de la parcela
            CreatePolylineFromParcel(parcel, doc.Database, ed)

            ed.WriteMessage(vbLf & "Polyline created successfully.")
        Catch ex As Exception
            ed.WriteMessage(vbLf & "Error: " & ex.Message)
        End Try
    End Sub
    Public Function GetParcelBoudary(CatchHandle As String, Optional ByRef IDNum As Integer = 0) As Polyline
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Dim Polyline As Polyline = Nothing
        Using trans As Transaction = db.TransactionManager.StartTransaction()

            Try
                Dim parcel As Parcel = TryCast(trans.GetObject(CLHandle.GetObjIdBStr(CatchHandle), OpenMode.ForRead), Parcel)
                CreatePolylineFromParcel(parcel, db, ed, Polyline)
                IDNum = parcel.Number
                trans.Commit()
                Return Polyline
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
                trans.Abort()
                Return Nothing
            Finally
                If trans.IsDisposed = False Then trans.Dispose()
            End Try
        End Using
    End Function


    Public Function SelectParcel(ed As Editor) As Parcel
        ' Solicitar al usuario que seleccione una parcela
        Dim parcelPrompt As New PromptEntityOptions(vbLf & "Select a Parcel: ")
        parcelPrompt.SetRejectMessage(vbLf & "Object is not a Parcel.")
        parcelPrompt.AddAllowedClass(GetType(Parcel), True)

        Dim parcelResult As PromptEntityResult = ed.GetEntity(parcelPrompt)
        If parcelResult.Status <> PromptStatus.OK Then
            ed.WriteMessage(vbLf & "Command canceled.")
            Return Nothing
        End If

        ' Abrir la transacción para obtener la parcela seleccionada
        Dim db As Database = ed.Document.Database
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim parcel As Parcel = TryCast(trans.GetObject(parcelResult.ObjectId, OpenMode.ForRead), Parcel)
            trans.Commit()
            Return parcel
        End Using
    End Function

    Public Sub CreatePolylineFromParcel(parcel As Parcel, db As Database, ed As Editor, Optional ByRef polyline As Polyline = Nothing)
        If ed Is Nothing Then
            Throw New ArgumentNullException(NameOf(ed))
        End If
        ' Abrir la transacción para crear la polilínea de la parcela
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            polyline = parcel.BaseCurve2d()

            ' Agregar la polilínea al espacio actual
            Dim blockTable As BlockTable = TryCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim blockTableRecord As BlockTableRecord = TryCast(trans.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
            blockTableRecord.AppendEntity(polyline)
            trans.AddNewlyCreatedDBObject(polyline, True)

            ' Confirmar la transacción
            trans.Commit()
        End Using
    End Sub
End Class

Public Module ParcelExtensions
    <System.Runtime.CompilerServices.Extension>
    Public Function BaseCurve2d(parcel As Parcel) As Polyline
        ' Intentar convertir BaseCurve a Polyline
        Dim poly As Polyline = TryCast(parcel.BaseCurve, Polyline)
        If poly IsNot Nothing Then
            Return poly
        End If

        ' Crear una nueva Polilínea para la representación 2D
        poly = New Polyline()

        ' Obtener el objeto COM para la parcela
        Dim comParcel As Object = parcel.AcadObject

        ' Obtener los bucles de la parcela
        Dim loops As Object = comParcel.GetType().InvokeMember("ParcelLoops", Reflection.BindingFlags.GetProperty, Nothing, comParcel, New Object() {0})

        ' Recorrer cada segmento en el bucle de la parcela y agregar vértices a la polilínea
        Dim count As Integer = CInt(loops.GetType().InvokeMember("Count", Reflection.BindingFlags.GetProperty, Nothing, loops, Nothing))
        For i As Integer = 0 To count - 1
            Dim segment As Object = loops.GetType().InvokeMember("Item", Reflection.BindingFlags.GetProperty, Nothing, loops, New Object() {i})
            Dim startX As Double = CDbl(segment.GetType().InvokeMember("StartX", Reflection.BindingFlags.GetProperty, Nothing, segment, Nothing))
            Dim startY As Double = CDbl(segment.GetType().InvokeMember("StartY", Reflection.BindingFlags.GetProperty, Nothing, segment, Nothing))
            Dim startZ As Double = CDbl(segment.GetType().InvokeMember("StartZ", Reflection.BindingFlags.GetProperty, Nothing, segment, Nothing))
            Dim endX As Double = CDbl(segment.GetType().InvokeMember("EndX", Reflection.BindingFlags.GetProperty, Nothing, segment, Nothing))
            Dim endY As Double = CDbl(segment.GetType().InvokeMember("EndY", Reflection.BindingFlags.GetProperty, Nothing, segment, Nothing))
            Dim endZ As Double = CDbl(segment.GetType().InvokeMember("EndZ", Reflection.BindingFlags.GetProperty, Nothing, segment, Nothing))
            Dim bulge As Double = 0

            ' Verificar si la propiedad Bulge existe y obtener su valor
            Dim bulgeProperty = segment.GetType().GetProperty("Bulge")
            If bulgeProperty IsNot Nothing Then
                bulge = CDbl(bulgeProperty.GetValue(segment, Nothing))
            End If

            ' Proyectar las coordenadas 3D en un plano 2D (ignorando la coordenada Z)
            Dim startPoint2D As New Point2d(startX, startY)
            Dim endPoint2D As New Point2d(endX, endY)

            ' Agregar el vértice a la polilínea
            poly.AddVertexAt(i, startPoint2D, bulge, 0, 0)
        Next

        poly.Closed = True
        Return poly
    End Function
End Module
