Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Public Class PolilineasAyacentes

    Public Shared Function ObtenerPolilineasAdyacentes(PL As Polyline) As List(Of (Polyline, List(Of Polyline)))
        Dim resultado As New List(Of (Polyline, List(Of Polyline)))

        ' Obtener el documento actual de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener el bloque modelo (ModelSpace)
                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                ' Obtener puntos inicial y final de la polilínea dada
                Dim puntoInicial As Point3d = PL.StartPoint
                Dim puntoFinal As Point3d = PL.EndPoint

                ' Crear una lista para almacenar las polilíneas adyacentes
                Dim polilineasAdyacentes As New List(Of Polyline)

                ' Recorrer todas las entidades en el bloque modelo
                For Each objId As ObjectId In btr
                    Dim entidad As Entity = trans.GetObject(objId, OpenMode.ForRead)
                    If TypeOf entidad Is Polyline AndAlso Not objId.Equals(PL.ObjectId) Then
                        Dim otraPolilinea As Polyline = CType(entidad, Polyline)

                        ' Obtener puntos inicial y final de la polilínea actual
                        Dim otraPuntoInicial As Point3d = otraPolilinea.StartPoint
                        Dim otraPuntoFinal As Point3d = otraPolilinea.EndPoint

                        ' Verificar si comparte punto inicial o final con la polilínea dada
                        If puntoInicial.IsEqualTo(otraPuntoInicial) OrElse puntoInicial.IsEqualTo(otraPuntoFinal) OrElse
                           puntoFinal.IsEqualTo(otraPuntoInicial) OrElse puntoFinal.IsEqualTo(otraPuntoFinal) Then
                            polilineasAdyacentes.Add(otraPolilinea)
                        End If
                    End If
                Next

                ' Agregar la polilínea dada y su lista de polilíneas adyacentes al resultado
                resultado.Add((PL, polilineasAdyacentes))

                ' Confirmar los cambios
                trans.Commit()
            Catch ex As Exception
                ed.WriteMessage("Error: " & ex.Message)
            Finally
                trans.Dispose()
            End Try
        End Using

        Return resultado
    End Function

End Class

