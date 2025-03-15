'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
'Imports Autodesk.AutoCAD.Geometry
'Imports Autodesk.Civil.DatabaseServices

'Public Module ParcelPolylineCreator

'    ' Función para crear una polilínea a partir de los bordes de una parcela
'    Public Sub CreateParcelBoundaryPolyline(parcelId As ObjectId)
'        ' Obtener el documento actual de AutoCAD
'        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'        Dim db As Database = doc.Database

'        ' Empezar una transacción
'        Using trans As Transaction = db.TransactionManager.StartTransaction()
'            Try
'                ' Abrir la parcela para lectura
'                Dim parcel As Parcel = trans.GetObject(parcelId, OpenMode.ForRead)

'                ' Crear una nueva polilínea
'                Using polyLine As New Polyline()
'                    ' Configurar los parámetros de la polilínea
'                    polyLine.ColorIndex = 1 ' Color: por ejemplo, rojo (color index 1)
'                    polyLine.Closed = True ' La polilínea estará cerrada

'                    ' Obtener los límites de la parcela como una colección de puntos
'                    Dim boundaryPoints As Point3dCollection = New Point3dCollection()
'                    parcel.Get(boundaryPoints)

'                    ' Agregar los puntos de los límites a la polilínea
'                    For Each point As Point3d In boundaryPoints
'                        polyLine.AddVertexAt(polyLine.NumberOfVertices, point.ToPoint2d(), 0, 0, 0)
'                    Next

'                    ' Abrir el espacio de modelos para escribir
'                    Dim blockTable As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
'                    Dim blockSpace As BlockTableRecord = trans.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

'                    ' Agregar la polilínea al espacio de modelos
'                    Dim polyLineId As ObjectId = blockSpace.AppendEntity(polyLine)
'                    trans.AddNewlyCreatedDBObject(polyLine, True)

'                    ' Completar la transacción
'                    trans.Commit()

'                    ' Informar al usuario
'                    doc.Editor.WriteMessage("Polilínea creada con éxito a partir de los bordes de la parcela.")
'                End Using
'            Catch ex As Exception
'                ' En caso de error, abortar la transacción
'                trans.Abort()
'                doc.Editor.WriteMessage($"Error al crear la polilínea: {ex.Message}")
'            End Try
'        End Using
'    End Sub

'End Module
