Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
Imports System.Linq

' Clase que contiene los comandos de AutoCAD
Public Class Cmds

    ' Comando para obtener los puntos 3D de un grupo de puntos en Civil 3D
    <CommandMethod("ObtenerPuntos3DDeGrupo")>
    Public Sub ObtenerBordes()
        ' Obtener el documento y el editor de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor

        ' Devuelve los bordes definidos por puntos aleatorios (CogoPoint)
        Dim pointGroupName As String = "a-ACC13T1 - LV. REGULARIZACION LOSA LI - KM 2+730 - 2+825"

        Dim copt As List(Of CogoPoint) = GetCOGOPointCollections(pointGroupName)

        Dim ListPL As New List(Of Polyline)() ' Inicialización de la lista de polilíneas

        Dim FirstCogoPoint As CogoPoint = copt(0)

        Dim Vertices As New List(Of Point3d)()

        ' Línea que converge en un solo vértice
        For Each Cogo As CogoPoint In copt
            Vertices.Add(Cogo.Location)
        Next

        Dim VerticesOrdenados As Dictionary(Of String, Point3d) = CGPointHelper.OrderVertices(Vertices)
        Dim pl As Polyline = CGPointHelper.CrearPL(VerticesOrdenados.Values.ToList())
        pl.Closed = True
        CGPointHelper.AddToModal(PL)
    End Sub

    ' Función para obtener la colección de CogoPoints
    Public Function GetCOGOPointCollections(pointGroupName As String) As List(Of CogoPoint)
        ' Obtener el documento actual de Civil 3D
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        ' Iniciar una transacción
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Obtener la colección de grupos de puntos de Civil 3D
                Dim civildoc As CivilDocument = CivilApplication.ActiveDocument
                Dim pointGroups As PointGroupCollection = civildoc.PointGroups
                Dim copt As New List(Of CogoPoint)()

                ' Verificar si hay grupos de puntos
                If pointGroups.Count > 0 Then
                    ' Obtener el nombre del primer grupo de puntos

                    If pointGroupName = [String].Empty Then
                        Return Nothing
                    End If

                    ' Obtener el grupo de puntos por nombre
                    Dim pointGroupId As ObjectId = GetPointGroupIdByName(pointGroupName)
                    Dim group As PointGroup = TryCast(pointGroupId.GetObject(OpenMode.ForRead), PointGroup)

                    ' Obtener los números de los puntos del grupo
                    Dim pointNumbers As UInteger() = group.GetPointNumbers()

                    ' Recorrer los números de puntos y obtener los CogoPoints correspondientes
                    For Each pointNumber As UInteger In pointNumbers
                        Dim colCogop As CogoPointCollection = civildoc.CogoPoints()
                        Dim ccogoPoint As CogoPoint = trans.GetObject(colCogop.GetPointByPointNumber(pointNumber), OpenMode.ForRead)
                        copt.Add(ccogoPoint)
                    Next
                Else
                    ed.WriteMessage(vbLf & "No hay grupos de puntos disponibles.")
                    Return Nothing
                End If

                ' Completar la transacción
                trans.Commit()
                Return copt
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Error: " & ex.Message)
                Return Nothing
            Finally
                trans.Dispose()
            End Try
        End Using
    End Function


    ' Obtener el ObjectId de un grupo de puntos por su nombre
    Private Function GetPointGroupIdByName(groupName As String) As ObjectId
        Dim civildoc As CivilDocument = CivilApplication.ActiveDocument
        Dim pointGroups As PointGroupCollection = civildoc.PointGroups

        If pointGroups.Contains(groupName) Then
            Dim pointGroupId As ObjectId = pointGroups(groupName)
            Return pointGroupId
        Else
            Throw New Exception("No se encontró el grupo de puntos con ese nombre.")
        End If
    End Function
End Class
