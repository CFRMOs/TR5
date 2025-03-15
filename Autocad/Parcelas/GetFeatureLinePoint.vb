'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.EditorInput
'Imports Autodesk.AutoCAD.Geometry
'Imports Autodesk.AutoCAD.Interop
'Imports Autodesk.AutoCAD.Interop.Common
'Imports Autodesk.AECC.Interop.Land
'Imports Autodesk.AECC.Interop.UiLand
'Imports System.Diagnostics
'Imports System.Globalization
'Imports System.Reflection
'Imports Autodesk.AutoCAD.Runtime

'Public Class GetFeatureLinePoints
'    <CommandMethod("GetFeatureLinePoints")>
'    Public Sub GetFeatureLinePoints()
'        Try
'            ' Obtener la aplicación de Civil 3D
'            Dim civilApp As Object = Civil3DHelper.GetCivilApp()
'            If civilApp Is Nothing Then
'                Debug.Print("Error: No se pudo obtener la aplicación de Civil 3D.")
'                Return
'            End If

'            ' Obtener el documento activo de Civil 3D
'            Dim oAeccDocument As AeccDocument = CType(civilApp.ActiveDocument, AeccDocument)

'            ' Obtener el editor del documento
'            Dim editor As Editor = Application.DocumentManager.MdiActiveDocument.Editor

'            ' Solicitar al usuario que seleccione una parcela
'            Dim objEnt As AcadObject = Nothing
'            Dim varPick As Object = Nothing

'            oAeccDocument.Utility.GetEntity(objEnt, varPick, "Seleccione la parcela")

'            If Not TypeOf objEnt Is AeccParcel Then
'                Debug.Print("La entidad seleccionada no es una parcela.")
'                Return
'            End If

'            Dim selectedParcel As AeccParcel = CType(objEnt, AeccParcel)
'            Debug.Print("Selected Parcel Name: " & selectedParcel.Name)

'            ' Convertir la parcela a polilínea 2D
'            Dim polyline As AcadPolyline = CreatePolylineFromParcel(selectedParcel)

'            ' Agregar la polilínea al espacio de trabajo del modelo
'            If polyline IsNot Nothing Then
'                ' Obtener el espacio de trabajo del modelo
'                Dim acadApp As AcadApplication = CType(Application.AcadApplication, AcadApplication)
'                Dim activeDoc As AcadDocument = acadApp.ActiveDocument
'                Dim modelSpace As AcadModelSpace = CType(activeDoc.ModelSpace, AcadModelSpace)

'                ' Agregar la polilínea al ModelSpace
'                modelSpace.AppendEntity(polyline)
'                activeDoc.Regen(AcRegenType.acActiveViewport)
'            End If

'        Catch ex As AccessViolationException
'            Debug.Print("AccessViolationException: " & ex.Message)
'        Catch ex As Exception
'            Debug.Print("Error: " & ex.Message)
'            Debug.Print(ex.StackTrace)
'        End Try
'    End Sub

'    Public Function CreatePolylineFromParcel(parcel As AeccParcel) As AcadPolyline
'        Try
'            ' Crear una lista de puntos para la polilínea
'            Dim vertexList As New List(Of Double)()

'            ' Obtener los segmentos de la parcela y agregar vértices a la lista
'            Dim loops As AeccParcelLoops = parcel.ParcelLoops

'            If loops Is Nothing Then
'                Debug.Print("No se pudo obtener los ParcelLoops.")
'                Return Nothing
'            End If

'            For Each loopSegment As AeccParcelLoop In loops
'                For Each segment As AeccParcelSegment In loopSegment
'                    Dim startPoint As AeccPoint = segment.StartPoint
'                    vertexList.Add(startPoint.X)
'                    vertexList.Add(startPoint.Y)
'                Next
'            Next

'            ' Crear la polilínea si hay suficientes puntos
'            If vertexList.Count >= 4 Then
'                ' Obtener el espacio de trabajo del modelo
'                Dim acadApp As AcadApplication = CType(Application.AcadApplication, AcadApplication)
'                Dim activeDoc As AcadDocument = acadApp.ActiveDocument
'                Dim modelSpace As AcadModelSpace = CType(activeDoc.ModelSpace, AcadModelSpace)

'                ' Agregar la polilínea al ModelSpace
'                Dim poly As AcadPolyline = modelSpace.AddPolyline(vertexList.ToArray())
'                poly.Closed = True
'                Return poly
'            Else
'                Debug.Print("No hay suficientes puntos para crear una polilínea.")
'                Return Nothing
'            End If

'        Catch ex As AccessViolationException
'            Debug.Print("AccessViolationException en CreatePolylineFromParcel: " & ex.Message)
'            Return Nothing
'        Catch ex As Exception
'            Debug.Print("Error en CreatePolylineFromParcel: " & ex.Message)
'            Debug.Print(ex.StackTrace)
'            Return Nothing
'        End Try
'    End Function
'End Class
