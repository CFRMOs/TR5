Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.LayerManager
Imports Autodesk.AutoCAD.Runtime

Public Class LayerColorChanger

    ' Definición de los tipos de cunetas y sus colores asociados
    Private Shared ReadOnly TiposCunetas As String() = {"BA-T1", "BA-T2", "DE-01", "DA-01", "DA-02",
                                                "CU-BO", "CU-BO1", "CU-T1", "CU-T2", "CU-TIPO-LIBRITO",
                                                "ZJ-DR1", "ZJ-DR2", "ZJ-DR3", "PV-T3", "PV-T2",
                                                "PV-T1", "BADEN", "Losa", "Aceras", "Bordillos"}

    '191,127,255 tipo librito
    Private Shared ReadOnly TColor As Color() = {Color.FromRgb(255, 0, 0),
                                         Color.FromRgb(0, 255, 0),
                                         Color.FromRgb(0, 0, 255),
                                         Color.FromRgb(255, 255, 0),
                                         Color.FromRgb(0, 255, 255),
                                         Color.FromRgb(255, 150, 0),
                                         Color.FromRgb(255, 127, 0),
                                         Color.FromRgb(15, 135, 255),
                                         Color.FromRgb(127, 0, 255),
                                         Color.FromRgb(82, 0, 165),
                                         Color.FromRgb(0, 165, 0),
                                         Color.FromRgb(128, 0, 128),
                                         Color.FromRgb(192, 192, 192),
                                         Color.FromRgb(128, 128, 128),
                                         Color.FromRgb(255, 128, 0),
                                         Color.FromRgb(128, 255, 0),
                                         Color.FromRgb(0, 128, 255),
                                         Color.FromRgb(128, 0, 255),
                                         Color.FromRgb(255, 128, 128),
                                         Color.FromRgb(128, 255, 128)}

    '191,127,255 tipo librito
    'Private Shared ReadOnly T As Color() = {Color.FromRgb(255, 0, 0),
    '                                     Color.FromRgb(0, 255, 0),
    '                                     Color.FromRgb(0, 0, 255),
    '                                     Color.FromRgb(255, 255, 0),
    '                                     Color.FromRgb(0, 255, 255),
    '                                     Color.FromRgb(255, 0, 255),
    '                                     Color.FromRgb(128, 0, 0),
    '                                     Color.FromRgb(15, 135, 255),
    '                                     Color.FromRgb(127, 0, 255),
    '                                     Color.FromRgb(128, 128, 0),
    '                                     Color.FromRgb(0, 165, 0),
    '                                     Color.FromRgb(128, 0, 128),
    '                                     Color.FromRgb(192, 192, 192),
    '                                     Color.FromRgb(128, 128, 128),
    '                                     Color.FromRgb(255, 128, 0),
    '                                     Color.FromRgb(128, 255, 0),
    '                                     Color.FromRgb(0, 128, 255),
    '                                     Color.FromRgb(128, 0, 255),
    '                                     Color.FromRgb(255, 128, 128),
    '                                     'Color.FromRgb(128, 255, 128)}

    ' Comando para cambiar el color a modo de prueba (debug)
    <CommandMethod("CMDChColorLY")>
    Public Sub CMDChColorLY()
        ' Llama a SetGPLYColor con el nombre del grupo "MED-15"
        SetGPLYColor("MED-14")
    End Sub

    ' Función para cambiar el color de un layer específico
    Public Shared Sub ChColorLY(layer As String, color As Color)
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Dim lt As LayerTable = trans.GetObject(AcCurD.LayerTableId, OpenMode.ForRead)
            If lt.Has(layer) Then
                Dim ltr As LayerTableRecord = trans.GetObject(lt(layer), OpenMode.ForWrite)
                ltr.Color = color
                ltr.LineWeight = 30
                ltr.DowngradeOpen()
            End If
            trans.Commit()
        End Using
    End Sub

    ' Función para recorrer el grupo dado de layers y cambiar sus colores
    Public Shared Sub SetGPLYColor(GName As String, Optional ApBaseC As Boolean = False)
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            ' Buscar el grupo de capas por su nombre en los filtros de capa
            Dim layerGroup As LayerGroup = GetLYGroup(GName)

            ' Verificar si se encontró el grupo de capas
            If layerGroup IsNot Nothing Then
                For Each layerId As ObjectId In layerGroup.LayerIds
                    Dim layer As LayerTableRecord = trans.GetObject(layerId, OpenMode.ForRead)
                    Dim layerName As String = layer.Name
                    If ApBaseC Then
                        ChColorLY(layerName, Color.FromRgb(102, 102, 102))
                    Else
                        For i As Integer = 0 To TiposCunetas.Length - 1
                            If layerName.Contains(TiposCunetas(i)) Then
                                ' Cambiar el color del layer
                                ChColorLY(layerName, TColor(i))
                                Exit For
                            End If
                        Next
                    End If
                    ' Verificar si el nombre del layer contiene alguno de los tipos especificados
                Next
            Else
                AcEd.WriteMessage(vbCrLf & "El grupo de capas '" & GName & "' no fue encontrado.")
            End If
            trans.Commit()
        End Using
    End Sub
    Public Shared Function GetLYGroup(GName As String) As LayerGroup
        Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim AcCurD As Database = AcDoc.Database
        Dim AcEd As Editor = AcDoc.Editor

        Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
            Try
                Dim layerFilters As LayerFilterTree = AcCurD.LayerFilters
                Dim layerGroup As LayerGroup = Nothing
                ' Buscar el grupo de capas por su nombre en los filtros de capa
                For Each filter As LayerFilter In layerFilters.Root.NestedFilters
                    If TypeOf filter Is LayerGroup AndAlso filter.Name = GName Then
                        layerGroup = DirectCast(filter, LayerGroup)
                        Exit For
                    End If
                Next
                Return layerGroup
                trans.Commit()
            Catch ex As Exception
                trans.Abort()
                AcEd.WriteMessage(("Exception: " & ex.Message))
                Return Nothing
            Finally
                trans.Dispose()
            End Try
        End Using

    End Function
    'canbios de medicion realtos en el dibujo 
    Public Shared Sub SetcurrentMed(ComMediciones As System.Windows.Forms.ComboBox)
        Dim LYCCH As New LayerColorChanger
        Dim Smed As String = ComMediciones.SelectedValue
        For Each Med As String In ComMediciones.Items
            If Smed <> Med Then SetGPLYColor(Med, True)
        Next
        SetGPLYColor(Smed)
    End Sub
    Public Shared Sub SetAllGPLYColor(GroupLayers As String())
        For Each Med As String In GroupLayers
            SetGPLYColor(Med)
        Next
    End Sub
End Class
' Función auxiliar para verificar la existencia de un grupo de capas (obsoleta en esta versión)
'Private Function GetLayerGroupByName(layerGroupName As String) As LayerGroup
'    Dim AcDoc As Document = Application.DocumentManager.MdiActiveDocument
'    Dim AcCurD As Database = AcDoc.Database
'    Dim AcEd As Editor = AcDoc.Editor
'
'    Using trans As Transaction = AcCurD.TransactionManager.StartTransaction()
'        Dim filterTree As LayerFilterTree = AcCurD.LayerFilters
'        Dim rootFilters As LayerFilterCollection = filterTree.Root.NestedFilters
'
'        ' Buscar el grupo de capas por su nombre
'        For Each filter As LayerFilter In rootFilters
'            If filter.Name = layerGroupName AndAlso TypeOf filter Is LayerGroup Then
'                Return DirectCast(filter, LayerGroup)
'            End If
'        Next
'        trans.Commit()
'    End Using
'
'    Return Nothing
'End Function