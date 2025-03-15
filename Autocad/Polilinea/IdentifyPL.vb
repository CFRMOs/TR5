'FindPolylineByLength:identifica una polilinea que cumple con una longitud dada a cierta precicion (0.000,0.00,0.0)
'que posea los mismos rango de estacionamientos dados 
'f(longitud as double,precicion as double, EstST as double,EstEnd as double, Alignment as Alignment) as entity
Imports System.Diagnostics
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry


Public Class IdentifyPL
    ''encontar la fila de la tabla que contenga el TabName buscado 
    ''devolver un listado con los calores en la columnas solisitada y la fila encontrada 
    Public Shared Function GetDataFromTbl(ByRef CAcadHelp As ACAdHelpers, Optional ByRef Handle As String = "") As List(Of Object)
        Try
            Dim Headers As New List(Of String) From {"Longitud", "StartStation", "EndStation", "Side", "AlignmentHDI"}
            Dim ExRNGSe As New ExcelRangeSelector()
            Dim Data As List(Of Object) = ExRNGSe.SelectRowOnTbl("CunetasGeneral", Headers, Handle)
            If Data?.Count <> 0 Then
                If CLHandle.CheckIfExistHd(Data(4)) Then
                    CAcadHelp.Alignment = CType(CLHandle.GetEntityByStrHandle(Data(4)), Alignment)
                End If
                Return Data
            End If
            Return Nothing
        Catch ex As Exception
            Debug.WriteLine("Error en IdentifyPL.GetDataFromTbl: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Shared Function SetALingment(ByRef CAcadHelp As ACAdHelpers, HandleAling As String) As Boolean

        If CLHandle.CheckIfExistHd(HandleAling) AndAlso CAcadHelp?.Alignment.Handle.ToString() <> HandleAling Then
            CAcadHelp.Alignment = CType(CLHandle.GetEntityByStrHandle(HandleAling), Alignment)
        End If
        Return CAcadHelp.Alignment IsNot Nothing AndAlso CAcadHelp.Alignment.Handle.ToString() = HandleAling
    End Function

    Public Shared Function FindPolyline(ByRef CAcadHelp As ACAdHelpers, Optional ByRef Handle As String = "") As String
        Try

            Dim Data As List(Of Object) = GetDataFromTbl(CAcadHelp, Handle)
            If Data?.Count <> 0 Then
                Dim AcEnt As Entity = IdentifyPL.Fpol(Data(2), 2, Data(0), Data(1), CAcadHelp)
                If AcEnt IsNot Nothing Then
                    Return AcEnt?.Handle.ToString()
                End If
            End If
            Return String.Empty
        Catch ex As Exception
            Debug.WriteLine("Error en IdentifyPL.Fpol: " & ex.Message)
            Return String.Empty
        End Try
    End Function
    Public Shared Function Fpol(longitud As Double, precision As Integer, EstST As Double, EstEnd As Double, ByRef CAcadHelp As ACAdHelpers) As Entity
        ' Obtener el documento actual y la base de datos
        CAcadHelp.CheckThisDrowing()

        Dim Tolerancia As Double = 2

        Dim AcEnt As Entity = FindPolylineByLength(longitud, precision, Tolerancia, EstST, EstEnd, CAcadHelp.Alignment)
        'TabName = AcEnt.TabName.ToString()

        Return AcEnt
    End Function

    ' Método para identificar una polilínea que cumple con los criterios dados
    Public Shared Function FindPolylineByLength(longitud As Double, precision As Integer, Tolerancia As Double, EstST As Double, EstEnd As Double, alignment As Alignment) As Entity
        ' Obtener el documento activo de AutoCAD
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        ' Obtener la base de datos asociada con el documento activo
        Dim db As Database = doc.Database
        ' Obtener el editor del documento (usado para mensajes de usuario y otras interacciones)
        Dim ed As Editor = doc.Editor

        ' Iniciar una transacción para leer y modificar la base de datos de AutoCAD
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            ' Abrir la tabla de bloques en modo lectura
            Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            ' Abrir el registro del espacio del modelo en modo lectura
            Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

            ' Recorrer cada objeto en el espacio del modelo
            For Each objId As ObjectId In btr
                ' Obtener la entidad asociada con el ObjectId
                Dim entity As Entity = CType(trans.GetObject(objId, OpenMode.ForRead), Entity)

                ' Verificar si la entidad es una polilínea
                If TypeOf entity Is Polyline Then
                    ' Convertir la entidad a una polilínea
                    Dim pline As Polyline = CType(entity, Polyline)
                    Dim side As String = String.Empty
                    ' Obtener la estación inicial y final de la polilínea usando el alineamiento proporcionado
                    Dim startStation As Double
                    Dim endStation As Double
                    Dim plineLength As Double
                    ' Verificar si la longitud de la polilínea está dentro del rango de precisión especificado
                    'considera un una tolerancia por error (+ o -) no mas ni meno de 2
                    If Math.Abs(Math.Round(pline.Length, precision) - Math.Round(longitud, precision)) <= Tolerancia Then

                        PLRelatedInf(pline, alignment, plineLength, startStation, endStation, side)
                        plineLength = Math.Round(plineLength, precision)
                        longitud = Math.Round(longitud, precision)
                        startStation = Math.Round(startStation, precision)
                        endStation = Math.Round(endStation, precision)


                        ' Verificar si las estaciones están dentro del rango de estaciones especificado
                        If (startStation >= EstST AndAlso startStation <= EstEnd) AndAlso (endStation >= EstST AndAlso endStation <= EstEnd) Then
                            ' Si la polilínea cumple con todos los criterios, devolverla
                            Return entity
                        Else
                            ed.WriteMessage("Errores de estacionamientos:")
                        End If
                    End If
                End If
            Next
            ' Si no se encuentra ninguna polilínea que cumpla con los criterios, devolver Nothing
            Return Nothing
        End Using
    End Function
    Shared Sub PLRelatedInf(PL As Polyline, Alignment As Alignment, Optional ByRef Len As Double = 0, Optional ByRef startStation As Double = 0, Optional ByRef endStation As Double = 0, Optional ByRef Side As String = "")
        Dim Area As Double
        Dim Side1 As String = String.Empty
        Dim Side2 As String = String.Empty
        Dim StartPT As Point3d
        Dim EndPT As Point3d

        CStationOffsetLabel.ProcessPolyline(Alignment, PL, StartPT, EndPT, Len, Area, startStation, endStation, Side1, Side2)
        CStationOffsetLabel.GetMxMnOPPL(startStation, endStation)

        If Side1 = Side2 Then Side = Side1
    End Sub

End Class
' Solicitar al usuario que seleccione un alineamiento
'Dim peo As New PromptEntityOptions(vbLf & "Seleccione un alineamiento:")
'peo.SetRejectMessage("Debe seleccionar un alineamiento." & vbLf)
'peo.AddAllowedClass(GetType(Alignment), False)
'Dim per As PromptEntityResult = editor.GetEntity(peo)
'If per.Status <> PromptStatus.OK Then
'    editor.WriteMessage(vbLf & "Comando cancelado.")
'    Exit Sub
'End If
'' Obtener el objeto alineamiento
'Dim alingId As ObjectId = per.ObjectId
'Dim alignment As Alignment = TryCast(trans.GetObject(alingId, OpenMode.ForRead), Alignment)

