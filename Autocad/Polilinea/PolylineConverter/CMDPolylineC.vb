Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class CMDPolylineC
    <CommandMethod("CreatePolylineFromPoints")>
    Public Sub CreatePolylineFromPoints()
        ' Obtener el documento actual y la base de datos
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor

        ' Solicitar los puntos al usuario
        Dim points As New List(Of Point3d)
        Dim promptPointOptions As New PromptPointOptions("Selecciona el primer punto o presiona ENTER para terminar:")
        Dim promptPointResult As PromptPointResult = acEd.GetPoint(promptPointOptions)

        While promptPointResult.Status = PromptStatus.OK
            points.Add(promptPointResult.Value)
            promptPointOptions.Message = "Selecciona el siguiente punto o presiona ENTER para terminar:"
            promptPointResult = acEd.GetPoint(promptPointOptions)
        End While

        ' Verificar si se introdujeron puntos
        If points.Count > 1 Then
            ' Llamar a la función que crea la polilínea y obtener la polilínea creada
            Dim polyline As Polyline = CreatePolyline(points)

            ' Verificar si la polilínea fue creada correctamente
            If polyline IsNot Nothing Then
                acDoc.Editor.WriteMessage(vbLf & "Polilínea creada exitosamente.")
            Else
                acDoc.Editor.WriteMessage(vbLf & "Error al crear la polilínea.")
            End If
        Else
            acDoc.Editor.WriteMessage(vbLf & "Se requieren al menos dos puntos para crear una polilínea.")
        End If
    End Sub
    <CommandMethod("CreatePolylineFrLine")>
    Public Sub CreatePolylineFrLine()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor

        ' Solicitar la selección de una línea al usuario
        Dim promptEntityOptions As New PromptEntityOptions("Selecciona una línea:")
        promptEntityOptions.SetRejectMessage("Elija una línea.")
        promptEntityOptions.AddAllowedClass(GetType(Line), False)
        Dim promptEntityResult As PromptEntityResult = acEd.GetEntity(promptEntityOptions)

        If promptEntityResult.Status = PromptStatus.OK Then
            ' Iniciar una transacción
            Using acTrans As Transaction = acDoc.TransactionManager.StartTransaction()
                Try
                    ' Obtener la línea seleccionada
                    Dim acLine As Line = acTrans.GetObject(promptEntityResult.ObjectId, OpenMode.ForRead)
                    ConvertLineToPolyline(acLine)
                Catch ex As System.Exception
                    acDoc.Editor.WriteMessage("Error: " & ex.Message)
                    acTrans.Abort()
                End Try
            End Using
        Else
            acDoc.Editor.WriteMessage(vbLf & "Comando cancelado.")
        End If
    End Sub
    <CommandMethod("CPolylineFr3dPL")>
    Public Sub CPolylineFr3dPL()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor

        ' Solicitar la selección de una línea al usuario
        Dim promptEntityOptions As New PromptEntityOptions("Selecciona una 3DPolyline:")
        promptEntityOptions.SetRejectMessage("Elija una 3DPolyline.")
        promptEntityOptions.AddAllowedClass(GetType(Polyline3d), False)
        Dim promptEntityResult As PromptEntityResult = acEd.GetEntity(promptEntityOptions)

        If promptEntityResult.Status = PromptStatus.OK Then
            ' Iniciar una transacción
            Using acTrans As Transaction = acDoc.TransactionManager.StartTransaction()
                Try
                    ' Obtener la línea seleccionada
                    Dim acPolyline3d As Polyline3d = acTrans.GetObject(promptEntityResult.ObjectId, OpenMode.ForRead)
                    CrearPLWR(CPoLy3dToPL(acPolyline3d))
                Catch ex As System.Exception
                    acDoc.Editor.WriteMessage("Error: " & ex.Message)
                    acTrans.Abort()
                End Try
            End Using
        Else
            acDoc.Editor.WriteMessage(vbLf & "Comando cancelado.")
        End If
    End Sub
    <CommandMethod("CMDDPL")>
    Public Sub Dividitestpl(Optional ByRef Polyline As Polyline = Nothing, Optional ByRef Alignment As Alignment = Nothing, Optional Station As Double = 0, Optional ByRef result As List(Of Polyline) = Nothing)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor
        Dim point As Point3d
        Dim PolyId As ObjectId

        ' Verifica si se proporciona la polilínea
        If Polyline Is Nothing Then
            Dim promptLineOptions As New PromptEntityOptions("Seleccionar una polilínea")
            promptLineOptions.SetRejectMessage("Debe seleccionar una polilínea.")
            promptLineOptions.AddAllowedClass(GetType(Polyline), True)

            Dim promptResult As PromptEntityResult = acEd.GetEntity(promptLineOptions)

            If promptResult.Status <> PromptStatus.OK Then Exit Sub ' Salir si el usuario cancela o hay error

            PolyId = promptResult.ObjectId

            ' Abrir la polilínea en modo lectura
            Using trans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Polyline = TryCast(trans.GetObject(PolyId, OpenMode.ForWrite), Polyline)
                If Polyline Is Nothing Then
                    acEd.WriteMessage("La entidad seleccionada no es una polilínea.")
                    Exit Sub
                End If
                Polyline.Elevation = 0
                trans.Commit()
            End Using
        Else
            PolyId = Polyline.Id
        End If

        Dim modifier As New PolylineModifier()
        modifier.ChangePolylineProperties(Polyline.Handle.ToString(), "FLUXO3", 5.5, 0.5, Color.FromColorIndex(ColorMethod.ByBlock, 0))
        ' Obtener el punto de división
        If Station = 0 Then
            ' Si no se proporciona la estación, pedimos un punto al usuario
            Dim promptPointOptions As New PromptPointOptions("Seleccionar un punto en el dibujo:")
            Dim promptPointResult As PromptPointResult = acEd.GetPoint(promptPointOptions)

            If promptPointResult.Status <> PromptStatus.OK Then Exit Sub ' Salir si el usuario cancela

            point = promptPointResult.Value
        Else
            ' Si se proporciona una estación, obtenemos el punto correspondiente
            Dim PtStation As Point3d = CStationOffsetLabel.GetPoint3dByStation(Station, Alignment)
            Dim Vector As Vector3d = GetPerpendicularVectorFromAlignment(Alignment, Station)
            'Usar el vector para crear una polilinea y extenderla la polilinea en cuestion
            point = FindIntersectionWithPolyline(PtStation, Vector, Polyline)

            CGPointHelper.AddCGPoint(PtStation.X, PtStation.Y, PtStation.Z, "ESTACION:" & Station.ToString("0+000.00"))

            CGPointHelper.AddCGPoint(point.X, point.Y, point.Z, "Punto Divisor at:" & Station.ToString("0+000.00"))

        End If

        ' Dividir la polilínea en el punto encontrado
        'Dim result As New List(Of Polyline)
        result = NewPolilineaDivider.DividirPolilineaEnPunto(point, PolyId)

    End Sub
    ' Función para extender una línea infinitamente y encontrar la intersección con una polilínea
    Public Function FindIntersectionWithPolyline(PtStation As Point3d, vector As Vector3d, polyline As Polyline) As Point3d
        ' Crear una línea que se extiende en la dirección del vector (usaremos una extensión grande en lugar de infinito)
        Dim lineStart As Point3d = PtStation - (vector * 10000) ' Extender hacia atrás 10,000 unidades
        Dim lineEnd As Point3d = PtStation + (vector * 10000) ' Extender hacia adelante 10,000 unidades
        Dim extendedLine As New Line(lineStart, lineEnd)

        ' Iniciar una transacción para buscar la intersección
        Using acTrans As Transaction = polyline.Database.TransactionManager.StartTransaction()
            Try
                ' Encontrar la intersección entre la línea extendida y la polilínea
                Dim intersectionPoints As New Point3dCollection()
                polyline.IntersectWith(extendedLine, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero)

                ' Si se encuentra una intersección, devolver el primer punto de intersección
                If intersectionPoints.Count > 0 Then
                    Return intersectionPoints(0)
                Else
                    ' Si no se encuentra intersección, devolver un punto nulo o un valor por defecto
                    Return Point3d.Origin ' O podrías devolver un valor nulo personalizado
                End If

            Catch ex As Exception
                ' Manejar cualquier error
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Error al calcular la intersección: " & ex.Message)
                Return Point3d.Origin ' O manejar el error de otra forma
            Finally
                acTrans.Commit()
            End Try
        End Using
    End Function

    ' Método para obtener el vector perpendicular al alineamiento en una estación
    Private Function GetPerpendicularVectorFromAlignment(Alignment As Alignment, Station As Double) As Vector3d
        ' Avanzar una pequeña distancia (1 mm o 0.001 unidades) a lo largo del alineamiento
        Dim station2 As Double = Station + 0.001

        ' Obtener los puntos en el alineamiento en la estación original y la avanzada
        Dim ptStation As Point3d = CStationOffsetLabel.GetPoint3dByStation(Station, Alignment)
        Dim ptStation2 As Point3d = CStationOffsetLabel.GetPoint3dByStation(station2, Alignment)

        ' Calcular el vector tangente entre los dos puntos
        Dim tangentVec As Vector3d = ptStation.GetVectorTo(ptStation2).GetNormal()

        ' Obtener el vector perpendicular al tangente
        Dim perpendicularVec As Vector3d = tangentVec.CrossProduct(Vector3d.ZAxis).GetNormal()

        Return perpendicularVec
    End Function



    Public Function GetOrtoPointOnLineByPointStation(Polyline As Polyline, PtStation As Point3d) As Point3d
        Dim closestPoint As Point3d = Point3d.Origin
        Dim minDistance As Double = Double.MaxValue
        Dim bulge As Double
        Dim closestSegmentIndex As Integer = -1 ' Para almacenar el índice del segmento más cercano
        Dim isCurve As Boolean = False ' Indica si el segmento más cercano es curvo

        ' Recorremos todos los segmentos de la polilínea para encontrar el más cercano
        For i As Integer = 0 To Polyline.NumberOfVertices - 2
            Dim pt1 As Point3d = Polyline.GetPoint3dAt(i)
            Dim pt2 As Point3d = Polyline.GetPoint3dAt(i + 1)
            bulge = Polyline.GetBulgeAt(i)

            ' Proyectar el punto en el segmento actual (primero en recto)
            Dim projectedPoint As Point3d = GetClosestPointOnSegment(pt1, pt2, PtStation)
            Dim distance As Double = PtStation.DistanceTo(projectedPoint)

            ' Si es el punto más cercano, actualizamos
            If distance < minDistance Then
                minDistance = distance
                closestPoint = projectedPoint
                closestSegmentIndex = i ' Guardamos el índice del segmento más cercano
                ' Identificamos si este segmento tiene bulge (curvo)
                If bulge <> 0 Then
                    isCurve = True
                Else
                    isCurve = False
                End If
            End If
        Next

        ' Si el segmento más cercano es curvo, corregimos la proyección usando el arco en ese segmento
        If isCurve AndAlso closestSegmentIndex >= 0 Then
            ' Obtener los puntos del segmento más cercano que es curvo
            Dim pt1 As Point3d = Polyline.GetPoint3dAt(closestSegmentIndex)
            Dim pt2 As Point3d = Polyline.GetPoint3dAt(closestSegmentIndex + 1)
            bulge = Polyline.GetBulgeAt(closestSegmentIndex)
            CGPointHelper.AddCGPoint(closestPoint.X, closestPoint.Y, closestPoint.Z, "Segmento recto")
            ' Proyectamos en el arco del segmento identificado
            Dim projectedPointArc As Point3d = ProjectPointOnPolyline(pt1, pt2, bulge, closestPoint)

            ' GetClosestPointOnArc(pt1, pt2, bulge, closestPoint)
            Dim distanceArc As Double = closestPoint.DistanceTo(projectedPointArc)

            ' Si el punto en el arco es más cercano, actualizamos
            If distanceArc < minDistance Then
                closestPoint = projectedPointArc
            End If
        End If

        '' Verificar si el punto más cercano está en la polilínea
        'If Not IsPointOnPolyline(Polyline, closestPoint) Then
        '    ' Si no está en la polilínea, devolvemos el punto más cercano con tolerancia
        '    If minDistance < Tolerance.Global.EqualPoint Then
        '        Return closestPoint
        '    Else
        '        ' Si no está dentro de la tolerancia, lanzar una excepción o manejarlo según sea necesario
        '        Throw New InvalidOperationException("No se pudo proyectar un punto en la polilínea.")
        '    End If
        'End If

        ' Devolvemos el punto proyectado más cercano
        Return closestPoint
    End Function

    ' Función auxiliar para proyectar el punto sobre un segmento
    Private Function GetClosestPointOnSegment(pt1 As Point3d, pt2 As Point3d, ptStation As Point3d) As Point3d
        ' Vector de la línea (pt1->pt2)
        Dim lineVec As Vector3d = pt1.GetVectorTo(pt2)
        Dim stationVec As Vector3d = pt1.GetVectorTo(ptStation)

        ' Proyección del vector stationVec sobre lineVec
        Dim projectionLength As Double = stationVec.DotProduct(lineVec) / lineVec.LengthSqrd

        ' Si la proyección está fuera del segmento, ajustamos los valores
        If projectionLength < 0 Then
            Return pt1
        ElseIf projectionLength > 1 Then
            Return pt2
        End If

        ' Obtener el punto proyectado en el segmento
        Dim projectedPoint As Point3d = pt1 + (lineVec * projectionLength)
        Return projectedPoint
    End Function

    ' Función auxiliar para proyectar el punto sobre un segmento curvo
    Private Function GetClosestPointOnArc(pt1 As Point3d, pt2 As Point3d, bulge As Double, ptStation As Point3d) As Point3d
        ' Obtener el punto medio entre pt1 y pt2 usando vectores
        Dim chordMidpoint As Point3d = pt1 + (pt2 - pt1) * 0.5
        Dim halfChordLength As Double = pt1.DistanceTo(chordMidpoint)
        Dim sagitta As Double = bulge * halfChordLength

        ' Encontrar el centro del círculo
        Dim chordVec As Vector3d = pt1.GetVectorTo(pt2)
        Dim normal As Vector3d = chordVec.CrossProduct(Vector3d.ZAxis).GetNormal() ' Perpendicular a la cuerda
        Dim circleCenter As Point3d = chordMidpoint + normal * sagitta

        ' Proyectar el punto ptStation sobre el arco
        Dim radius As Double = circleCenter.DistanceTo(pt1)
        Dim radialVec As Vector3d = circleCenter.GetVectorTo(ptStation).GetNormal() * radius

        ' El punto más cercano en el arco es el que está sobre el radio proyectado
        Dim projectedPoint As Point3d = circleCenter + radialVec
        Return projectedPoint
    End Function
    Private Function IsPointOnPolyline(pl As Polyline, pt As Point3d) As Boolean
        Dim isOn As Boolean = False

        ' Recorremos todos los segmentos de la polilínea
        For i As Integer = 0 To pl.NumberOfVertices - 2
            Dim seg As Curve3d = Nothing

            ' Obtener el tipo de segmento (arco o línea)
            Dim segType As SegmentType = pl.GetSegmentType(i)

            If segType = SegmentType.Arc Then
                seg = pl.GetArcSegmentAt(i) ' Segmento de arco
            ElseIf segType = SegmentType.Line Then
                seg = pl.GetLineSegmentAt(i) ' Segmento de línea
            End If

            ' Verificar si el punto está en el segmento
            If seg IsNot Nothing Then
                isOn = seg.IsOn(pt, Tolerance.Global)
                If isOn Then
                    Exit For ' Si el punto está en el segmento, salir del bucle
                End If
            End If
        Next
        Return isOn
    End Function




    ' Método principal que proyecta un punto sobre una polilínea, eligiendo el método correcto según las características del segmento
    Public Function ProjectPointOnPolyline(pt1 As Point3d, pt2 As Point3d, bulge As Double, ptStation As Point3d) As Point3d
        Dim chordLength As Double = pt1.DistanceTo(pt2)

        ' Determinar qué método usar en función del bulge y las características del segmento
        If Math.Abs(bulge) < Tolerance.Global.EqualPoint Then
            ' Caso 1: Segmento recto o bulge muy pequeño
            Return ProjectOnStraightSegment(pt1, pt2, ptStation)

        ElseIf Math.Abs(bulge) < 0.1 Then
            ' Caso 2: Bulge pequeño (curvatura mínima)
            Return ProjectOnSmallBulgeSegment(pt1, pt2, bulge, ptStation)

        ElseIf Math.Abs(bulge) >= 0.1 Then
            ' Caso 3: Bulge grande (curvatura pronunciada)
            Return ProjectOnLargeBulgeSegment(pt1, pt2, bulge, ptStation)

        ElseIf bulge < 0 Then
            ' Caso 4: Bulge negativo (curvatura en dirección opuesta)
            Return ProjectOnNegativeBulgeSegment(pt1, pt2, bulge, ptStation)

            'ElseIf IsPolylineClosed(pt1, pt2) Then
            '    ' Caso 5: Polilínea cerrada (como un bucle)
            '    Return ProjectOnClosedPolyline(pt1, pt2, bulge, ptStation)

            'ElseIf IsPolyline3D(pt1, pt2) Then
            '    ' Caso 6: Segmento 3D
            '    Return ProjectOn3DSegment(pt1, pt2, bulge, ptStation)

        Else
            ' Si ninguno de los casos aplica, devolver el punto proyectado en la cuerda
            Return ProjectOnStraightSegment(pt1, pt2, ptStation)
        End If
    End Function

    ' Método para proyectar un punto sobre un segmento recto o con bulge cercano a cero
    Private Function ProjectOnStraightSegment(pt1 As Point3d, pt2 As Point3d, ptStation As Point3d) As Point3d
        Return GetClosestPointOnSegment(pt1, pt2, ptStation)
    End Function

    ' Método para proyectar un punto sobre un segmento con un bulge pequeño
    Private Function ProjectOnSmallBulgeSegment(pt1 As Point3d, pt2 As Point3d, bulge As Double, ptStation As Point3d) As Point3d
        Dim projectedPointOnLine As Point3d = GetClosestPointOnSegment(pt1, pt2, ptStation)
        ' Calcular el ajuste en el arco, la curvatura es mínima
        Dim chordVec As Vector3d = pt1.GetVectorTo(pt2)
        Dim sagitta As Double = Math.Abs(bulge) * chordVec.Length / 2
        Dim normal As Vector3d = chordVec.CrossProduct(Vector3d.ZAxis).GetNormal()

        ' Ajustar la proyección para caer sobre el arco
        Return projectedPointOnLine + normal * sagitta
    End Function

    ' Método para proyectar un punto sobre un segmento con bulge grande
    Private Function ProjectOnLargeBulgeSegment(pt1 As Point3d, pt2 As Point3d, bulge As Double, ptStation As Point3d) As Point3d
        ' Proyección más precisa para arcos grandes usando trigonometría completa
        Dim chordVec As Vector3d = pt1.GetVectorTo(pt2)
        Dim chordLength As Double = chordVec.Length
        Dim midPoint As Point3d = pt1 + (chordVec * 0.5)

        ' Calcular la sagitta y proyectar usando el bulge
        Dim sagitta As Double = Math.Abs(bulge) * chordLength / 2
        Dim normal As Vector3d = chordVec.CrossProduct(Vector3d.ZAxis).GetNormal()
        Return midPoint + normal * sagitta
    End Function

    ' Método para proyectar un punto sobre un segmento con bulge negativo
    Private Function ProjectOnNegativeBulgeSegment(pt1 As Point3d, pt2 As Point3d, bulge As Double, ptStation As Point3d) As Point3d
        ' Proyectar como en el caso del bulge positivo, pero invertimos la dirección del normal
        Dim projectedPointOnLine As Point3d = GetClosestPointOnSegment(pt1, pt2, ptStation)
        Dim chordVec As Vector3d = pt1.GetVectorTo(pt2)
        Dim sagitta As Double = Math.Abs(bulge) * chordVec.Length / 2
        Dim normal As Vector3d = -chordVec.CrossProduct(Vector3d.ZAxis).GetNormal() ' Invertir normal

        Return projectedPointOnLine + normal * sagitta
    End Function

    ' Método para manejar polilíneas cerradas
    Private Function ProjectOnClosedPolyline(pt1 As Point3d, pt2 As Point3d, bulge As Double, ptStation As Point3d) As Point3d
        ' Proyectar el punto sobre un segmento cerrado o bucle
        ' Asegurarse de manejar correctamente las conexiones entre el primer y último punto
        Dim projectedPointOnLine As Point3d = GetClosestPointOnSegment(pt1, pt2, ptStation)
        ' Proyección normal, pero con lógica para manejar bucles si es necesario
        Return ProjectOnSmallBulgeSegment(pt1, pt2, bulge, ptStation)
    End Function

    ' Método para manejar polilíneas en 3D
    Private Function ProjectOn3DSegment(pt1 As Point3d, pt2 As Point3d, bulge As Double, ptStation As Point3d) As Point3d
        ' Convertir a un plano adecuado si es necesario, o realizar una proyección en 3D
        ' Aquí podrías necesitar proyectar el punto en el plano correcto antes de usar los cálculos
        ' Asumimos que la proyección en el plano Z es necesaria para polilíneas 3D
        Dim projectedPointOnLine As Point3d = GetClosestPointOnSegment(pt1, pt2, ptStation)
        ' Ajustar usando el bulge
        Return ProjectOnSmallBulgeSegment(pt1, pt2, bulge, ptStation)
    End Function


End Class
