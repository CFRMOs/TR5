Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports DBObject = Autodesk.AutoCAD.DatabaseServices.DBObject
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity

Public Module Parcelas
    Public ARTitulos As Array = {"Numero", "Area", "Nombre", "TabName", "MXstation", "Mnstation"}
    Public ARFT As Array = {"0.00", "0.00", "@", "@", "@", "0+000.00", "0+000.00"}

    <CommandMethod("CMDPSTEst")>
    Public Sub CMDPSTEst()
        ' Obtener la colección de parcelas del dibujo activo
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
            Try

                Dim ent As Entity
                Dim promptentOptions As New PromptEntityOptions("Seleccionar un Parcel:")
                ent = SelectEntityByType("Parcel", promptentOptions)

                Dim PolyL As Polyline = GetSegmentParcels(ent)
                ent = Nothing
                trans.Commit()

            Catch ex As Exception
                ed.WriteMessage(vbCrLf & "Error : " & ex.Message)
                trans.Abort()
            Finally
                trans.Dispose()
            End Try
        End Using
    End Sub
    <CommandMethod("CMDTestParcel")>
    Public Sub CMDTestParcel()
        ' Obtener la colección de parcelas del dibujo activo
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        'Dim xl As New ExcelAppTR
        'Dim ShNAme As String = "Reporte Parcelas Autocad"
        'Dim xlSh As Worksheet = xl.GetSH(ShNAme)
        'If xlSh Is Nothing Then Exit Sub

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
            Try
                Dim promptALingOptions As New PromptEntityOptions("Seleccionar un Alignment:")

                Dim acEntAlignment As Entity

                acEntAlignment = SelectEntityByType("Alignment", promptALingOptions)


                Dim AR As Array

                If acEntAlignment Is Nothing Then Exit Sub
                ' Iterar sobre las entidades del modelo
                For Each objId As ObjectId In btr
                    Dim ent As Entity = CType(trans.GetObject(objId, OpenMode.ForRead), Entity)
                    'ed.WriteMessage(vbCrLf & "Type" & TypeName(ent))
                    ' Verificar si la entidad es una parcela
                    If TypeName(ent) = "Parcel" Then
                        Dim parcel As Parcel = CType(ent, Parcel)
                        Dim PolyL As Polyline = GetSegmentParcels(ent)
                        Dim MXstation As Double, Mnstation As Double
                        CStationOffsetLabel.GetMxMnBorder(PolyL, CType(acEntAlignment, Alignment), MXstation, Mnstation)
                        'Dim Points As Point2dCollection = CollectPLPoints(EntPl)
                        'Dim entBOr As Entity = GetSegmentParcels(ent)
                        Dim HND As String = parcel.Handle.ToString()
                        '' Acceder a las propiedades de la parcela
                        Dim parcelNumber As String = parcel.Number
                        Dim area As Double = parcel.Area

                        '' Aquí puedes hacer lo que necesites con los datos de la parcela
                        AR = {parcelNumber, area, parcel.Name, HND, MXstation, Mnstation}
                        ''xl.FTtransferdata(AR, ARTitulos, ARFT)

                    End If
                    ent = Nothing
                Next
                trans.Commit()

            Catch ex As Exception
                ed.WriteMessage(vbCrLf & "Error : " & ex.Message)
                trans.Abort()
            Finally
                trans.Dispose()
            End Try
        End Using
    End Sub
    <CommandMethod("CMDParcel")>
    Public Sub CMDParcel()
        ' Obtener la colección de parcelas del dibujo activo
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        'Dim xl As New ExcelAppTR
        'Dim ShNAme As String = "Reporte Parcelas Autocad"
        'Dim xlSh As Worksheet = xl.GetSH(ShNAme)
        'If xlSh Is Nothing Then Exit Sub

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
            Try
                Dim promptALingOptions As New PromptEntityOptions("Seleccionar un Alignment:")

                Dim acEntAlignment As Entity

                acEntAlignment = SelectEntityByType("Alignment", promptALingOptions)
                Dim AR As Array

                If acEntAlignment Is Nothing Then Exit Sub
                ' Iterar sobre las entidades del modelo
                'For Each objId As ObjectId In btr
                Dim ent As Entity
                Dim promptentOptions As New PromptEntityOptions("Seleccionar un Parcel:")

                ent = SelectEntityByType("Parcel", promptentOptions)
                'ed.WriteMessage(vbCrLf & "Type" & TypeName(ent))
                ' Verificar si la entidad es una parcela
                If TypeName(ent) = "Parcel" Then
                    Dim parcel As Parcel = CType(ent, Parcel)
                    Dim PolyL As Polyline = GetSegmentParcels(ent)
                    Dim MXstation As Double, Mnstation As Double
                    CStationOffsetLabel.GetMxMnBorder(PolyL, CType(acEntAlignment, Alignment), MXstation, Mnstation)

                    Dim HND As String = parcel.Handle.ToString()
                    '' Acceder a las propiedades de la parcela
                    Dim parcelNumber As String = parcel.Number
                    Dim area As Double = parcel.Area

                    '' Aquí puedes hacer lo que necesites con los datos de la parcela
                    AR = {parcelNumber, area, parcel.Name, HND, MXstation, Mnstation}
                    'xl.FTtransferdata(AR, ARTitulos, ARFT)

                End If
                ent = Nothing
                'Next
                trans.Commit()

            Catch ex As Exception
                ed.WriteMessage(vbCrLf & "Error : " & ex.Message)
                trans.Abort()
            Finally
                trans.Dispose()
            End Try
        End Using
    End Sub
    <CommandMethod("CMDPolyline")>
    Public Sub CMDPolyline()
        ' Obtener la colección de parcelas del dibujo activo
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        'Dim xl As New ExcelAppTR
        'Dim ShNAme As String = "Reporte Parcelas Autocad"
        'Dim xlSh As Worksheet = xl.GetSH(ShNAme)
        'If xlSh Is Nothing Then Exit Sub

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
            Try
                Dim promptALingOptions As New PromptEntityOptions("Seleccionar un Alignment:")

                Dim acEntAlignment As Entity
                acEntAlignment = SelectEntityByType("Alignment", promptALingOptions)
                Dim AR As Array
                If acEntAlignment Is Nothing Then Exit Sub
                ' Iterar sobre las entidades del modelo
                'For Each objId As ObjectId In btr
                Dim ent As Entity
                Dim promptentOptions As New PromptEntityOptions("Seleccionar un Parcel:")

                ent = SelectEntityByType("Polyline", promptentOptions)
                'ed.WriteMessage(vbCrLf & "Type" & TypeName(ent))
                ' Verificar si la entidad es una parcela
                If TypeName(ent) = "Polyline" Then

                    Dim PolyL As Polyline = CType(ent, Polyline)
                    Dim MXstation As Double, Mnstation As Double
                    CStationOffsetLabel.GetMxMnBorder(PolyL, CType(acEntAlignment, Alignment), MXstation, Mnstation)

                    Dim HND As String = PolyL.Handle.ToString()

                    Dim area As Double = PolyL.Area

                    '' Aquí puedes hacer lo que necesites con los datos de la parcela
                    AR = {0, area, "poli", HND, MXstation, Mnstation}
                    'xl.FTtransferdata(AR, ARTitulos, ARFT)

                End If
                ent = Nothing
                'Next
                trans.Commit()

            Catch ex As Exception
                ed.WriteMessage(vbCrLf & "Error : " & ex.Message)
                trans.Abort()
            Finally
                trans.Dispose()
            End Try
        End Using
    End Sub

    'Public Sub TransferirAExcel(AR As Array, ARTitulos As Array, ARFT As Array)
    '    'Dim xl As New ExcelAppTR
    '    'Dim ShNAme As String = "Reporte Parcelas Autocad"
    '    'Dim xlSh As Worksheet = xl.GetSH(ShNAme)
    '    'If xlSh Is Nothing Then Exit Sub
    '    'Dim AR As Array = {HandlePL.Value, StPtStation, EDPtStation, Len(), EDStation_OffSet_Side, HandleAlign.Value}
    '    'Dim ARTitulos As Array = {"TabName", "PK-Inicial", "PK-Final", "Longitud", "Lado", "AlignHDL"}
    '    'Dim ARFT As Array = {"@", "0+000.00", "0+000.00", "0.00", "@", "@"}

    '    'xl.FTtransferdata(AR, ARTitulos, ARFT)
    'End Sub
    Public Function GetSegmentParcels(ent As Entity) As Polyline
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim parcel As Parcel = CType(ent, Parcel)
                Dim objs As New DBObjectCollection()

                ent.Explode(objs)
                For Each Obj As DBObject In objs
                    If TypeName(Obj) = "Polyline" Then
                        Return CType(Obj, Polyline)
                    End If
                Next

            Catch ex As Exception
                ed.WriteMessage(vbCrLf & "Error : " & ex.Message)
                trans.Abort()
            Finally
                trans.Dispose()
            End Try

        End Using

        ' Return Nothing if no polyline is found
        Return Nothing
    End Function

End Module
