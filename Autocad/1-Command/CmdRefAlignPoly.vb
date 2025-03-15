Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
'Imports ExcelInteropManager
Public Class CmdRefAlignPoly

    <CommandMethod("CmdRefAlignPoly")>
    Public Sub GetEnt()
        '' Get the current document and database

        Dim acDoc As Document = GetDocumentManager().MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = acDoc.Editor
        Dim selectedAlignment As ObjectId = ObjectId.Null ' Initialize selected alignment variable
        ' get worksh
        'Dim xl As New ExcelAppTR
        'Dim xlSh As Worksheet = xl.GetSH()
        'If xlSh Is Nothing Then Exit Sub
        '' Request for objects to be selected in the drawing area
        'Dim selector As New CSelectorEntity
        'comprobacion de main handle of Alignment object
        '___draw some jammes dummies 
        '___

        Dim promptOptions As New PromptEntityOptions("Seleccionar un Polyline:")
        Dim acEnt As Entity = SelectEntityByType("Polyline", promptOptions)
        If acEnt Is Nothing Then Exit Sub
        Dim HandlePL As Handle = acEnt.Handle


        Dim promptALingOptions As New PromptEntityOptions("Seleccionar un Alignment:")
        Dim acEntAlignment As Entity = SelectEntityByType("Alignment", promptALingOptions)
        If acEntAlignment Is Nothing Then Exit Sub
        Dim HandleAlign As Handle = acEntAlignment.Handle

        '*---------------
        '  alternativa  -
        '*---------------

        Dim StartPT As Point3d
        Dim EndPT As Point3d
        Dim Len As Double

        'considerando varios tipos de polyline 
        If TypeName(acEnt) = "Polyline" Then
            Dim PL0 As Polyline = acEnt
            StartPT = PL0.StartPoint
            EndPT = PL0.EndPoint
            Len = PL0.Length
        ElseIf TypeName(acEnt) = "Polyline2d" Then
            Dim PL1 As Polyline2d = acEnt
            StartPT = PL1.StartPoint
            EndPT = PL1.EndPoint
            Len = PL1.Length
        End If

        Dim StPtStation As Double
        Dim StPGL_Elev As Double
        Dim StStatOffSet As Double
        Dim STStation_OffSet_Side As String = vbNull

        CStationOffsetLabel.GETStationByPoint(acEntAlignment, StartPT, StPtStation, StPGL_Elev, StStatOffSet, STStation_OffSet_Side)

        Dim EDPtStation As Double
        Dim EDPGL_Elev As Double
        Dim EDStatOffSet As Double
        Dim EDStation_OffSet_Side As String = vbNull

        CStationOffsetLabel.GETStationByPoint(acEntAlignment, EndPT, EDPtStation, EDPGL_Elev, EDStatOffSet, EDStation_OffSet_Side)

        Ed.WriteMessage(vbCrLf)
        WriteStation(StPtStation, Ed)
        WriteStation(EDPtStation, Ed)
        Ed.WriteMessage(Len & vbCrLf)

        Dim AR As Array = {HandlePL.Value, StPtStation, EDPtStation, Len, EDStation_OffSet_Side, HandleAlign.Value}
        Dim ARTitulos As Array = {"TabName", "PK-Inicial", "PK-Final", "Longitud", "Lado", "AlignHDL"}
        Dim ARFT As Array = {"@", "0+000.00", "0+000.00", "0.00", "@", "@"}

        'xl.FTtransferdata(AR, ARTitulos, ARFT)

    End Sub
    Public Function GetSingleEntity(BasePrompt As String, Type As Array) As List(Of Entity)
        '' Request for objects to be selected in the drawing area
        Dim ListEnt As List(Of Entity) = Nothing
        For i = LBound(Type) To UBound(Type)
            Dim promptOptions As New PromptEntityOptions(BasePrompt & " " & Type(i) & " :") '"Seleccionar un Polyline:")
            Dim acEnt As Entity = SelectEntityByType(Type(i).ToString, promptOptions) '"Polyline", promptOptions)
            ListEnt.Add(acEnt)
        Next
        Return ListEnt
    End Function
    Private Sub WriteStation(valor As Double, Ed As Editor)
        Dim valorFormateado As String = valor.ToString("0+000.00")
        Ed.WriteMessage(valorFormateado & vbCrLf)
    End Sub
End Class


