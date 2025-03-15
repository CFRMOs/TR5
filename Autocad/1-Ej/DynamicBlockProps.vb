Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Namespace DynamicBlocks
    Public Class Commands
        <CommandMethod("DBP")>
        Public Shared Sub DynamicBlockProps()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Dim pso As New PromptStringOptions(
                vbLf & "Enter dynamic block name or enter to select: ") With {
                .AllowSpaces = True
                }
            Dim pr As PromptResult = ed.GetString(pso)

            If pr.Status <> PromptStatus.OK Then
                Return
            End If

            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim br As BlockReference = Nothing

                ' If a null string was entered allow entity selection
                If pr.StringResult = "" Then
                    ' Select a block reference
                    Dim peo As New PromptEntityOptions(
                        vbLf & "Select dynamic block reference: ")

                    peo.SetRejectMessage(vbLf & "Entity is not a block.")
                    peo.AddAllowedClass(GetType(BlockReference), False)

                    Dim per As PromptEntityResult = ed.GetEntity(peo)

                    If per.Status <> PromptStatus.OK Then
                        Return
                    End If

                    ' Access the selected block reference
                    br = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead), BlockReference)
                Else
                    ' Otherwise we look up the block by name
                    Dim bt As BlockTable = TryCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                    If Not bt.Has(pr.StringResult) Then
                        ed.WriteMessage(vbLf & "Block """ & pr.StringResult & """ does not exist.")
                        Return
                    End If

                    ' Create a new block reference referring to the block
                    br = New BlockReference(New Point3d(), bt(pr.StringResult))
                End If

                Dim btr As BlockTableRecord = CType(tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)

                ' Call our function to display the block properties
                DisplayDynBlockProperties(ed, br, btr.Name)

                ' Committing is cheaper than aborting
                tr.Commit()
            End Using
        End Sub

        Private Shared Sub DisplayDynBlockProperties(ByVal ed As Editor, ByVal br As BlockReference, ByVal name As String)
            ' Only continue is we have a valid dynamic block
            If br IsNot Nothing AndAlso br.IsDynamicBlock Then
                ed.WriteMessage(vbLf & "Dynamic properties for ""{0}""", name)

                ' Get the dynamic block's property collection
                Dim pc As DynamicBlockReferencePropertyCollection = br.DynamicBlockReferencePropertyCollection

                ' Loop through, getting the info for each property
                For Each prop As DynamicBlockReferenceProperty In pc
                    ' Start with the property name, type and description
                    ed.WriteMessage(vbLf & "Property: ""{0}"" : {1}", prop.PropertyName, prop.UnitsType)

                    If prop.Description <> "" Then
                        ed.WriteMessage(vbLf & "  Description: {0}", prop.Description)
                    End If

                    ' Is it read-only?
                    If prop.ReadOnly Then
                        ed.WriteMessage(" (Read Only)")
                    End If

                    ' Get the allowed values, if it's constrained
                    Dim first As Boolean = True

                    For Each value As Object In prop.GetAllowedValues()
                        ed.WriteMessage(If(first, vbLf & "  Allowed values: [", ", "))
                        ed.WriteMessage("""{0}""", value)
                        first = False
                    Next

                    If Not first Then
                        ed.WriteMessage("]")
                    End If

                    ' And finally the current value
                    ed.WriteMessage(vbLf & "  Current value: ""{0}""", prop.Value)
                Next
            End If
        End Sub
    End Class
End Namespace
