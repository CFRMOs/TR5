''Imports System
''Imports Autodesk.AutoCAD.ApplicationServices
''Imports Autodesk.AutoCAD.EditorInput
''Imports Autodesk.AutoCAD.Runtime
''Imports Autodesk.AECC.Interop.Land
''Imports Autodesk.AECC.Interop.UiLand

'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
'Imports Autodesk.AutoCAD.Runtime

'Namespace ParcelBoundary
'    Public Class ParcelBoundaryApp
'        Private m_oAcadApp As Autodesk.AutoCAD.Interop.AcadApplication = Nothing
'        Private m_oAeccApp As Autodesk.AECC.Interop.UiLand.AeccApplication = Nothing
'        Private m_oAeccDoc As Autodesk.AECC.Interop.UiLand.AeccDocument = Nothing
'        Private m_oAeccDb As IAeccDatabase = Nothing

'        Private m_sAcadProdID As String = "AutoCAD.Application"
'        Private m_sAeccAppProgId As String = "AeccXUiLand.AeccApplication"

'        Private m_sMessage As String = ""

'        <CommandMethod("CREATEBOUNDARY")> Public Sub CreateBoundary()
'            '            'Start Civil-3D, or get it if it's already running
'            Try
'                '                m_oAcadApp = CType(System.Runtime.InteropServices.Marshal.GetActiveObject(m_sAcadProdID), Autodesk.AutoCAD.Interop.AcadApplication)
'            Catch ex As System.Exception
'                '                Dim AcadProg As Type = System.Type.GetTypeFromProgID(m_sAcadProdID)
'                '                m_oAcadApp = CType(System.Activator.CreateInstance(AcadProg, True), Autodesk.AutoCAD.Interop.AcadApplication)
'            End Try

'            If m_oAcadApp IsNot Nothing Then
'                m_oAcadApp.Visible = True
'                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
'                Dim acCurDb As Database = doc.Database
'                'm_oAeccApp = CType(m_oAcadApp.GetInterfaceObject(m_sAeccAppProgId), IAeccApplication)
'                'm_oAeccDoc = CType(m_oAeccApp.ActiveDocument, IAeccDocument)

'                '                ' get the Database object via a late bind
'                '                m_oAeccDb = CType(m_oAeccDoc.GetType().GetProperty("Database").GetValue(m_oAeccDoc, Nothing), IAeccDatabase)

'                '                ' Get the first parcel for demonstration
'                Dim oParcel As IAeccParcel = acCurDb
'                '                = m_oAeccDb.Sites.Item(0).Parcels.Item(0)

'                '                ' Loop through all elements used to make parcel "oParcel"
'                '                Dim i As Integer
'                '                For i = 0 To oParcel.ParcelLoops.Count - 1
'                '                    Dim oElement As IAeccParcelSegmentElement
'                '                    oElement = oParcel.ParcelLoops.Item(i)

'                '                    m_sMessage += "Element " & i & " of segment " & oElement.ParcelSegment.Name & ": " & oElement.StartX & "," & oElement.StartY & " to " & oElement.EndX & ", " & oElement.EndY & vbCrLf

'                '                    If TypeOf oElement Is IAeccParcelSegmentLine Then
'                '                        Dim oSegmentLine As IAeccParcelSegmentLine = CType(oElement, IAeccParcelSegmentLine)
'                '                        m_sMessage += " is a line." & vbCrLf
'                '                    ElseIf TypeOf oElement Is IAeccParcelSegmentCurve Then
'                '                        Dim oSegmentCurve As IAeccParcelSegmentCurve = CType(oElement, IAeccParcelSegmentCurve)
'                '                        m_sMessage += " is a curve with a radius of:" & oSegmentCurve.Radius & vbCrLf
'                '                    End If
'                '                Next

'                '                Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
'                '                ed.WriteMessage(m_sMessage)
'                '                m_sMessage = ""
'            End If
'        End Sub
'    End Class
'End Namespace
