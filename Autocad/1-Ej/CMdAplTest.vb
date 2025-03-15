Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices

Public Module CMdAplTest
    <CommandMethod("CMdAplTest")>
    Public Sub CMdAplTest()
        Dim civDoc = CivilApplication.ActiveDocument
        Using tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Dim corr As Corridor = tr.GetObject(civDoc.CorridorCollection(0), OpenMode.ForRead)
            For Each bl As Baseline In corr.Baselines
                For Each station In bl.SortedStations()
                    Dim appliedassy = bl.GetAppliedAssemblyAtStation(station)
                    Dim pts As CalculatedPointCollection = appliedassy.Points 'this returns the CalulatedPointCollection For ALL points
                    Dim ptsbycode As CalculatedPointCollection = appliedassy.GetPointsByCode("EPS") 'whereas this gets the points for the single code "EPS"
                Next
            Next
        End Using
    End Sub

End Module
