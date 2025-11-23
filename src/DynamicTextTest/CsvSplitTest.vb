Imports System
Imports Xunit
Imports ZoppaLibrary
Imports ZoppaLibrary.Collections
Imports ZoppaLibrary.LegacyFiles
Imports ZoppaLibrary.Strings

Public Class CsvSplitTest

    <Fact>
    Public Sub Split_CsvLine_ShouldReturnCorrectParts()
        Dim csvLine As String = "Name,Age,Location
造田, 49, 福山
あいり, 20, 広島
"
        Dim splitter As CsvSplitter = CsvSplitter.CreateSplitter(csvLine)
        Dim parts1 = splitter.Split()
        Assert.Equal("造田", parts1("Name").ToString())
        Assert.Equal("49", parts1("Age").ToString())
        Assert.Equal("福山", parts1("Location").ToString())

        Dim parts2 = splitter.Split()
        Assert.Equal("あいり", parts2("Name").ToString())
        Assert.Equal("20", parts2("Age").ToString())
        Assert.Equal("広島", parts2("Location").ToString())

        Dim parts3 = splitter.Split()
        Dim parts4 = splitter.Split()
    End Sub

End Class
