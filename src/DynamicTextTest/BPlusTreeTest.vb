Imports System
Imports System.Globalization
Imports System.Security.AccessControl
Imports Xunit
Imports ZoppaLibrary
Imports ZoppaLibrary.Collections

' dotnet test --collect:"XPlat Code Coverage" 
' ReportGenerator -reports:"G:\source\ZoppaDLogger\ZoppaDLoggerTest\TestResults\c6403942-e7bb-437e-bf4a-b7f5942b6228\coverage.cobertura.xml" -targetdir:"coveragereport" -reporttypes:Html


Public Class BPlusTreeTest

    <Fact>
    Public Sub Insert_MultipleValues()
        Dim btree As New BPlusTree(Of Integer)(2)
        For Each value In {
            32, 38, 34, 13, 11, 5, 12, 8, 25, 27, 26, 30, 1, 21, 22, 2, 3, 37, 23, 6,
            36, 17, 10, 9, 14, 4, 33, 40, 15, 24, 35, 29, 31, 39, 16, 28, 19, 20, 7, 18
        }
            btree.Insert(value)
        Next

        Dim count As Integer = 1
        For Each v In btree
            Assert.Equal(count, v)
            count += 1
        Next

        For i As Integer = 1 To 40
            Assert.True(btree.Contains(i), $"BPlusTree {i} が見つからない")
        Next
        Assert.False(btree.Contains(0), $"BPlusTree 0 が見つかった")
    End Sub

    <Fact>
    Public Sub Delete_MultipleValues()
        Dim btree As New BPlusTree(Of Integer)(2)
        For i As Integer = 1 To 20
            btree.Insert(i)
        Next

        For Each value In {
            18, 1, 7, 8, 5, 9, 19, 6, 16, 15,
            20, 11, 13, 14, 4, 10, 2, 12, 3, 17
        }
            btree.Remove(value)
        Next
    End Sub

    <Fact>
    Public Sub Delete_MultipleValues2()
        Dim btree As New BPlusTree(Of Integer)(4)
        For i As Integer = 1 To 40
            btree.Insert(i)
        Next

        For Each value In {
            32, 38, 34, 13, 11, 5, 12, 8, 25, 27,
            26, 30, 1, 21, 22, 2, 3, 37, 23, 6,
            36, 17, 10, 9, 14, 4, 33, 40, 15, 24,
            35, 29, 31, 39, 16, 28, 19, 20, 7, 18
        }
            btree.Remove(value)
        Next
    End Sub

    <Fact>
    Public Sub Delete_MultipleValues3()
        Dim datas = {
            4, 17, 13, 2, 7, 6, 12, 8, 1, 6,
            10, 11, 19, 14, 18, 2, 17, 2, 3, 5
        }

        Dim btree As New BPlusTree(Of Integer)(2)

        For Each value In datas
            btree.Insert(value)
        Next

        For Each value In datas
            btree.Remove(value)
        Next
    End Sub

    <Fact>
    Public Sub Delete_MultipleValues4()
        Dim datas = {
            81, 49, 94, 40, 62, 33, 50, 29, 4, 25, 58, 27, 93, 55, 9, 12, 83, 56, 88, 3,
            77, 89, 75, 22, 61, 86, 16, 19, 34, 2, 63, 47, 24, 45, 11, 30, 15, 65, 26, 20,
            71, 31, 95, 38, 90, 53, 17, 36, 41, 82, 100, 68, 21, 37, 60, 76, 32, 54, 74, 92,
            43, 8, 85, 14, 6, 72, 97, 1, 39, 18, 10, 78, 79, 57, 28, 73, 64, 46, 69, 99,
            23, 35, 44, 70, 52, 91, 84, 7, 67, 48, 98, 13, 42, 87, 96, 5, 66, 59, 80, 51
        }

        Dim btree As New BPlusTree(Of Integer)(6)

        For Each value In datas
            btree.Insert(value)
        Next

        For Each value In datas
            btree.Remove(value)
        Next
    End Sub

    <Fact>
    Public Sub Search_ExistingValue_ReturnsValue()
        Dim tree As New BPlusTree(Of Integer)()
        tree.Insert(10)
        tree.Insert(20)
        tree.Insert(30)
        Assert.Equal(10, tree.Search(10))
        Assert.Equal(20, tree.Search(20))
        Assert.Equal(30, tree.Search(30))
    End Sub

    <Fact>
    Public Sub Search_NonExistingValue_ReturnsNothing()
        Dim tree As New BPlusTree(Of Integer)()
        tree.Insert(10)
        tree.Insert(20)
        Assert.Equal(0, tree.Search(99)) ' IntegerのNothingは0
    End Sub

    <Fact>
    Public Sub Search_EmptyTree_ReturnsNothing()
        Dim tree As New BPlusTree(Of String)()
        Assert.Null(tree.Search("abc"))
    End Sub

    <Fact>
    Public Sub Search_DuplicateInsert_ReturnsFirstInserted()
        Dim tree As New BPlusTree(Of String)()
        tree.Insert("abc")
        tree.Insert("def")
        tree.Insert("abc")
        Assert.Equal("abc", tree.Search("abc"))
    End Sub

    <Fact>
    Public Sub Search_Null_ThrowsArgumentNullException()
        Dim tree As New BPlusTree(Of String)()
        Assert.Equal(Nothing, tree.Search(Nothing))
    End Sub

    Private Class MyEntry
        Implements IComparable(Of MyEntry)

        Public Property Key As Integer
        Public Property UniqueKey As Integer

        Public Function CompareTo(other As MyEntry) As Integer Implements IComparable(Of MyEntry).CompareTo
            Return Me.Key.CompareTo(other.Key)
        End Function
    End Class

    <Fact>
    Public Sub Insert_EntryWithUniqueKey()
        Dim btree As New BPlusTree(Of MyEntry)(2)
        Dim datas = New MyEntry(99) {}
        For i As Integer = 0 To 99
            datas(i) = New MyEntry With {.Key = i \ 5 + 1, .UniqueKey = i + 1}
        Next
        Dim rnd As New Random()
        For i As Integer = 0 To 98
            Dim index = rnd.Next(i, 100)
            Dim temp = datas(i)
            datas(i) = datas(index)
            datas(index) = temp
        Next
        For i As Integer = 0 To 99
            Dim imax = 1
            If btree.Contains(datas(i)) Then
                imax = btree.Search(datas(i)).UniqueKey + 1
            End If
            datas(i).UniqueKey = imax
            btree.Insert(datas(i))
        Next

        For Each entry In btree
            Debug.WriteLine($"Key: {entry.Key}, UniqueKey: {entry.UniqueKey}")
        Next

        For i As Integer = 1 To 20
            btree.Remove(New MyEntry With {.Key = i})
        Next
        For Each entry In btree
            Debug.WriteLine($"Key: {entry.Key}, UniqueKey: {entry.UniqueKey}")
        Next
    End Sub

End Class
