Imports System
Imports Xunit
Imports ZoppaLibrary
Imports ZoppaLibrary.Collections

' dotnet test --collect:"XPlat Code Coverage" 
' ReportGenerator -reports:"G:\source\ZoppaDLogger\ZoppaDLoggerTest\TestResults\c6403942-e7bb-437e-bf4a-b7f5942b6228\coverage.cobertura.xml" -targetdir:"coveragereport" -reporttypes:Html

Public Class BtreeTest


    <Fact>
    Public Sub Insert_SingleValue_ShouldContainValue()
        Dim btree As New Btree(Of Integer)()
        btree.Insert(10)
    End Sub

    <Fact>
    Public Sub Insert_MultipleValues_ShouldNotThrow()
        Dim btree As New Btree(Of Integer)()
        For Each value In {32, 38, 34, 13, 11, 5, 12, 8, 25, 27, 26, 30, 1, 21, 22, 2, 3, 37, 23, 6, 36, 17, 10, 9, 14, 4, 33, 40, 15, 24, 35, 29, 31, 39, 16, 28, 19, 20, 7, 18}
            btree.Insert(value)
        Next
        Dim count As Integer = 1
        For Each v In btree
            Assert.Equal(count, v)
            count += 1
        Next
    End Sub

    <Fact>
    Public Sub Insert_DuplicateValue_ShouldThrowException()
        Dim btree As New Btree(Of Integer)()
        btree.Insert(7)
        Assert.Throws(Of BtreeException)(
            Sub()
                btree.Insert(7) ' ここで例外が発生するはず
            End Sub)
    End Sub

    <Fact>
    Public Sub Insert_NullValue_ShouldThrowException()
        Dim btree As New Btree(Of String)()
        Assert.Throws(Of ArgumentNullException)(
            Sub()
                btree.Insert(Nothing) ' ここで例外が発生するはず
            End Sub)
    End Sub

    <Fact>
    Public Sub Remove_ExistingValue_ShouldNotContainValue()
        Dim btree As New Btree(Of Integer)()
        btree.Insert(10)
        btree.Insert(20)
        btree.Insert(30)
        btree.Remove(20)
        For Each v In btree
            Assert.NotEqual(20, v)
        Next
    End Sub

    <Fact>
    Public Sub Remove_NullValue_ShouldThrowException()
        Dim btree As New Btree(Of String)()
        btree.Insert(1)
        Assert.Throws(Of ArgumentNullException)(
            Sub()
                btree.Remove(Nothing)
            End Sub)
    End Sub

    <Fact>
    Public Sub Remove_NonExistingValue_ShouldNotThrow()
        Dim btree As New Btree(Of Integer)()
        btree.Insert(10)
        Assert.Throws(Of BtreeException)(
            Sub()
                ' 存在しない値を削除しようとしても例外は発生しない
                btree.Remove(20)
            End Sub)
    End Sub

    <Fact>
    Public Sub Remove_AllValues_ShouldBeEmpty()
        Dim btree As New Btree(Of Integer)()
        For Each value In {
            7, 1, 87, 42, 93, 88, 34, 62, 35, 74, 69, 67, 91, 28, 32, 38, 68, 6, 20, 46,
            63, 17, 52, 58, 70, 81, 85, 61, 15, 48, 57, 19, 56, 13, 40, 84, 8, 71, 16, 64,
            94, 89, 53, 47, 25, 49, 23, 5, 33, 75, 45, 100, 27, 73, 9, 39, 37, 2, 18, 83,
            41, 59, 22, 72, 3, 65, 60, 26, 77, 50, 51, 80, 24, 54, 96, 86, 31, 29, 76, 97,
            90, 10, 95, 92, 36, 44, 43, 78, 66, 21, 30, 12, 14, 11, 55, 82, 79, 98, 4, 99
        }
            btree.Insert(value)
        Next

        Assert.True(btree.Contains(88))
        Assert.True(btree.Contains(17))
        Assert.True(btree.Contains(55))
        Assert.False(btree.Contains(0))
        Assert.False(btree.Contains(101))

        For i As Integer = 1 To 100
            btree.Remove(i)
        Next
        Assert.Equal(0, btree.Count)
    End Sub

    <Fact>
    Public Sub Remove_AllValues_ShouldBeEmpty2()
        Dim btree As New Btree(Of Integer)()
        For Each value In {
            89, 43, 66, 129, 116, 44, 102, 99, 125, 130, 83, 136, 126, 96, 58, 91, 114, 54, 94, 25,
            34, 68, 145, 120, 67, 16, 138, 78, 12, 38, 141, 19, 11, 8, 143, 90, 39, 131, 80, 112,
            69, 142, 157, 110, 128, 13, 133, 45, 5, 49, 81, 23, 47, 32, 92, 159, 155, 144, 95, 104,
            22, 27, 15, 148, 28, 140, 109, 30, 151, 98, 134, 122, 7, 106, 17, 18, 52, 71, 103, 46,
            60, 41, 63, 111, 160, 119, 156, 56, 74, 59, 127, 75, 65, 4, 35, 20, 121, 14, 79, 61,
            132, 76, 10, 70, 139, 137, 40, 149, 1, 2, 3, 50, 55, 64, 57, 73, 117, 72, 29, 152,
            62, 84, 31, 105, 82, 33, 48, 26, 53, 51, 100, 118, 87, 24, 101, 97, 146, 36, 42, 154,
            150, 88, 108, 124, 135, 21, 9, 113, 37, 85, 153, 6, 158, 107, 123, 77, 115, 147, 93, 86
        }
            btree.Insert(value)
        Next
        Assert.Equal(btree.Search(88), 88)
        Assert.Equal(btree.Search(17), 17)
        Assert.Equal(btree.Search(55), 55)
        For i As Integer = 160 To 1 Step -1
            btree.Remove(i)
        Next

        Assert.Equal(0, btree.Count)

        Assert.Equal(btree.Search(55), Nothing)
    End Sub

    <Fact>
    Public Sub Remove_AllValues_ShouldBeEmpty3()
        Dim btree As New Btree(Of Integer)()
        For i As Integer = 1 To 24
            btree.Insert(i)
        Next
        For i As Integer = 1 To 4
            btree.Remove(i)
        Next
        Assert.Equal(20, btree.Count)
        For i As Integer = 5 To 24
            Assert.True(btree.Contains(i))
        Next
    End Sub

    <Fact>
    Public Sub Remove_AllValues_ShouldBeEmpty4()
        Dim btree As New Btree(Of Integer)()
        For i As Integer = 1 To 100
            btree.Insert(i)
        Next
        For Each v In {5, 10, 15, 20}
            btree.Remove(v)
        Next
        Assert.Equal(96, btree.Count)
    End Sub

    <Fact>
    Public Sub Clear_EmptiesTree()
        Dim tree As New Btree(Of Integer)()
        ' いくつか挿入
        tree.Insert(10)
        tree.Insert(20)
        tree.Insert(30)
        Assert.True(tree.Contains(10))
        Assert.True(tree.Contains(20))
        Assert.True(tree.Contains(30))

        ' クリア
        tree.Clear()

        ' すべて消えていること
        Assert.False(tree.Contains(10))
        Assert.False(tree.Contains(20))
        Assert.False(tree.Contains(30))

        ' 追加できること
        tree.Insert(100)
        Assert.True(tree.Contains(100))
    End Sub

    <Fact>
    Public Sub Clear_OnEmptyTree_DoesNotThrow()
        Dim tree As New Btree(Of String)()
        tree.Clear()
        Assert.False(tree.Contains("abc"))
        ' クリア後も追加できる
        tree.Insert("abc")
        Assert.True(tree.Contains("abc"))
    End Sub

    <Fact>
    Public Sub Clear_MultipleTimes()
        Dim tree As New Btree(Of Integer)()
        tree.Insert(1)
        tree.Clear()
        tree.Clear()
        Assert.False(tree.Contains(1))
        tree.Insert(2)
        Assert.True(tree.Contains(2))
        tree.Clear()
        Assert.False(tree.Contains(2))
    End Sub

    <Fact>
    Public Sub Remove_AllValues_ShouldBeEmpty5()
        Dim btree As New Btree(Of Integer)(2)
        For i As Integer = 1 To 40
            btree.Insert(i)
        Next
        For Each v In {
            8, 16, 35, 18, 17, 23, 36, 40, 19, 24, 31, 38, 25, 10, 15, 30, 21, 9, 14, 6,
            33, 5, 34, 4, 28, 2, 39, 11, 12, 20, 1, 32, 13, 37, 22, 3, 26, 7, 29, 27
        }
            btree.Remove(v)
        Next
        Assert.Equal(0, btree.Count)
    End Sub

End Class

