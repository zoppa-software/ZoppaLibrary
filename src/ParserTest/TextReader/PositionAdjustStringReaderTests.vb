Imports System
Imports System.IO
Imports ZoppaLibrary.Parser
Imports Xunit


Public Class PositionAdjustStringReaderTests

    <Fact>
    Public Sub Peek_DoesNotAdvanceAndReturnsNextChar()
        Using tr = New PositionAdjustStringReader("abc")
            Dim p = tr.Peek()
            Assert.Equal(AscW("a"c), p)
            Assert.Equal(0, tr.Position)
            Dim r = tr.Read()
            Assert.Equal(AscW("a"c), r)
            Assert.Equal(1, tr.Position)
        End Using
    End Sub

    <Fact>
    Public Sub Read_SequentialReadsAndEof()
        Using tr = New PositionAdjustStringReader("ab")
            Assert.Equal(AscW("a"c), tr.Read())
            Assert.Equal(AscW("b"c), tr.Read())
            Assert.Equal(-1, tr.Read()) ' EOF
        End Using
    End Sub

    <Fact>
    Public Sub Read_BufferFillsAndReturnsCount()
        Using tr = New PositionAdjustStringReader("hello")
            Dim buf As Char() = New Char(9) {} ' length 10
            Dim readCount = tr.Read(buf, 2, 3)
            Assert.Equal(3, readCount)
            Assert.Equal("h"c, buf(2))
            Assert.Equal("e"c, buf(3))
            Assert.Equal("l"c, buf(4))
        End Using
    End Sub

    <Fact>
    Public Sub ReadChar_ReturnsNothingAtEof()
        Using tr = New PositionAdjustStringReader("")
            Dim c = tr.ReadChar()
            Assert.False(c.HasValue)
        End Using
    End Sub

    <Fact>
    Public Sub Snapshot_RestorePositionAndReadFromBuffer()
        Using tr = New PositionAdjustStringReader("xyz")
            Assert.Equal(AscW("x"c), tr.Read())
            Assert.Equal(AscW("y"c), tr.Read())
            Dim snap = tr.MemoryPosition()
            Assert.Equal(AscW("z"c), tr.Read()) ' consume z
            snap.Restore()
            Assert.Equal(2, tr.Position)
            Dim nc1 = tr.Peek()
            Assert.Equal(AscW("z"c), nc1)
            Assert.Equal(2, tr.Position)
            Dim nc2 = tr.Read()
            Assert.Equal(AscW("z"c), nc2)
            Assert.Equal(3, tr.Position)
        End Using
    End Sub

    <Fact>
    Public Sub ReadChar_ReadsAllAndReturnsNothingAtEnd()
        Dim tr = New PositionAdjustStringReader("ABC")

        Dim c1 = tr.ReadChar()
        Assert.True(c1.HasValue)
        Assert.Equal("A"c, c1.Value)
        Assert.Equal(1, tr.Position)

        Dim c2 = tr.ReadChar()
        Assert.True(c2.HasValue)
        Assert.Equal("B"c, c2.Value)
        Assert.Equal(2, tr.Position)

        Dim c3 = tr.ReadChar()
        Assert.True(c3.HasValue)
        Assert.Equal("C"c, c3.Value)
        Assert.Equal(3, tr.Position)

        Dim c4 = tr.ReadChar()
        Assert.False(c4.HasValue)
        Assert.Equal(3, tr.Position)
    End Sub

    <Fact>
    Public Sub Peek_DoesNotAdvance_PositionAndReadDoesAdvance()
        Dim tr = New PositionAdjustStringReader("X")

        Dim p = tr.Peek()
        Assert.Equal(AscW("X"c), p)
        Assert.Equal(0, tr.Position)

        Dim r = tr.Read()
        Assert.Equal(AscW("X"c), r)
        Assert.Equal(1, tr.Position)

        Dim r2 = tr.Read()
        Assert.Equal(-1, r2)
    End Sub

    <Fact>
    Public Sub Substring_ReturnsExpectedWithoutChangingPosition()
        Dim input = "HelloWorld"
        Dim tr = New PositionAdjustStringReader(input)

        Dim s = tr.Substring(0, 5)
        Assert.Equal("Hello", s)
        Assert.Equal(0, tr.Position)

        Dim s2 = tr.Substring(5, 5)
        Assert.Equal("World", s2)
    End Sub

    <Fact>
    Public Sub MemoryPosition_Restore_RewindsPosition()
        Dim tr = New PositionAdjustStringReader("ABCD")

        tr.Read() ' A -> pos 1
        tr.Read() ' B -> pos 2

        Dim snap = tr.MemoryPosition()
        tr.Read() ' C -> pos 3

        snap.Restore()
        Assert.Equal(2, tr.Position)

        ' after restore, reading returns the same next char as before restore
        Dim nxt = tr.Read()
        Assert.Equal(AscW("C"c), nxt)
    End Sub

    <Fact>
    Public Sub SubChar_OutOfRange_Throws()
        Dim tr = New PositionAdjustStringReader("AB")

        ' read 0 and 1 via SubChar to populate internal buffer
        Dim c1 = tr.SubChar(0)
        Dim c2 = tr.SubChar(1)

        ' accessing beyond populated+1 should throw
        Assert.Throws(Of IndexOutOfRangeException)(Sub() tr.SubChar(3))
    End Sub

End Class

