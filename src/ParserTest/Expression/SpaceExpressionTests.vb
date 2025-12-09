Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit

Public Class SpaceExpressionTests

    <Fact>
    Public Sub Match_Space_AdvancesReader()
        Dim tr = New PositionAdjustStringReader(" ")
        Dim expr = New SpaceExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(1, r.[End])
        Assert.Equal(1, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_Space_AdvancesReader2()
        Dim tr = New PositionAdjustStringReader("   " & vbTab & 1)
        Dim expr = New SpaceExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(4, r.[End])
        Assert.Equal(4, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_ControlWhitespace_AdvancesReader()
        Dim controlSymbols = New String() {vbLf, vbTab, vbCr, vbFormFeed, vbBack}
        Dim expr = New SpaceExpression()
        For Each s In controlSymbols
            Dim tr = New PositionAdjustStringReader(s)
            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(1, r.[End])
            Assert.Equal(1, tr.Position)
        Next
    End Sub

    <Fact>
    Public Sub Match_EOF_ReturnsEmptyRangeAndDoesNotAdvanceReader()
        Dim tr = New PositionAdjustStringReader(String.Empty)
        Dim expr = New SpaceExpression()

        Dim r = expr.Match(tr)

        Assert.False(r.Enable)
        Assert.Equal(tr.Position, r.[Start])
        Assert.Equal(tr.Position, r.[End])
        Assert.Equal(0, tr.Position)
    End Sub

    <Theory>
    <InlineData("a")>
    <InlineData("0")>
    <InlineData("_")>
    <InlineData("|")>
    Public Sub Match_NonSpace_ReturnsEmptyRangeAndDoesNotAdvanceReader(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New SpaceExpression()

        Dim r = expr.Match(tr)

        Assert.False(r.Enable)
        Assert.Equal(tr.Position, r.[Start])
        Assert.Equal(tr.Position, r.[End])
        Assert.Equal(0, tr.Position)
    End Sub

End Class