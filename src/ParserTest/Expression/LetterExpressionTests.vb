Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit

Public Class LetterExpressionTests

    <Fact>
    Public Sub Match_Uppercase_ReturnsRange()
        Dim tr = New PositionAdjustStringReader("A")
        Dim expr = New LetterExpression()
        Dim r = expr.Match(tr)
        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(1, r.[End])
        Assert.Equal(1, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_Lowercase_ReturnsRange()
        Dim tr = New PositionAdjustStringReader("z")
        Dim expr = New LetterExpression()
        Dim r = expr.Match(tr)
        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(1, r.[End])
        Assert.Equal(1, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_NonLetter_ReturnsInvalid()
        Dim tr = New PositionAdjustStringReader("1")
        Dim expr = New LetterExpression()
        Dim r = expr.Match(tr)
        Assert.False(r.Enable)
    End Sub

    <Fact>
    Public Sub Match_Empty_ReturnsInvalid()
        Dim tr = New PositionAdjustStringReader(String.Empty)
        Dim expr = New LetterExpression()
        Dim r = expr.Match(tr)
        Assert.False(r.Enable)
    End Sub

    <Fact>
    Public Sub Match_FirstLetterOnly_AdvancesOne()
        Dim tr = New PositionAdjustStringReader("bcd")
        Dim expr = New LetterExpression()
        Dim r = expr.Match(tr)
        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(1, r.[End])
        Assert.Equal(1, tr.Position)
        Dim nxt = expr.Match(tr)
        Assert.True(nxt.Enable)
        Assert.Equal(1, nxt.[Start])
        Assert.Equal(2, nxt.[End])
        Assert.Equal(2, tr.Position)
    End Sub

End Class