Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit
Imports ZoppaLibrary.BNF

Public Class TerminatorExpressionTests

    <Fact>
    Public Sub Match_Semicolon_ReturnsRange()
        Dim tr = New PositionAdjustStringReader(";")
        Dim expr = New TerminatorExpression()
        Dim r = expr.Match(tr)
        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(1, r.[End])
    End Sub

    <Fact>
    Public Sub Match_Period_ReturnsRange()
        Dim tr = New PositionAdjustStringReader(".")
        Dim expr = New TerminatorExpression()
        Dim r = expr.Match(tr)
        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(1, r.[End])
    End Sub

    <Fact>
    Public Sub Match_NonTerminator_ReturnsInvalid()
        Dim tr = New PositionAdjustStringReader("a")
        Dim expr = New TerminatorExpression()
        Dim r = expr.Match(tr)
        Assert.False(r.Enable)
    End Sub

    <Fact>
    Public Sub Match_Empty_ReturnsInvalid()
        Dim tr = New PositionAdjustStringReader(String.Empty)
        Dim expr = New TerminatorExpression()
        Dim r = expr.Match(tr)
        Assert.False(r.Enable)
    End Sub

End Class