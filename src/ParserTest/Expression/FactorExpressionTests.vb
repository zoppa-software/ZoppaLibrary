Option Explicit On
Option Strict On

Imports ZoppaLibrary.Parser
Imports Xunit

Public Class FactorExpressionTests

    <Fact>
    Public Sub Match_TermWithStar_NoSpace_ReturnsMatchAndAdvancesReader()
        Dim tr = New PositionAdjustStringReader("a*")
        Dim expr = New FactorExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(2, r.[End])
        Assert.Equal(2, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_TermWithStar_WithSpace_ReturnsMatchAndAdvancesReader()
        Dim tr = New PositionAdjustStringReader("a *")
        Dim expr = New FactorExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(3, r.[End])
        Assert.Equal(3, tr.Position)
    End Sub

    <Theory>
    <InlineData("a+")>
    <InlineData("a?")>
    Public Sub Match_TermWithPlusOrQuestion_ReturnsMatchAndAdvancesReader(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New FactorExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_TermMinusTerm_NoSpaces_ReturnsMatchAndAdvancesReader()
        Dim tr = New PositionAdjustStringReader("a-b")
        Dim expr = New FactorExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(3, r.[End])
        Assert.Equal(3, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_TermMinusTerm_WithSpaces_ReturnsMatchAndAdvancesReader()
        Dim tr = New PositionAdjustStringReader("a - b")
        Dim expr = New FactorExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(5, r.[End])
        Assert.Equal(5, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_MissingRhsAfterMinus_ReturnsInvalidAndRestoresReader()
        Dim tr = New PositionAdjustStringReader("a-")
        Dim expr = New FactorExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_NoOperatorAfterTerm_ReturnsInvalidAndReaderUnchanged()
        Dim tr = New PositionAdjustStringReader("a")
        Dim expr = New FactorExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(1, r.[End])
        Assert.Equal(1, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_TermNotPresent_ReturnsInvalid()
        Dim tr = New PositionAdjustStringReader("+")
        Dim expr = New FactorExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

End Class