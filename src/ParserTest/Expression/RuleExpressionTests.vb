Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit

Public Class RuleExpressionTests

    <Fact>
    Public Sub Match_SimpleRule_WithSpaces_ReturnsMatchAndAdvancesReader()
        Dim input = "id = 'x';"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New RuleExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_SimpleRule_NoSpaces_ReturnsMatchAndAdvancesReader()
        Dim input = "id='val'."
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New RuleExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_RhsIsConcatenation_ReturnsMatchAndAdvancesReader()
        Dim input = "rule = a*,b+;"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New RuleExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_MissingEquals_ReturnsInvalidAndRestoresReader()
        Dim input = "id 'x';"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New RuleExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_MissingTerminator_ReturnsInvalidAndRestoresReader()
        Dim input = "id = 'x'"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New RuleExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_InvalidLhs_ReturnsInvalidAndRestoresReader()
        Dim input = "1bad = 'x';"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New RuleExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_MissingRhs_ReturnsInvalidAndRestoresReader()
        Dim input = "id = ;"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New RuleExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

End Class