Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit

Public Class TermExpressionTests

    <Fact>
    Public Sub Match_TerminalAtStart_ReturnsRangeAndAdvancesReader()
        Dim input = "'x'"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TermExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_IdentifierAtStart_ReturnsRangeAndAdvancesReader()
        Dim input = "abc123"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TermExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(6, r.[End])
        Assert.Equal(6, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_BracketedIdentifier_ReturnsRangeAndAdvancesReader()
        Dim input = "(a)"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TermExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(3, r.[End])
        Assert.Equal(3, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_BracketedTerminalWithSpaces_ReturnsRangeAndAdvancesReader()
        Dim input = "( 'a' )"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TermExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_Bracketed_MissingClosingBracket_ReturnsInvalidAndRestoresReader()
        Dim input = "(a"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TermExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_NoMatch_ReturnsInvalidAndReaderUnchanged()
        Dim input = "+"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TermExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_Terminal_ReturnsRange()
        Dim tr As New PositionAdjustString("'a'")
        Dim expr As New TermExpression()
        Dim startPos = tr.Position

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(startPos, r.Start)
        Assert.Equal(tr.Position, r.End)
        Assert.Equal("'a'", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_Identifier_ReturnsRange()
        Dim tr As New PositionAdjustString("abc")
        Dim expr As New TermExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal("abc", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_ParenthesesContainingIdentifier_ReturnsRange()
        Dim tr As New PositionAdjustString("(x)")
        Dim expr As New TermExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal("(x)", r.ToString())
    End Sub

    <Fact>
    Public Sub Match_UnmatchedBracket_RestoresPositionAndReturnsInvalid()
        Dim tr As New PositionAdjustString("(a")
        Dim expr As New TermExpression()
        Dim startPos = tr.Position

        Dim r = expr.Match(tr)

        Assert.False(r.Enable)
        Assert.Equal(startPos, tr.Position)
    End Sub

End Class