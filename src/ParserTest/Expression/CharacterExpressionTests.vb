Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit

Public Class CharacterExpressionTests

    <Theory>
    <InlineData("a")>
    <InlineData("Z")>
    <InlineData("m")>
    Public Sub Match_WhenInputIsLetter_ReturnsMatchAndAdvancesReader(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New CharacterExpression()

        Dim r = expr.Match(tr)

        Assert.NotEqual(ExpressionRange.Invalid, r)
        Assert.Equal(1, tr.Position)
    End Sub

    <Theory>
    <InlineData("0")>
    <InlineData("5")>
    <InlineData("9")>
    Public Sub Match_WhenInputIsDigit_ReturnsMatchAndAdvancesReader(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New CharacterExpression()

        Dim r = expr.Match(tr)

        Assert.NotEqual(ExpressionRange.Invalid, r)
        Assert.Equal(1, tr.Position)
    End Sub

    <Theory>
    <InlineData("[")>
    <InlineData("]")>
    <InlineData("(")>
    <InlineData(")")>
    <InlineData("+")>
    <InlineData("*")>
    <InlineData("?")>
    <InlineData(".")>
    Public Sub Match_WhenInputIsSymbol_ReturnsMatchAndAdvancesReader(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New CharacterExpression()

        Dim r = expr.Match(tr)

        Assert.NotEqual(ExpressionRange.Invalid, r)
        Assert.Equal(1, tr.Position)
    End Sub

    <Theory>
    <InlineData("_")>
    <InlineData(" ")>
    Public Sub Match_UnderscoreOrSpace_ReturnsMatchAndAdvancesReader(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New CharacterExpression()

        Dim r = expr.Match(tr)

        Assert.NotEqual(ExpressionRange.Invalid, r)
        Assert.Equal(1, tr.Position)
    End Sub

    <Theory>
    <InlineData("")>
    <InlineData("@")>
    Public Sub Match_WhenInputDoesNotMatch_ReturnsInvalidAndReaderUnchanged(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New CharacterExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

End Class