Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit

Public Class IdentifierExpressionTests

    <Theory>
    <InlineData("a", 1)>
    <InlineData("Z", 1)>
    <InlineData("ab", 2)>
    <InlineData("a1", 2)>
    <InlineData("a_b", 3)>
    <InlineData("Abc123", 6)>
    Public Sub Match_ValidIdentifiers_ReturnsRangeAndAdvancesReader(input As String, expectedLength As Integer)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New IdentifierExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(expectedLength, r.[End])
        Assert.Equal(expectedLength, tr.Position)
    End Sub

    <Theory>
    <InlineData("a+", 1)>
    <InlineData("a1+", 2)>
    <InlineData("abc_def+", 7)>
    Public Sub Match_StopsAtNonIdentifierCharacter_ReturnsPartialRange(input As String, expectedLength As Integer)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New IdentifierExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(expectedLength, r.[End])
        Assert.Equal(expectedLength, tr.Position)
    End Sub

    <Theory>
    <InlineData("")>
    <InlineData("1a")>
    <InlineData("_a")>
    <InlineData("0abc")>
    Public Sub Match_WhenNotStartingWithLetter_ReturnsInvalidAndReaderUnchanged(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New IdentifierExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

End Class