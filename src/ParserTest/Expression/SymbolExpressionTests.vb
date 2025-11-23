Option Explicit On
Option Strict On

Imports ZoppaLibrary.Parser
Imports Xunit

Public Class SymbolExpressionTests

    <Theory>
    <InlineData("[")>
    <InlineData("]")>
    <InlineData("{")>
    <InlineData("}")>
    <InlineData("(")>
    <InlineData(")")>
    <InlineData("<")>
    <InlineData(">")>
    <InlineData("'")>
    <InlineData("""")>
    <InlineData("=")>
    <InlineData("|")>
    <InlineData(".")>
    <InlineData(",")>
    <InlineData(";")>
    <InlineData("-")>
    <InlineData("+")>
    <InlineData("*")>
    <InlineData("?")>
    Public Sub Match_WhenInputStartsWithVisibleSymbol_ReturnsMatchAndAdvancesReader(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New SymbolExpression()

        Dim r = expr.Match(tr)

        Assert.NotEqual(ExpressionRange.Invalid, r)
        Assert.Equal(1, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_ControlCharacters_AreRecognizedAndAdvanceReader()
        Dim controlSymbols = New String() {vbLf, vbTab, vbCr, vbFormFeed, vbBack}
        For Each s In controlSymbols
            Dim tr = New PositionAdjustStringReader(s)
            Dim expr = New SymbolExpression()

            Dim r = expr.Match(tr)

            Assert.NotEqual(ExpressionRange.Invalid, r)
            Assert.Equal(1, tr.Position)
        Next
    End Sub

    <Theory>
    <InlineData("a")>
    <InlineData("0")>
    <InlineData(" ")>
    <InlineData("")>
    <InlineData("Z")>
    Public Sub Match_WhenInputDoesNotStartWithSymbol_ReturnsInvalidAndReaderUnchanged(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New SymbolExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

End Class