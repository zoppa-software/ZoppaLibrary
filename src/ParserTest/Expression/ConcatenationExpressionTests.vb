Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit

Public Class ConcatenationExpressionTests

    <Fact>
    Public Sub Match_SingleFactorWithoutComma_ReturnsMatchAndAdvancesReader()
        Dim input = "a*"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New ConcatenationExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_MultipleFactorsSeparatedByComma_ReturnsMatchAndAdvancesReader()
        Dim input = "a*,b+"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New ConcatenationExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_TrailingComma_IsAcceptedAndIncludedInRange()
        Dim input = "a*,"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New ConcatenationExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_WithSpaces_AroundBlocks_HandlesSpacesCorrectly()
        Dim input = "  a*  ,  b+  "
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New ConcatenationExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Theory>
    <InlineData("")>
    Public Sub Match_InvalidWhenNoFactor_ReturnsInvalidAndReaderUnchanged(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New ConcatenationExpression()

        Dim r = expr.Match(tr)

        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

End Class