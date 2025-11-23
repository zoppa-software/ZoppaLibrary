Option Explicit On
Option Strict On

Imports ZoppaLibrary.Parser
Imports Xunit

Public Class DigitExpressionTests

    <Theory>
    <InlineData("0")>
    <InlineData("5")>
    <InlineData("9")>
    Public Sub Match_WhenInputStartsWithDigit_ReturnsMatchAndAdvancesReader(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New DigitExpression()

        Dim r = expr.Match(tr)

        ' マッチ成功 (Invalid ではない) かつリーダ位置が 1 に進むことを確認
        Assert.NotEqual(ExpressionRange.Invalid, r)
        Assert.Equal(1, tr.Position)
    End Sub

    <Theory>
    <InlineData("a")>
    <InlineData("")>
    <InlineData("%")>
    Public Sub Match_WhenInputDoesNotStartWithDigit_ReturnsInvalidAndReaderUnchanged(input As String)
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New DigitExpression()

        Dim r = expr.Match(tr)

        ' マッチ失敗 (Invalid) かつリーダ位置は変わらないことを確認
        Assert.Equal(ExpressionRange.Invalid, r)
        Assert.Equal(0, tr.Position)
    End Sub

End Class