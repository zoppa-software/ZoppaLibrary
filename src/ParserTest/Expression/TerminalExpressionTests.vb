Option Explicit On
Option Strict On

Imports ZoppaLibrary.EBNF
Imports Xunit

Public Class TerminalExpressionTests

    <Fact>
    Public Sub Match_SingleQuotedOneChar_ReturnsMatchAndAdvancesReader()
        Dim input = "'a'"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TerminalExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_DoubleQuotedMultipleChars_ReturnsMatchAndAdvancesReader()
        Dim inner = "hello123_ *"
        Dim input = ChrW(34) & inner & ChrW(34) ' ダブルクォートで囲む
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TerminalExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(input.Length, r.[End])
        Assert.Equal(input.Length, tr.Position)
    End Sub

    <Fact>
    Public Sub Match_StopsAtClosingQuote_LeavesRemainingInput()
        Dim input = "'ab'c"
        Dim tr = New PositionAdjustStringReader(input)
        Dim expr = New TerminalExpression()

        Dim r = expr.Match(tr)

        Assert.True(r.Enable)
        Assert.Equal(0, r.[Start])
        Assert.Equal(4, r.[End]) ' "'ab'" の長さ
        Assert.Equal(4, tr.Position)
        ' 残りの文字が消費されていないことを確認
        Dim remaining = ChrW(tr.Read())
        Assert.Equal("c"c, remaining)
    End Sub

    <Fact>
    Public Sub Match_EmptyQuotes_ReturnsInvalid()
        Dim cases = New String() {"''", ChrW(34) & ChrW(34)}
        Dim expr = New TerminalExpression()

        For Each inp In cases
            Dim tr = New PositionAdjustStringReader(inp)
            Dim r = expr.Match(tr)
            Assert.Equal(ExpressionRange.Invalid, r)
        Next
    End Sub

    <Fact>
    Public Sub Match_MissingClosingQuote_ReturnsInvalid()
        Dim cases = New String() {"'abc", ChrW(34) & "abc"} ' 閉じクォートがないケース
        Dim expr = New TerminalExpression()

        For Each inp In cases
            Dim tr = New PositionAdjustStringReader(inp)
            Dim r = expr.Match(tr)
            Assert.Equal(ExpressionRange.Invalid, r)
        Next
    End Sub

End Class