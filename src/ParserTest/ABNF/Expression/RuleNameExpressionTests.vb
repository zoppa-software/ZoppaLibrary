Option Explicit On
Option Strict On

Imports Xunit
Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.BNF
Imports ZoppaLibrary.EBNF

Namespace ABNF

    Public Class RuleNameExpressionTests

        <Fact>
        Public Sub Match_SingleAlpha_ReturnsRangeAndAdvancesReader()
            Dim tr = New PositionAdjustStringReader("a")
            Dim expr = New RuleNameExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(1, r.[End])
            Assert.Equal(1, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_MultipleAlphaDigitHyphen_ReturnsRangeAndAdvancesReader()
            Dim input = "rule-1name2"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleNameExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_StopsAtInvalidCharacter_ReturnsPartialRange()
            Dim tr = New PositionAdjustStringReader("abc$def")
            Dim expr = New RuleNameExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(3, r.[End])
            Assert.Equal(3, tr.Position)
            Assert.Equal(AscW("$"c), tr.Peek())
        End Sub

        <Fact>
        Public Sub Match_StartsWithDigit_ReturnsInvalidAndReaderUnchanged()
            Dim tr = New PositionAdjustStringReader("1abc")
            Dim expr = New RuleNameExpression()

            Dim r = expr.Match(tr)

            Assert.Equal(ExpressionRange.Invalid, r)
            Assert.Equal(0, tr.Position)
        End Sub

        <Theory>
        <InlineData("")>
        <InlineData("_")>
        <InlineData("-")>
        <InlineData("!name")>
        Public Sub Match_InvalidStartCharacters_ReturnsInvalidAndReaderUnchanged(input As String)
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleNameExpression()

            Dim r = expr.Match(tr)

            Assert.Equal(ExpressionRange.Invalid, r)
            Assert.Equal(0, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_ConsecutiveHyphens_AreAllowed()
            Dim tr = New PositionAdjustStringReader("a--b")
            Dim expr = New RuleNameExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(4, r.[End])
            Assert.Equal(4, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_AfterValidName_StartsFromCurrentPosition()
            Dim tr = New PositionAdjustStringReader("prefix rule-name rest")
            ' advance to start of rule-name
            For i As Integer = 1 To 7
                tr.Read()
            Next
            Dim expr = New RuleNameExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(7, r.[Start])
            Assert.Equal(16, r.[End]) ' "rule-name" length = 9, 7+9=16
            Assert.Equal(16, tr.Position)
        End Sub

    End Class

End Namespace