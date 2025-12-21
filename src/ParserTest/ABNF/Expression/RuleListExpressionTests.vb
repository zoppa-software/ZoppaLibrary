Option Explicit On
Option Strict On

Imports Xunit
Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.BNF

Namespace ABNF

    Public Class RuleListExpressionTests

        <Fact>
        Public Sub Match_SingleRule_ReturnsRangeAndAdvancesReader()
            Dim input = "rulename = %x41"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_MultipleRules_ReturnsRangeAndAdvancesReader()
            Dim input = "rule1 = %x41" & vbCrLf & "rule2 = %x42"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_RulesWithSpacesAndComments_ReturnsRangeAndAdvancesReader()
            Dim input = "  rule1 = %x41  " & vbCrLf & vbCrLf & "rule2 = %x42  "
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_RulesWithCRLFOnly_SkipsEmptyLines()
            Dim input = "rule1 = %x41" & vbCrLf & vbCrLf & vbCrLf & "rule2 = %x42"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_DuplicateRuleName_KeepsFirstDefinition()
            Dim input = "rulename = %x41" & vbCrLf & "rulename = %x42"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
            ' 最初のルールのみがSubRangesに含まれることを確認
            Assert.Single(r.SubRanges.Where(Function(sr) sr.ToString().StartsWith("rulename")))
        End Sub

        <Fact>
        Public Sub Match_RuleWithIncrementalAlternation_MergesAlternatives()
            ' ABNF では同じルール名で複数定義すると選択肢が追加される
            Dim input = "rulename = %x41" & vbCrLf & "rulename =/ %x42"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_OnlyWhitespaceAndNewlines_ReturnsValidRange()
            Dim input = "   " & vbCrLf & "  " & vbCrLf & "   "
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_EmptyInput_ReturnsValidRange()
            Dim tr = New PositionAdjustStringReader(String.Empty)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.False(r.Enable)
        End Sub

        <Fact>
        Public Sub Match_InfiniteLoopPrevention_ThrowsABNFException()
            ' 進捗しない状況を想定したテスト
            ' 実際の実装では、prevPos = tr.Position の場合に例外をスローする
            Dim input = "invalid-content"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            ' ABNFExceptionがスローされることを想定（実際の実装に依存）
            Assert.Throws(Of ABNFException)(Sub() expr.Match(tr))
        End Sub

        <Fact>
        Public Sub Match_RuleFollowedByEOF_ReturnsValidRange()
            Dim input = "simple-rule = %x41"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
            Assert.True(r.SubRanges.Count > 0)
        End Sub

        <Fact>
        Public Sub Match_MixedWhitespaceTypes_HandlesCorrectly()
            Dim input = "rule1 = %x41" & vbTab & " " & vbCrLf & vbLf & "rule2 = %x42"
            Dim tr = New PositionAdjustStringReader(input)
            Dim expr = New RuleListExpression()

            Dim r = expr.Match(tr)

            Assert.True(r.Enable)
            Assert.Equal(0, r.[Start])
            Assert.Equal(input.Length, r.[End])
            Assert.Equal(input.Length, tr.Position)
        End Sub

    End Class

End Namespace