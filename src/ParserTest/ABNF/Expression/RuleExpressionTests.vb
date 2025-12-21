Option Explicit On
Option Strict On

Imports Xunit
Imports ZoppaLibrary.BNF
Imports ZoppaLibrary.ABNF

Namespace ABNF

    Public Class RuleExpressionTests

        <Fact>
        Public Sub Match_ValidRule_ReturnsValidExpressionRange()
            ' Arrange
            Dim input = "rulename = element CRLF"
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal(0, result.Start)
                Assert.True(result.SubRanges.Count > 0)
            End Using
        End Sub

        <Fact>
        Public Sub Match_RuleWithIncrementalAlternative_ReturnsValidExpressionRange()
            ' Arrange
            Dim input = "rulename =/ element"
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal(3, result.SubRanges.Count) ' rulename + elements
            End Using
        End Sub

        <Fact>
        Public Sub Match_RuleWithWhitespace_IgnoresWhitespace()
            ' Arrange
            Dim input = "rulename   =   element   "
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal(0, result.Start)
                Assert.True(result.End <= input.Length)
            End Using
        End Sub

        <Fact>
        Public Sub Match_RuleWithComment_HandlesCommentCorrectly()
            ' Arrange
            Dim input = "rulename = element ; comment" & vbCrLf
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.True(result.SubRanges.Count >= 2) ' rulename + elements
            End Using
        End Sub

        <Fact>
        Public Sub Match_NoRuleName_ReturnsInvalidRange()
            ' Arrange
            Dim input = "= element"
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.False(result.Enable)
            End Using
        End Sub

        <Fact>
        Public Sub Match_NoEquals_ReturnsInvalidRange()
            ' Arrange
            Dim input = "rulename element"
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.False(result.Enable)
            End Using
        End Sub

        <Fact>
        Public Sub Match_NoElements_ReturnsInvalidRange()
            ' Arrange
            Dim input = "rulename = "
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.False(result.Enable)
            End Using
        End Sub

        <Fact>
        Public Sub Match_EmptyString_ReturnsInvalidRange()
            ' Arrange
            Dim input = ""
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.False(result.Enable)
            End Using
        End Sub

        <Fact>
        Public Sub Match_NullReader_ThrowsArgumentNullException()
            ' Arrange
            Dim ruleExpr As New RuleExpression()

            ' Act & Assert
            Assert.Throws(Of NullReferenceException)(
                Sub() ruleExpr.Match(Nothing)
            )
        End Sub

        <Fact>
        Public Sub Match_ComplexRule_ParsesCorrectly()
            ' Arrange
            Dim input = "complex-rule = ""literal"" SP 1*DIGIT" & vbCrLf
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal("complex-rule", result.SubRanges(0).ToString())
            End Using
        End Sub

        <Fact>
        Public Sub Match_RuleWithAlternatives_ParsesCorrectly()
            ' Arrange
            Dim input = "choice-rule = ""option1"" / ""option2"" / ""option3"""
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.True(result.SubRanges.Count >= 2) ' name + elements
            End Using
        End Sub

        <Fact>
        Public Sub Match_RuleWithOptionalElements_ParsesCorrectly()
            ' Arrange
            Dim input = "optional-rule = required [optional] ""literal"""
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal("optional-rule", result.SubRanges(0).ToString())
            End Using
        End Sub

        <Fact>
        Public Sub Match_RuleWithRepetition_ParsesCorrectly()
            ' Arrange
            Dim input = "repeat-rule = 1*3DIGIT"
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal("repeat-rule", result.SubRanges(0).ToString())
            End Using
        End Sub

        <Fact>
        Public Sub Match_MultipleRules_ProcessesFirstRuleOnly()
            ' Arrange
            Dim input = "first = ""value1""" & vbCrLf & "second = ""value2"""
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal("first", result.SubRanges(0).ToString())
                ' Position should be after first rule
                Assert.True(tr.Position < input.Length)
            End Using
        End Sub

        <Fact>
        Public Sub Match_RuleWithQuotedStrings_ParsesCorrectly()
            ' Arrange
            Dim input = "quoted-rule = ""Hello World"" SP %x22 ; double quote"
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal("quoted-rule", result.SubRanges(0).ToString())
            End Using
        End Sub

        <Fact>
        Public Sub Match_RuleWithNumericValues_ParsesCorrectly()
            ' Arrange
            Dim input = "numeric-rule = %x41-5A / %d65-90"
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.Equal("numeric-rule", result.SubRanges(0).ToString())
            End Using
        End Sub

        <Fact>
        Public Sub Match_PositionAfterMatch_IsCorrect()
            ' Arrange
            Dim input = "test-rule = ""test""" & vbCrLf
            Using tr As New PositionAdjustStringReader(input)
                Dim ruleExpr As New RuleExpression()
                Dim initialPos = tr.Position

                ' Act
                Dim result = ruleExpr.Match(tr)

                ' Assert
                Assert.True(result.Enable)
                Assert.True(tr.Position > initialPos)
            End Using
        End Sub

    End Class

End Namespace