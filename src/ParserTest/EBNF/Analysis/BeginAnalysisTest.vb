Option Explicit On
Option Strict On

Imports Xunit
Imports ZoppaLibrary.BNF
Imports ZoppaLibrary.EBNF

Namespace Analysis

    Public Class BeginAnalysisTest

        <Fact>
        Public Sub New_WithValidRange_CreatesInstance()
            ' Arrange
            Dim range As New ExpressionRange(Nothing, Nothing, 0, 0, Array.Empty(Of ExpressionRange)())

            ' Act
            Dim analysis As New BeginAnalysis(range)

            ' Assert
            Assert.NotNull(analysis)
            Assert.NotNull(analysis.Pattern)
            Assert.Empty(analysis.Pattern)
        End Sub

        <Fact>
        Public Sub Match_AlwaysReturnsSuccess()
            ' Arrange
            Dim range As New ExpressionRange(Nothing, Nothing, 0, 0, Array.Empty(Of ExpressionRange)())
            Dim analysis As New BeginAnalysis(range)
            Dim tr = New PositionAdjustStringReader("test input")
            Dim env As New EBNFEnvironment()
            Dim ruleTable As New SortedDictionary(Of String, RuleAnalysis)()
            Dim specialMethods As New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
            Dim answers As New List(Of EBNFAnalysisItem)()

            ' Act
            Dim result = analysis.Match(tr, env, ruleTable, specialMethods, "testRule", answers)

            ' Assert
            Assert.True(result.sccess)
            Assert.Equal(0, result.shift)
        End Sub

        <Fact>
        Public Sub Match_DoesNotAdvanceReader()
            ' Arrange
            Dim range As New ExpressionRange(Nothing, Nothing, 0, 0, Array.Empty(Of ExpressionRange)())
            Dim analysis As New BeginAnalysis(range)
            Dim tr = New PositionAdjustStringReader("test input")
            Dim env As New EBNFEnvironment()
            Dim ruleTable As New SortedDictionary(Of String, RuleAnalysis)()
            Dim specialMethods As New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
            Dim answers As New List(Of EBNFAnalysisItem)()
            Dim initialPosition = tr.Position

            ' Act
            analysis.Match(tr, env, ruleTable, specialMethods, "testRule", answers)

            ' Assert
            Assert.Equal(initialPosition, tr.Position)
        End Sub

        <Fact>
        Public Sub Match_WithEmptyInput_ReturnsSuccess()
            ' Arrange
            Dim range As New ExpressionRange(Nothing, Nothing, 0, 0, Array.Empty(Of ExpressionRange)())
            Dim analysis As New BeginAnalysis(range)
            Dim tr = New PositionAdjustStringReader("")
            Dim env As New EBNFEnvironment()
            Dim ruleTable As New SortedDictionary(Of String, RuleAnalysis)()
            Dim specialMethods As New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
            Dim answers As New List(Of EBNFAnalysisItem)()

            ' Act
            Dim result = analysis.Match(tr, env, ruleTable, specialMethods, "testRule", answers)

            ' Assert
            Assert.True(result.sccess)
            Assert.Equal(0, result.shift)
        End Sub

        <Fact>
        Public Sub Match_WithNullParameters_ReturnsSuccess()
            ' Arrange
            Dim range As New ExpressionRange(Nothing, Nothing, 0, 0, Array.Empty(Of ExpressionRange)())
            Dim analysis As New BeginAnalysis(range)
            Dim tr = New PositionAdjustStringReader("test")
            Dim env As New EBNFEnvironment()
            Dim ruleTable As New SortedDictionary(Of String, RuleAnalysis)()
            Dim specialMethods As New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
            Dim answers As New List(Of EBNFAnalysisItem)()

            ' Act
            Dim result = analysis.Match(tr, env, ruleTable, specialMethods, Nothing, answers)

            ' Assert
            Assert.True(result.sccess)
            Assert.Equal(0, result.shift)
        End Sub

        <Fact>
        Public Sub Match_DoesNotModifyAnswers()
            ' Arrange
            Dim range As New ExpressionRange(Nothing, Nothing, 0, 0, Array.Empty(Of ExpressionRange)())
            Dim analysis As New BeginAnalysis(range)
            Dim tr = New PositionAdjustStringReader("test input")
            Dim env As New EBNFEnvironment()
            Dim ruleTable As New SortedDictionary(Of String, RuleAnalysis)()
            Dim specialMethods As New SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean))()
            Dim answers As New List(Of EBNFAnalysisItem)()
            Dim initialCount = answers.Count

            ' Act
            analysis.Match(tr, env, ruleTable, specialMethods, "testRule", answers)

            ' Assert
            Assert.Equal(initialCount, answers.Count)
        End Sub

        <Fact>
        Public Sub ToString_ReturnsExpectedString()
            ' Arrange
            Dim range As New ExpressionRange(Nothing, Nothing, 0, 0, Array.Empty(Of ExpressionRange)())
            Dim analysis As New BeginAnalysis(range)

            ' Act
            Dim result = analysis.ToString()

            ' Assert
            Assert.Equal("<ŠJŽn>", result)
        End Sub

        <Fact>
        Public Sub Pattern_IsEmptyList()
            ' Arrange
            Dim range As New ExpressionRange(Nothing, Nothing, 0, 0, Array.Empty(Of ExpressionRange)())
            Dim analysis As New BeginAnalysis(range)

            ' Assert
            Assert.NotNull(analysis.Pattern)
            Assert.IsType(Of List(Of IAnalysis))(analysis.Pattern)
            Assert.Empty(analysis.Pattern)
        End Sub

    End Class

End Namespace
