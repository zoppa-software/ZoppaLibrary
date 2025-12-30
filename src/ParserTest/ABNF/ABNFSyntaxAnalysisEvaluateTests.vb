Option Explicit On
Option Strict On

Imports System
Imports System.Text
Imports Xunit
Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.BNF
Imports ZoppaLibrary.Strings

''' <summary>
''' ABNFSyntaxAnalysis.Evaluateメソッドの詳細テストクラス
''' </summary>
Public Class ABNFSyntaxAnalysisEvaluateTests

#Region "基本的な成功ケース"

    <Fact>
    Public Sub Evaluate_SimpleRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "simple = ""test"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("test"))

        ' Act
        Dim result = env.Evaluate("simple", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("simple", result.Identifier)
        Assert.Equal(0, result.Start)
        Assert.Equal(4, result.End)
        Assert.Equal(New Byte() {&H74, &H65, &H73, &H74}, result.GetBytes().ToArray())
    End Sub

    <Fact>
    Public Sub Evaluate_AlternationRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "choice = ""hello"" / ""world"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target1 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("hello"))
        Dim target2 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("world"))

        ' Act
        Dim result1 = env.Evaluate("choice", target1)
        Dim result2 = env.Evaluate("choice", target2)

        ' Assert
        Assert.Equal("choice", result1.Identifier)
        Assert.Equal(New Byte() {&H68, &H65, &H6C, &H6C, &H6F}, result1.GetBytes().ToArray())
        Assert.Equal("choice", result2.Identifier)
        Assert.Equal(New Byte() {&H77, &H6F, &H72, &H6C, &H64}, result2.GetBytes().ToArray())
    End Sub

    <Fact>
    Public Sub Evaluate_ConcatenationRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "concat = ""hello"" "" "" ""world"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("hello world"))

        ' Act
        Dim result = env.Evaluate("concat", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("concat", result.Identifier)
        Assert.Equal(11, result.End - result.Start)
    End Sub

    <Fact>
    Public Sub Evaluate_RepetitionRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "repeat = 3DIGIT"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("123"))

        ' Act
        Dim result = env.Evaluate("repeat", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("repeat", result.Identifier)
        Assert.Equal(3, result.End - result.Start)
        Assert.Equal(New Byte() {&H31, &H32, &H33}, result.GetBytes().ToArray())
    End Sub

    <Fact>
    Public Sub Evaluate_OptionalRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "optional = [""prefix""] ""main"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target1 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("prefixmain"))
        Dim target2 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("main"))

        ' Act
        Dim result1 = env.Evaluate("optional", target1)
        Dim result2 = env.Evaluate("optional", target2)

        ' Assert
        Assert.Equal(10, result1.End - result1.Start)
        Assert.Equal(4, result2.End - result2.Start)
    End Sub

#End Region

#Region "特殊メソッド使用ケース"

    <Fact>
    Public Sub Evaluate_AlphaRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "alpha-test = 3ALPHA"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("ABC"))

        ' Act
        Dim result = env.Evaluate("alpha-test", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("alpha-test", result.Identifier)
        Assert.Equal(New Byte() {&H41, &H42, &H43}, result.GetBytes().ToArray())
    End Sub

    <Fact>
    Public Sub Evaluate_DigitRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "digit-test = 2*4DIGIT"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target1 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("12"))
        Dim target2 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("1234"))

        ' Act
        Dim result1 = env.Evaluate("digit-test", target1)
        Dim result2 = env.Evaluate("digit-test", target2)

        ' Assert
        Assert.Equal(2, result1.End - result1.Start)
        Assert.Equal(4, result2.End - result2.Start)
    End Sub

    <Fact>
    Public Sub Evaluate_HexDigitRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "hex-test = 4HEXDIG"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("A1F9"))

        ' Act
        Dim result = env.Evaluate("hex-test", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("hex-test", result.Identifier)
        Assert.Equal(New Byte() {&H41, &H31, &H46, &H39}, result.GetBytes().ToArray())
    End Sub

    <Fact>
    Public Sub Evaluate_WhitespaceRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "ws-test = ""start"" WSP ""end"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target1 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("start end"))
        Dim target2 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("start" & vbTab & "end"))

        ' Act
        Dim result1 = env.Evaluate("ws-test", target1)
        Dim result2 = env.Evaluate("ws-test", target2)

        ' Assert
        Assert.Equal(9, result1.End - result1.Start)
        Assert.Equal(9, result2.End - result2.Start)
    End Sub

#End Region

#Region "エラーケース"

    <Fact>
    Public Sub Evaluate_NonexistentRule_ThrowsException()
        ' Arrange
        Dim input = "existing = ""test"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("test"))

        ' Act & Assert
        Dim exception = Assert.Throws(Of ABNFException)(
            Sub()
                env.Evaluate("nonexistent", target)
            End Sub
        )
        Assert.Contains("指定された識別子 'nonexistent' はルールに存在しません", exception.Message)
    End Sub

    <Fact>
    Public Sub Evaluate_PartialMatch_ThrowsException()
        ' Arrange
        Dim input = "exact = ""hello"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("helloworld"))

        ' Act & Assert
        Dim exception = Assert.Throws(Of ABNFException)(
            Sub()
                env.Evaluate("exact", target)
            End Sub
        )
        Assert.Contains("識別子 'exact' の解析に失敗しました", exception.Message)
    End Sub

    <Fact>
    Public Sub Evaluate_NoMatch_ThrowsException()
        ' Arrange
        Dim input = "nomatch = ""expected"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("different"))

        ' Act & Assert
        Dim exception = Assert.Throws(Of ABNFException)(
            Sub()
                env.Evaluate("nomatch", target)
            End Sub
        )
        Assert.Contains("識別子 'nomatch' の解析に失敗しました", exception.Message)
    End Sub

    <Fact>
    Public Sub Evaluate_EmptyInput_ThrowsException()
        ' Arrange
        Dim input = "required = ""required"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(New Byte() {})

        ' Act & Assert
        Assert.Throws(Of ABNFException)(
            Sub()
                env.Evaluate("required", target)
            End Sub
        )
    End Sub

    <Fact>
    Public Sub Evaluate_RepetitionCountMismatch_ThrowsException()
        ' Arrange
        Dim input = "exact3 = 3DIGIT"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target1 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("12"))    ' 2桁のみ
        Dim target2 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("1234"))  ' 4桁

        ' Act & Assert
        Assert.Throws(Of ABNFException)(
            Sub()
                env.Evaluate("exact3", target1)
            End Sub
        )
        Assert.Throws(Of ABNFException)(
            Sub()
                env.Evaluate("exact3", target2)
            End Sub
        )
    End Sub

#End Region

#Region "複雑なケース"

    <Fact>
    Public Sub Evaluate_NestedRules_ReturnsCorrectResult()
        ' Arrange
        Dim input = "" &
"main-rule = prefix content suffix
prefix = ""["" 
content = 1*ALPHA
suffix = ""]"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("[hello]"))

        ' Act
        Dim result = env.Evaluate("main-rule", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("main-rule", result.Identifier)
        Assert.Equal(7, result.End - result.Start)

        ' サブルールの確認
        Assert.NotNull(result("prefix"))
        Assert.NotNull(result("content"))
        Assert.NotNull(result("suffix"))
    End Sub

    <Fact>
    Public Sub Evaluate_RecursiveRule_ReturnsCorrectResult()
        ' Arrange
        Dim input = "" &
"list = item *("","" item)
item = 1*ALPHA"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("apple,banana,cherry"))

        ' Act
        Dim result = env.Evaluate("list", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("list", result.Identifier)
        Assert.Equal(19, result.End - result.Start)
    End Sub

    <Fact>
    Public Sub Evaluate_NumericValue_ReturnsCorrectResult()
        ' Arrange
        Dim input = "hex-val = %x41-5A"  ' A-Z
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("M"))  ' &H4D

        ' Act
        Dim result = env.Evaluate("hex-val", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("hex-val", result.Identifier)
        Assert.Equal(New Byte() {&H4D}, result.GetBytes().ToArray())
    End Sub

    <Fact>
    Public Sub Evaluate_ProseValue_ReturnsCorrectResult()
        ' Arrange
        Dim input = "prose-val = <any ASCII character>"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        ' 散文値の実装は環境依存のため、スキップまたは模擬実装が必要
    End Sub

#End Region

#Region "パフォーマンス・境界値テスト"

    <Fact>
    Public Sub Evaluate_LargeInput_PerformsWell()
        ' Arrange
        Dim input = "many-digits = *DIGIT"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim largeNumber = New String("1"c, 1000)
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes(largeNumber))

        ' Act
        Dim result = env.Evaluate("many-digits", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal(1000, result.End - result.Start)
    End Sub

    <Fact>
    Public Sub Evaluate_DeepNesting_HandlesCorrectly()
        ' Arrange
        Dim input = "" &
"level1 = level2 
level2 = level3 
level3 = level4 
level4 = level5 
level5 = ""deep"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("deep"))

        ' Act
        Dim result = env.Evaluate("level1", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("level1", result.Identifier)
        Assert.Equal(4, result.End - result.Start)
    End Sub

    <Fact>
    Public Sub Evaluate_SpecialCharacters_HandlesCorrectly()
        ' Arrange
        Dim input = "special = DQUOTE ""content"" DQUOTE"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("""content"""))

        ' Act
        Dim result = env.Evaluate("special", target)

        ' Assert
        Assert.NotNull(result)
        Assert.Equal("special", result.Identifier)
        Assert.Equal(9, result.End - result.Start)
    End Sub

#End Region

#Region "カスタム特殊メソッド"

    <Fact>
    Public Sub Evaluate_CustomSpecialMethod_ReturnsCorrectResult()
        ' Arrange
        Dim input = "custom = CUSTOM-VOWEL"
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))

        ' カスタムメソッド追加
        env.AddSpecialMethods("CUSTOM-VOWEL",
            Function(tr As PositionAdjustBytes) As Boolean
                Dim c = tr.Peek()
                If c = AscW("a"c) OrElse c = AscW("e"c) OrElse c = AscW("i"c) OrElse
                   c = AscW("o"c) OrElse c = AscW("u"c) Then
                    tr.Read()
                    Return True
                End If
                Return False
            End Function)

        Dim target1 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("a"))
        Dim target2 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("e"))
        Dim target3 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("x"))

        ' Act
        Dim result1 = env.Evaluate("custom", target1)
        Dim result2 = env.Evaluate("custom", target2)

        ' Assert
        Assert.NotNull(result1)
        Assert.NotNull(result2)
        Assert.Equal(New Byte() {&H61}, result1.GetBytes().ToArray()) ' 'a'
        Assert.Equal(New Byte() {&H65}, result2.GetBytes().ToArray()) ' 'e'

        Assert.Throws(Of ABNFException)(
            Sub()
                env.Evaluate("custom", target3) ' 'x' - 母音ではない
            End Sub
        )
    End Sub

#End Region

#Region "位置情報テスト"

    <Fact>
    Public Sub Evaluate_PositionInformation_IsCorrect()
        ' Arrange
        Dim input = "positioned = ""start"" ""middle"" ""end"""
        Dim env = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input))
        Dim target = New PositionAdjustBytes(Encoding.UTF8.GetBytes("startmiddleend"))

        ' Act
        Dim result = env.Evaluate("positioned", target)

        ' Assert
        Assert.Equal(0, result.Start)
        Assert.Equal(14, result.End)
        Assert.Equal(14, target.Position) ' 入力が完全に消費されている
    End Sub

    <Fact>
    Public Sub Evaluate_MultipleEvaluations_MaintainsState()
        ' Arrange
        Dim input1 = "rule1 = ""hello"""
        Dim input2 = "rule2 = ""world"""
        Dim env1 = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input1))
        Dim env2 = ABNFSyntaxAnalysis.CompileEnvironment(New PositionAdjustString(input2))

        Dim target1 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("hello"))
        Dim target2 = New PositionAdjustBytes(Encoding.UTF8.GetBytes("world"))

        ' Act
        Dim result1 = env1.Evaluate("rule1", target1)
        Dim result2 = env2.Evaluate("rule2", target2)

        ' Assert
        Assert.Equal("rule1", result1.Identifier)
        Assert.Equal("rule2", result2.Identifier)
        Assert.NotEqual(result1.Identifier, result2.Identifier)
    End Sub

#End Region

End Class