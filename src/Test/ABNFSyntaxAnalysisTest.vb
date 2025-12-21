Imports System
Imports Xunit
Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.Strings

Public Class ABNFSyntaxAnalysisTest

    <Fact>
    Public Sub CompileEnvironment_SimpleRule_CreatesEnvironmentWithRule()
        ' Arrange: シンプルなABNFルールを作成
        Dim ruleText = U8String.NewString("rule-name = ""literal""")
        Dim reader = ruleText.GetReader()

        ' Act: CompileEnvironmentを実行
        Dim environment = ABNFSyntaxAnalysis.CompileEnvironment(reader)

        ' Assert: 環境が正しく作成されたことを確認
        Assert.NotNull(environment)
        Assert.NotNull(environment.RuleTable)
        Assert.True(environment.RuleTable.ContainsKey("rule-name"))

        ' ルールが正しく追加されているかを確認
        Dim rule = environment.RuleTable("rule-name")
        Assert.NotNull(rule)
        Assert.Equal("rule-name", rule.RuleName)
    End Sub

    <Fact>
    Public Sub CompileEnvironment_MultipleRules_CreatesEnvironmentWithAllRules()
        ' Arrange: 複数のABNFルールを作成
        Dim ruleText = U8String.NewString("rule1 = ""first""" & vbCrLf &
                                         "rule2 = ""second""" & vbCrLf &
                                         "rule3 = ""third""")
        Dim reader = ruleText.GetReader()

        ' Act: CompileEnvironmentを実行
        Dim environment = ABNFSyntaxAnalysis.CompileEnvironment(reader)

        ' Assert: すべてのルールが正しく作成されたことを確認
        Assert.NotNull(environment)
        Assert.Equal(3, environment.RuleTable.Count)
        Assert.True(environment.RuleTable.ContainsKey("rule1"))
        Assert.True(environment.RuleTable.ContainsKey("rule2"))
        Assert.True(environment.RuleTable.ContainsKey("rule3"))
    End Sub

    <Fact>
    Public Sub CompileEnvironment_RuleWithAlternatives_CreatesCorrectRule()
        ' Arrange: 選択肢を持つルールを作成
        Dim ruleText = U8String.NewString("choice-rule = ""option1"" / ""option2"" / ""option3""")
        Dim reader = ruleText.GetReader()

        ' Act: CompileEnvironmentを実行
        Dim environment = ABNFSyntaxAnalysis.CompileEnvironment(reader)

        ' Assert: 選択肢ルールが正しく作成されたことを確認
        Assert.NotNull(environment)
        Assert.True(environment.RuleTable.ContainsKey("choice-rule"))

        Dim rule = environment.RuleTable("choice-rule")
        Assert.NotNull(rule)
        Assert.Equal("choice-rule", rule.RuleName)
    End Sub

    <Fact>
    Public Sub CompileEnvironment_EmptyInput_CreatesEmptyEnvironment()
        ' Arrange: 空の入力を作成
        Dim emptyText = U8String.NewString("")
        Dim reader = emptyText.GetReader()

        ' Act: CompileEnvironmentを実行
        Dim environment = ABNFSyntaxAnalysis.CompileEnvironment(reader)

        ' Assert: 空の環境が作成されることを確認
        Assert.NotNull(environment)
        Assert.NotNull(environment.RuleTable)
        Assert.Equal(0, environment.RuleTable.Count)
    End Sub

    <Fact>
    Public Sub CompileEnvironment_RuleWithRepetition_CreatesCorrectRule()
        ' Arrange: 繰り返しを持つルールを作成
        Dim ruleText = U8String.NewString("repeat-rule = 1*3(""a"" / ""b"")")
        Dim reader = ruleText.GetReader()

        ' Act: CompileEnvironmentを実行
        Dim environment = ABNFSyntaxAnalysis.CompileEnvironment(reader)

        ' Assert: 繰り返しルールが正しく作成されたことを確認
        Assert.NotNull(environment)
        Assert.True(environment.RuleTable.ContainsKey("repeat-rule"))

        Dim rule = environment.RuleTable("repeat-rule")
        Assert.NotNull(rule)
        Assert.Equal("repeat-rule", rule.RuleName)
    End Sub

    <Fact>
    Public Sub CompileEnvironment_DuplicateRuleNames_KeepsFirstRule()
        ' Arrange: 重複するルール名を持つ入力を作成
        Dim ruleText = U8String.NewString("same-rule = ""first""" & vbCrLf &
                                         "same-rule = ""second""")
        Dim reader = ruleText.GetReader()

        ' Act: CompileEnvironmentを実行
        Dim environment = ABNFSyntaxAnalysis.CompileEnvironment(reader)

        ' Assert: 最初のルールのみが保持されることを確認
        Assert.NotNull(environment)
        Assert.True(environment.RuleTable.ContainsKey("same-rule"))
        Assert.Equal(1, environment.RuleTable.Count)

        Dim rule = environment.RuleTable("same-rule")
        Assert.Equal("same-rule", rule.RuleName)
    End Sub

    <Fact>
    Public Sub CompileEnvironment_MethodTableContainsSpecialMethods()
        ' Arrange: 任意の有効なルールを作成
        Dim ruleText = U8String.NewString("test-rule = ""test""")
        Dim reader = ruleText.GetReader()

        ' Act: CompileEnvironmentを実行
        Dim environment = ABNFSyntaxAnalysis.CompileEnvironment(reader)

        ' Assert: 特殊メソッドテーブルが正しく初期化されていることを確認
        Assert.NotNull(environment.MethodTable)
        Assert.True(environment.MethodTable.Count > 0)

        ' 基本的な特殊メソッドが含まれていることを確認
        Assert.True(environment.MethodTable.ContainsKey("ALPHA"))
        Assert.True(environment.MethodTable.ContainsKey("DIGIT"))
        Assert.True(environment.MethodTable.ContainsKey("HEXDIG"))
        Assert.True(environment.MethodTable.ContainsKey("SP"))
        Assert.True(environment.MethodTable.ContainsKey("HTAB"))
        Assert.True(environment.MethodTable.ContainsKey("CRLF"))
    End Sub

    <Fact>
    Public Sub CompileEnvironment_ComplexRule_CreatesCorrectStructure()
        ' Arrange: 複雑なABNFルールを作成
        Dim ruleText = U8String.NewString("uri = scheme "":"" hier-part [ ""?"" query ] [ ""#"" fragment ]" & vbCrLf &
                                         "scheme = ALPHA *( ALPHA / DIGIT / ""+"" / ""-"" / ""."" )" & vbCrLf &
                                         "query = *( pchar / ""/"" / ""?"" )")
        Dim reader = ruleText.GetReader()

        ' Act: CompileEnvironmentを実行
        Dim environment = ABNFSyntaxAnalysis.CompileEnvironment(reader)

        ' Assert: 複雑なルールが正しく解析されることを確認
        Assert.NotNull(environment)
        Assert.Equal(3, environment.RuleTable.Count)
        Assert.True(environment.RuleTable.ContainsKey("uri"))
        Assert.True(environment.RuleTable.ContainsKey("scheme"))
        Assert.True(environment.RuleTable.ContainsKey("query"))

        ' 各ルールが適切に作成されていることを確認
        Assert.Equal("uri", environment.RuleTable("uri").RuleName)
        Assert.Equal("scheme", environment.RuleTable("scheme").RuleName)
        Assert.Equal("query", environment.RuleTable("query").RuleName)
    End Sub

End Class