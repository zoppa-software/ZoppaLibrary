Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 構文解析機能を提供します（EBNF）
    ''' </summary>
    Public Module EBNFSyntaxAnalysis

        ''' <summary>
        ''' 指定されたルール群から構文解析環境を作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <returns>構文解析環境。</returns>
        Public Function CompileEnvironment(rules As String) As EBNFEnvironment
            Return CompileEnvironment(New PositionAdjustString(rules), Nothing)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す文字列。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As String,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)))) As EBNFEnvironment
            Return CompileEnvironment(New PositionAdjustString(rules), addSpecMethods)
        End Function

        ''' <summary>
        ''' 指定されたルール群から構文解析環境を作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>構文解析環境。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader) As EBNFEnvironment
            Return CompileEnvironment(rules, Nothing)
        End Function

        ''' <summary>
        ''' 指定されたルール群に基づいてルールテーブルを作成します。
        ''' </summary>
        ''' <param name="rules">ルール群を表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="addSpecMethods">特殊メソッドを追加するためのデリゲート。</param>
        ''' <returns>ルールテーブルを含む <see cref="EBNFEnvironment"/>。</returns>
        Public Function CompileEnvironment(rules As IPositionAdjustReader,
                                           addSpecMethods As Action(Of SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)))) As EBNFEnvironment
            ' 引数チェック
            If rules Is Nothing Then
                Throw New ArgumentNullException(NameOf(rules))
            End If

            '  メソッドテーブルを作成
            Dim answerEnv As New EBNFEnvironment()
            If addSpecMethods IsNot Nothing Then
                addSpecMethods(answerEnv.MethodTable)
            End If

            ' ルールテーブルを作成
            Dim expr = New GrammarExpression()
            Dim range = expr.Match(rules)
            For Each kvp In CreateRuleTable(range)
                answerEnv.RuleTable.Add(kvp.Key, kvp.Value)
            Next
            Return answerEnv
        End Function

        ''' <summary>
        ''' ルールテーブルを作成します。
        ''' </summary>
        ''' <param name="range">ルール群を表す <see cref="ExpressionRange"/>。</param>
        ''' <returns>ルールテーブル。</returns>
        Private Function CreateRuleTable(range As ExpressionRange) As SortedDictionary(Of String, RuleAnalysis)
            Dim ruleTable As New SortedDictionary(Of String, RuleAnalysis)()

            ' ルール名ごとの RuleAnalysis を作成
            For Each sr In range.SubRanges
                Dim key = sr.SubRanges(0).ToString()
                If Not ruleTable.ContainsKey(key) Then
                    ruleTable.Add(key, New RuleAnalysis(key, sr.SubRanges(1)))
                End If
            Next

            ' 単純ルートかチェックする
            For Each kvp In ruleTable
                kvp.Value.CheckSimpleRoute(ruleTable)
            Next

            Return ruleTable
        End Function

    End Module

End Namespace
