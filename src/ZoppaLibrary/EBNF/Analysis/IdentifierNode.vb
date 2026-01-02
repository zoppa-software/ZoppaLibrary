Option Explicit On
Option Strict On

Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.ABNF.ABNFSyntaxAnalysis
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 識別子ノード。
    ''' </summary>
    NotInheritable Class IdentifierNode
        Inherits AnalysisNode

        ''' <summary>位置ごとのマッチャーキャッシュ。</summary>
        Private ReadOnly _matchers As New SortedDictionary(Of Integer, IAnalysisMatcher)()

        ''' <summary>ルール名。</summary>
        Private ReadOnly _identName As String

        ''' <summary>評価範囲。</summary>
        Public Overrides ReadOnly Property Range As ExpressionRange

        ''' <summary>
        ''' 再試行可能かを取得する。
        ''' </summary>
        Public Overrides ReadOnly Property IsRetry As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ノードID。</param>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(id As Integer, range As ExpressionRange)
            MyBase.New(id)
            Me._identName = range.ToString()
            Me.Range = range
        End Sub

        ''' <summary>
        ''' キャッシュをクリアします。（実装）
        ''' </summary>
        Protected Overrides Sub ClearCacheImpl()
            Me._matchers.Clear()
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public Overrides Function Match(tr As IPositionAdjustReader, env As EBNFEnvironment, ruleName As String) As (success As Boolean, answer As EBNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()

            ' 現在位置のマッチャーを取得
            Dim position = tr.Position
            Dim matcher = Me.GetMatcher(position, env)

            ' マッチを試みる
            Dim res = matcher.Match(tr, env)
            If res.success Then
                Return (True, New EBNFAnalysisItem(Me._identName, matcher.GetAnswer(), tr, position, tr.Position))
            Else
                Me._matchers.Remove(position)
                snapPos.Restore()
                Return (False, Nothing)
            End If
        End Function

        ''' <summary>
        ''' 次のパターンのマッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public Overrides Function MoveNext(tr As IPositionAdjustReader, env As EBNFEnvironment) As (success As Boolean, answer As EBNFAnalysisItem)
            ' 現在位置のマッチャーを取得
            Dim position = tr.Position
            Dim matcher = Me.GetMatcher(position, env)

            ' 次のパターンのマッチを試みる
            Dim res = matcher.MoveNext(tr, env)
            If res.success Then
                Return (True, New EBNFAnalysisItem(Me._identName, matcher.GetAnswer(), tr, position, tr.Position))
            Else
                Me._matchers.Remove(position)
                Return (False, Nothing)
            End If
        End Function

        ''' <summary>
        ''' 指定位置のマッチャーを取得する。
        ''' </summary>
        ''' <param name="position">位置。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <returns>マッチャー。</returns>
        Private Function GetMatcher(position As Integer, env As EBNFEnvironment) As IAnalysisMatcher
            Dim iterator As IAnalysisMatcher
            If Me._matchers.ContainsKey(position) Then
                iterator = Me._matchers(position)
            ElseIf env.RuleTable.ContainsKey(Me._identName) Then
                iterator = env.RuleTable(Me._identName).GetMatcher()
                Me._matchers.Add(position, iterator)
            ElseIf env.MethodTable.ContainsKey(Me._identName) Then
                iterator = New RuleAnalysis(Me._identName, env.MethodTable(Me._identName)).GetMatcher()
                Me._matchers.Add(position, iterator)
            Else
                Throw New KeyNotFoundException($"識別子 '{Me._identName}' はルールテーブルに存在しません。")
            End If
            Return iterator
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"Identifier:{Me._identName}"
        End Function

    End Class

End Namespace
