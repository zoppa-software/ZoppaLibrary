Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' サブルール名ノード。
    ''' </summary>
    NotInheritable Class RuleNameNode
        Inherits AnalysisNode

        ''' <summary>
        ''' 位置ごとのマッチャーキャッシュ。
        ''' </summary>
        Private ReadOnly _matchers As New SortedDictionary(Of Integer, AnalysisMatcher)()

        ''' <summary>
        ''' ルール名。
        ''' </summary>
        Private ReadOnly _ruleName As String

        ''' <summary>
        ''' 再試行可能かを取得する。
        ''' </summary
        Public Overrides ReadOnly Property IsRetry As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ノードID。</param>
        ''' <param name="range">式範囲。</param>
        Public Sub New(id As Integer, range As ExpressionRange)
            MyBase.New(id, range)
            Me._ruleName = range.ToString()
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>マッチ結果。</returns>
        Public Overrides Function Match(tr As PositionAdjustBytes, env As ABNFEnvironment, ruleName As String) As (success As Boolean, answer As ABNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()

            ' 現在位置のマッチャーを取得
            Dim position = tr.Position
            Dim matcher = Me.GetMatcher(position, env)

            ' マッチを試みる
            Dim res = matcher.Match(tr, env)
            If res.success Then
                Return (True, New ABNFAnalysisItem(Me._ruleName, matcher.GetAnswer(), tr, position, tr.Position))
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
        ''' <param name="env">ABNF環境。</param
        ''' <returns>マッチ結果。</returns>
        Public Overrides Function MoveNext(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, isRetry As Boolean, answer As ABNFAnalysisItem)
            ' 現在位置のマッチャーを取得
            Dim position = tr.Position
            Dim matcher = Me.GetMatcher(position, env)

            ' 次のパターンのマッチを試みる
            Dim res = matcher.MoveNext(tr, env)
            If res.success Then
                Return (True, True, New ABNFAnalysisItem(Me._ruleName, matcher.GetAnswer(), tr, position, tr.Position))
            Else
                Me._matchers.Remove(position)
                Return (False, True, Nothing)
            End If
        End Function

        ''' <summary>
        ''' 指定位置のマッチャーを取得する。
        ''' </summary>
        ''' <param name="position">位置。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <returns>マッチャー。</returns>
        Private Function GetMatcher(position As Integer, env As ABNFEnvironment) As AnalysisMatcher
            Dim iterator As AnalysisMatcher
            If Me._matchers.ContainsKey(position) Then
                iterator = Me._matchers(position)
            Else
                iterator = env.RuleTable(Me._ruleName).GetMatcher()
                Me._matchers.Add(position, iterator)
            End If
            Return iterator
        End Function

    End Class

End Namespace
