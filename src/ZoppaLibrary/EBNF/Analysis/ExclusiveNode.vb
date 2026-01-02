Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 排他ノード。
    ''' </summary>
    NotInheritable Class ExclusiveNode
        Inherits AnalysisNode

        ''' <summary>評価対象。</summary>
        Private ReadOnly _groupExpr As RuleAnalysis

        ''' <summary>それ以外。</summary>
        Private ReadOnly _exceptExpr As RuleAnalysis

        ''' <summary>評価範囲。</summary>
        Public Overrides ReadOnly Property Range As ExpressionRange

        ''' <summary>
        ''' 再試行可能かを取得する。
        ''' </summary>
        Public Overrides ReadOnly Property IsRetry As Boolean
            Get
                Return False
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ID。</param>
        ''' <param name="range">評価範囲。
        Public Sub New(id As Integer, range As ExpressionRange)
            MyBase.New(id)
            Me._groupExpr = New RuleAnalysis("", range.SubRanges(0))
            Me._exceptExpr = New RuleAnalysis("", range.SubRanges(2))
            Me.Range = range
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>マッチ結果。</returns>
        Public Overrides Function Match(tr As IPositionAdjustReader, env As EBNFEnvironment, ruleName As String) As (success As Boolean, answer As EBNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()
            Dim startPos = tr.Position

            ' 要素に対応するルールを評価
            Dim groupMarcher = Me._groupExpr.GetMatcher()
            groupMarcher.ClearCache()
            Dim groupRes = groupMarcher.Match(tr, env)
            If groupRes.success Then
                Dim exceptSnap = tr.MemoryPosition()
                snapPos.Restore()

                Dim exceptMarcher = Me._exceptExpr.GetMatcher()
                exceptMarcher.ClearCache()
                Dim exceptRes = exceptMarcher.Match(tr, env)
                If Not exceptRes.success Then
                    exceptSnap.Restore()
                    Return (True, New EBNFAnalysisItem("literal", New List(Of EBNFAnalysisItem), tr, startPos, tr.Position))
                End If
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos, Me.Range)

            ' 一致しない場合は偽を返す
            snapPos.Restore()
            Return (False, Nothing)
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
            Return (False, Nothing)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"Exclusive:{Me._groupExpr}-{Me._exceptExpr}"
        End Function

    End Class

End Namespace
