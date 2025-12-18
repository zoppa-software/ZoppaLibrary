Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 要素解析を表します。
    ''' </summary>
    Public NotInheritable Class FactorAnalysis
        Implements IAnalysis

        ''' <summary>評価対象。</summary>
        Private ReadOnly _groupExpr As RuleAnalysis

        ''' <summary>それ以外。</summary>
        Private ReadOnly _exceptExpr As RuleAnalysis

        ''' <summary>評価範囲。</summary>
        Private ReadOnly _range As ExpressionRange

        ''' <summary>
        ''' 解析パターンを取得する。
        ''' </summary>
        Public ReadOnly Property Pattern As List(Of IAnalysis) Implements IAnalysis.Pattern

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(range As ExpressionRange)
            Me._groupExpr = New RuleAnalysis("", range.SubRanges(0))
            Me._exceptExpr = New RuleAnalysis("", range.SubRanges(2))
            Me._range = range
            Me.Pattern = New List(Of IAnalysis)()
        End Sub

        ''' <summary>
        ''' 解析を実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <param name="ruleTable">ルール解析テーブル。</param>
        ''' <param name="specialMethods">特殊メソッドテーブル。</param>
        ''' <param name="ruleName">現在のルール名。</param>
        ''' <param name="answers">解析結果のリスト。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              env As EBNFEnvironment,
                              ruleTable As SortedDictionary(Of String, RuleAnalysis),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              ruleName As String,
                              answers As List(Of EBNFAnalysisItem)) As Boolean Implements IAnalysis.Match
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of EBNFAnalysisItem)()

            ' 要素に対応するルールを評価
            Dim hit = False
            If Me._groupExpr.Match(tr, env, ruleTable, specialMethods, ruleName, subAnswers) Then
                Dim exceptSnap = tr.MemoryPosition()
                snap.Restore()
                If Not Me._exceptExpr.Match(tr, env, ruleTable, specialMethods, ruleName, subAnswers) Then
                    exceptSnap.Restore()
                    hit = True
                End If
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos, Me._range)

            ' 次のパターンを評価
            If hit Then
                For Each evalExpr In Me.Pattern
                    If evalExpr.Match(tr, env, ruleTable, specialMethods, ruleName, answers) Then
                        Return True
                    End If
                Next
            End If

            ' どれもマッチしなかった場合は偽を返す
            snap.Restore()
            Return False
        End Function

    End Class

End Namespace