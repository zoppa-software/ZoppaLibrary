Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 特殊シーケンス解析を表します。
    ''' </summary>
    Public NotInheritable Class SpecialSeqAnalysis
        Implements IAnalysis

        ''' <summary>識別子名。</summary>
        Private ReadOnly _name As String

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
            Me._name = range.SubRanges(0).ToString().Trim()
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
                              answers As List(Of EBNFAnalysisItem)) As (sccess As Boolean, shift As Integer) Implements IAnalysis.Match
            ' 特殊メソッドが存在しない場合はコメント扱いで成功とする
            If Not specialMethods.ContainsKey(Me._name) Then
                Return (True, 0)
            End If

            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of EBNFAnalysisItem)()

            ' 特殊メソッドを評価
            Dim res = specialMethods(Me._name)(tr)
            If res Then
                answers.Add(New EBNFAnalysisItem(Me._name, subAnswers, tr, startPos, tr.Position))
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos, Me._range)

            ' 次のパターンを評価
            Dim nextRes As (sccess As Boolean, shift As Integer) = (res, 1)
            If res Then
                nextRes = Me.AnalysisNextPattern(tr, env, ruleTable, specialMethods, ruleName, answers)
            End If

            ' 解析に失敗した場合は位置を復元
            If Not nextRes.sccess Then
                snap.Restore()
            End If
            Return nextRes
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"<? {Me._name} ?>"
        End Function

    End Class

End Namespace
