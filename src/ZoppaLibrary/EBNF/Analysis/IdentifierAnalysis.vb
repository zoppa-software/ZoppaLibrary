Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 識別子解析を表します。
    ''' </summary>
    Public NotInheritable Class IdentifierAnalysis
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
            Me._name = range.ToString()
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
            ' 識別子に対応するルールが存在しない場合は例外をスロー
            If Not ruleTable.ContainsKey(Me._name) Then
                Throw New KeyNotFoundException($"識別子 '{Me._name}' はルールテーブルに存在しません。")
            End If

            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of EBNFAnalysisItem)()

            ' 識別子に対応するルールを評価
            Dim res = ruleTable(Me._name).Match(tr, env, ruleTable, specialMethods, ruleName, subAnswers)
            If res.sccess Then
                answers.Add(New EBNFAnalysisItem(Me._name, subAnswers, tr, startPos, tr.Position))
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos, Me._range)

            ' 次のパターンを評価
            If res.sccess Then
                res = Me.AnalysisNextPattern(tr, env, ruleTable, specialMethods, ruleName, answers)
            End If

            ' どれもマッチしなかった場合は偽を返す
            If Not res.sccess Then
                snap.Restore()
            End If
            Return res
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"<{Me._name}>"
        End Function

    End Class

End Namespace
