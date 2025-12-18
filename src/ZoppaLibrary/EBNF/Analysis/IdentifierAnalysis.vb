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
                              answers As List(Of EBNFAnalysisItem)) As Boolean Implements IAnalysis.Match
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of EBNFAnalysisItem)()

            ' 識別子に対応するルールを評価
            Dim hit = False
            If ruleTable.ContainsKey(Me._name) Then
                If ruleTable(Me._name).Match(tr, env, ruleTable, specialMethods, ruleName, subAnswers) Then
                    answers.Add(New EBNFAnalysisItem(Me._name, subAnswers, tr, startPos, tr.Position))
                    hit = True
                End If
            Else
                Throw New KeyNotFoundException($"識別子 '{Me._name}' はルールテーブルに存在しません。")
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
