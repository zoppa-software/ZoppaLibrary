Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 選択肢（オルタネーション）を表すコンパイル済み式。
    ''' </summary>
    Public NotInheritable Class AlternationCompiledExpression
        Implements ICompiledExpression

        ''' <summary>
        ''' 選択肢の対象範囲。
        ''' </summary>
        Private ReadOnly _target As ExpressionRange

        ''' <summary>
        ''' 選択肢の各要素を表すコンパイル済み式の列挙。
        ''' </summary>
        Private ReadOnly _subExprs() As ICompiledExpression

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="target">対象範囲。</param>
        ''' <param name="enumerable">各要素。</param>
        Public Sub New(target As ExpressionRange, enumerable As IEnumerable(Of ICompiledExpression))
            Me._target = target
            Me._subExprs = enumerable.ToArray()
        End Sub

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字列が
        ''' この式にマッチするかどうかを判定します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ruleTable">ルールテーブル。</param>
        ''' <param name="specialMethods">特殊メソッドのテーブル。</param>
        ''' <param name="answers">解析結果を格納する範囲のリスト。</param>
        ''' <param name="debugMode">デバッグモード。</param>
        ''' <param name="messages">返却メッセージリスト。</param>
        ''' <returns>マッチした場合は true。それ以外は false。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              ruleTable As SortedDictionary(Of String, RuleCompiledExpression),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              answers As List(Of AnalysisRange),
                              debugMode As Boolean,
                              messages As DebugMessage) As Boolean Implements ICompiledExpression.Match
            Dim snap = tr.MemoryPosition()
            Dim subAnswers As New List(Of AnalysisRange)()

            ' それぞれの選択肢を試す
            For Each subExpr In Me._subExprs
                snap.Restore()
                subAnswers.Clear()

                If subExpr.Match(tr, ruleTable, specialMethods, subAnswers, debugMode, messages) Then
                    ' マッチした場合は真を返す
                    If debugMode Then
                        messages.Add($"選択:{subExpr.ToString()}")
                    End If
                    answers.AddRange(subAnswers)
                    Return True
                End If
            Next

            ' どれもマッチしなかった場合は偽を返す
            snap.Restore()
            Return False
        End Function

    End Class

End Namespace
