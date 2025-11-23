Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 識別子にマッチする範囲を表します。
    ''' </summary>
    Public NotInheritable Class IdentifierCompiledExpression
        Implements ICompiledExpression

        ''' <summary>
        ''' 対象となる識別子の範囲。
        ''' </summary>
        Private ReadOnly _target As ExpressionRange

        ''' <summary>
        ''' 識別子の名前。
        ''' </summary>
        Private ReadOnly _name As String

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="target">対象範囲。</param>
        Public Sub New(target As ExpressionRange)
            Me._target = target
            Me._name = target.ToString()
        End Sub

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字列が
        ''' この式にマッチするかどうかを判定します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ruleTable">ルールテーブル。</param>
        ''' <param name="specialMethods">特殊メソッドのテーブル。</param>
        ''' <param name="answers">解析結果を格納する範囲のリスト。</param>
        ''' <returns>マッチした場合は true。それ以外は false。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              ruleTable As SortedDictionary(Of String, RuleCompiledExpression),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              answers As List(Of AnalysisRange)) As Boolean Implements ICompiledExpression.Match
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim subAnswers As New List(Of AnalysisRange)()

            If ruleTable.ContainsKey(Me._name) Then
                If ruleTable(Me._name).Pattern.Match(tr, ruleTable, specialMethods, subAnswers) Then
                    answers.Add(New AnalysisRange(Me._name, subAnswers, tr, startPos, tr.Position))
                    Return True
                Else
                    snap.Restore()
                    Return False
                End If
            Else
                Throw New KeyNotFoundException($"識別子 '{Me._name}' はルールテーブルに存在しません。")
            End If
        End Function

    End Class

End Namespace
