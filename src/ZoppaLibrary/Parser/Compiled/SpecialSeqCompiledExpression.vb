Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 特殊シーケンスにマッチする式を表します。
    ''' </summary>
    Public NotInheritable Class SpecialSeqCompiledExpression
        Implements ICompiledExpression

        ''' <summary>
        ''' 対象範囲。
        ''' </summary>
        Private _target As ExpressionRange

        ''' <summary>
        ''' メソッド名。
        ''' </summary>
        Private _methodName As String

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="target">対象範囲。</param>
        Public Sub New(target As ExpressionRange)
            Me._target = target
            Me._methodName = target.SubRanges(0).ToString().Trim()
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

            If specialMethods.ContainsKey(Me._methodName) AndAlso
               specialMethods(Me._methodName)(tr) Then
                answers.Add(New AnalysisRange(Me._methodName, New List(Of AnalysisRange)(), tr, startPos, tr.Position))
                Return True
            Else
                snap.Restore()
                Return False
            End If
        End Function

    End Class

End Namespace
