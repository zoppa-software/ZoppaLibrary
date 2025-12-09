Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 複数のコンパイル済み式を連結した式を表します。
    ''' </summary>
    NotInheritable Class ConcatenationCompiledExpression
        Implements ICompiledExpression

        ''' <summary>
        ''' 連結要素の対象範囲。
        ''' </summary>
        Private ReadOnly _target As ExpressionRange

        ''' <summary>
        ''' 連結する各要素を表すコンパイル済み式の列挙。
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
                              answers As List(Of EBNFAnalysisItem),
                              debugMode As Boolean,
                              messages As DebugMessage) As Boolean Implements ICompiledExpression.Match
            Dim snap = tr.MemoryPosition()
            Dim subAnswers As New List(Of EBNFAnalysisItem)()

            ' 連結要素を順番にマッチさせる
            For Each subExpr In Me._subExprs
                ' マッチしなかった場合は元の位置に戻す
                If Not subExpr.Match(tr, ruleTable, specialMethods, subAnswers, debugMode, messages) Then
                    snap.Restore()
                    Return False
                End If
            Next

            ' 全ての要素がマッチした場合は結果を追加する
            answers.AddRange(subAnswers)
            Return True
        End Function

    End Class

End Namespace
