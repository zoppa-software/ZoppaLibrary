Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' コンパイル済み式のインターフェイスを表します。
    ''' </summary>
    Public Interface ICompiledExpression

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
        Function Match(tr As IPositionAdjustReader,
                       ruleTable As SortedDictionary(Of String, RuleCompiledExpression),
                       specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                       answers As List(Of AnalysisRange),
                       debugMode As Boolean,
                       messages As DebugMessage) As Boolean

    End Interface

End Namespace
