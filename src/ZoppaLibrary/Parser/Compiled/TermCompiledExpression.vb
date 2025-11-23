Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 括弧で囲まれた式にマッチするコンパイル済み式を表します。
    ''' </summary>  
    Public NotInheritable Class TermCompiledExpression
        Implements ICompiledExpression

        ''' <summary>
        ''' 対象範囲。
        ''' </summary>
        Private ReadOnly _target As ExpressionRange

        ''' <summary>
        ''' パターン文字列。
        ''' </summary>
        Private ReadOnly _pattern As String

        ''' <summary>
        ''' 各要素。
        ''' </summary>
        Private ReadOnly _subExprs() As ICompiledExpression

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="target">対象範囲。</param>
        ''' <param name="enumerable">各要素。</param>
        Public Sub New(target As ExpressionRange, enumerable As IEnumerable(Of ICompiledExpression))
            Me._target = target
            Me._pattern = target.ToString().Substring(0, 1)
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
        ''' <returns>マッチした場合は true。それ以外は false。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              ruleTable As SortedDictionary(Of String, RuleCompiledExpression),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              answers As List(Of AnalysisRange)) As Boolean Implements ICompiledExpression.Match
            Dim startPos = tr.MemoryPosition()
            Dim snap = tr.MemoryPosition()

            Select Case Me._pattern
                Case "["c
                    ' オプションパターン
                    Me._subExprs(0).Match(tr, ruleTable, specialMethods, answers)
                    Return True

                Case "{"c
                    ' オプションパターン
                    Do While Me._subExprs(0).Match(tr, ruleTable, specialMethods, answers)
                        ' 空実装
                    Loop
                    Return True

                Case Else
                    ' グループパターン
                    If Me._subExprs(0).Match(tr, ruleTable, specialMethods, answers) Then
                        Return True
                    Else
                        snap.Restore()
                        Return False
                    End If
            End Select
        End Function

    End Class

End Namespace
