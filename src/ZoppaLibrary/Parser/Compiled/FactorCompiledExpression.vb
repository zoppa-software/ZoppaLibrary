Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 繰り返し記号付きの式を表します。
    ''' </summary>
    Public NotInheritable Class FactorCompiledExpression
        Implements ICompiledExpression

        ''' <summary>
        ''' 対象範囲。
        ''' </summary>
        Private ReadOnly _target As ExpressionRange

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

            Select Case Me._subExprs.Length
                Case 1
                    If Me._subExprs(0).Match(tr, ruleTable, specialMethods, answers) Then
                        Return True
                    Else
                        snap.Restore()
                        Return False
                    End If

                Case Else
                    Dim subAnswer As New List(Of AnalysisRange)()
                    Select Case Me._subExprs(1).ToString()
                        Case "?"c
                            Me._subExprs(0).Match(tr, ruleTable, specialMethods, subAnswer)
                            answers.AddRange(subAnswer)
                            Return True

                        Case "*"c
                            Do While Me._subExprs(0).Match(tr, ruleTable, specialMethods, subAnswer)
                                ' 空実装
                            Loop
                            answers.AddRange(subAnswer)
                            Return True

                        Case "+"c
                            Dim hit As Boolean
                            Do While Me._subExprs(0).Match(tr, ruleTable, specialMethods, subAnswer)
                                hit = True
                            Loop
                            If hit Then
                                answers.AddRange(subAnswer)
                                Return True
                            Else
                                snap.Restore()
                                Return False
                            End If

                        Case "-"c
                            If Me._subExprs(0).Match(tr, ruleTable, specialMethods, subAnswer) Then
                                Dim snap2 = tr.MemoryPosition()
                                snap.Restore()
                                If Not Me._subExprs(2).Match(tr, ruleTable, specialMethods, New List(Of AnalysisRange)()) Then
                                    snap2.Restore()
                                    answers.AddRange(subAnswer)
                                    Return True
                                End If
                            End If
                            snap.Restore()
                            Return False

                        Case Else
                            Throw New InvalidOperationException("不明な繰り返し記号です。")
                    End Select
            End Select
        End Function

    End Class

End Namespace
