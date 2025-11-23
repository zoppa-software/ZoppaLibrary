Option Explicit On
Option Strict On

Imports System.Text

Namespace Parser

    ''' <summary>
    ''' 文字列にマッチする式を表します。
    ''' </summary>
    Public NotInheritable Class CharacterCompiledExpression
        Implements ICompiledExpression

        ''' <summary>
        ''' マッチ対象の文字列。
        ''' </summary>
        Private ReadOnly _target As ExpressionRange

        ''' <summary>
        ''' 文字列の値。
        ''' </summary>
        Private ReadOnly _stringValue As String

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="target">対象範囲。</param>
        Public Sub New(target As ExpressionRange)
            Me._target = target
            Me._stringValue = target.ToString()
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

            Dim buf As New StringBuilder()
            Do While tr.Peek() <> -1 AndAlso buf.Length < Me._stringValue.Length
                Dim c = ChrW(tr.Read())
                If c = "\"c Then
                    Dim nextChar = ChrW(tr.Read())
                    Select Case nextChar
                        Case "n"c
                            buf.Append(vbLf)
                        Case "r"c
                            buf.Append(vbCr)
                        Case "t"c
                            buf.Append(vbTab)
                        Case "f"c
                            buf.Append(vbFormFeed)
                        Case "b"c
                            buf.Append(vbBack)
                        Case "\"c
                            buf.Append("\"c)
                        Case Else
                            Throw New InvalidCastException($"不明なエスケープシーケンスです: \{nextChar}")
                    End Select
                Else
                    buf.Append(c)
                End If
            Loop

            If buf.ToString() = Me._stringValue Then
                Return True
            Else
                snap.Restore()
                Return False
            End If
        End Function

        ''' <summary>
        ''' この式を文字列として表現します。
        ''' </summary>
        ''' <returns>この式の文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me._target.ToString()
        End Function

    End Class

End Namespace
