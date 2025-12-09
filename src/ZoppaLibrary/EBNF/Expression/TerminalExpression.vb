Option Explicit On
Option Strict On

Imports System.Text

Namespace EBNF

    ''' <summary>
    ''' 終端記号にマッチする式を表します。
    ''' terminal = "'" , character - "'" , { character - "'" } , "'"
    '''         | '"' , character - '"' , { character - '"' } , '"' ;
    ''' </summary>
    NotInheritable Class TerminalExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が終端記号にマッチすれば
        ''' マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim snap = tr.MemoryPosition()

            Dim startPos = tr.Position
            Dim ranges As New List(Of ExpressionRange)()

            Dim quoteChar = tr.Peek()
            If quoteChar = AscW("'"c) OrElse quoteChar = AscW(""""c) Then
                ' 開始の引用符を読み進める
                tr.Read()

                ' 最初の文字を確認
                Dim sb As New StringBuilder()
                Dim innerStartPos = tr.Position
                Dim innerEndPos = 0
                Dim esc = False

                Do While True
                    Select Case tr.Peek()
                        Case -1
                            ' 入力の終端に達した場合は終了
                            snap.Restore()
                            Return ExpressionRange.Invalid

                        Case quoteChar
                            ' 空の引用符は無効
                            innerEndPos = tr.Position
                            If innerStartPos >= innerEndPos Then
                                snap.Restore()
                                Return ExpressionRange.Invalid
                            End If
                            Exit Do

                        Case AscW("\"c)
                            ' エスケープシーケンスを読み進める
                            esc = True
                            tr.Read()

                            Dim ic = ChrW(tr.Peek())
                            Select Case ic
                                Case "n"c
                                    sb.Append(vbLf)
                                Case "r"c
                                    sb.Append(vbCr)
                                Case "t"c
                                    sb.Append(vbTab)
                                Case "f"c
                                    sb.Append(vbFormFeed)
                                Case "b"c
                                    sb.Append(vbBack)
                                Case "\"c
                                    sb.Append("\"c)
                                Case Else
                                    Throw New InvalidCastException($"不明なエスケープシーケンスです: \{ic}")
                            End Select
                            tr.Read()

                        Case Else
                            sb.Append(ChrW(tr.Read()))
                    End Select
                Loop

                ' 終了の引用符を読み進める
                Dim nc = tr.Peek()
                If nc = quoteChar Then
                    ' 引用符にマッチした場合は読み進める
                    innerEndPos = tr.Position
                    tr.Read()
                Else
                    ' 引用符がマッチしなかった場合は終了
                    snap.Restore()
                    Return ExpressionRange.Invalid
                End If

                ' マッチした範囲を返す
                Dim innerExpr = If(esc,
                    New ExpressionRange(CharacterExpr, New PositionAdjustString(sb.ToString()), 0, sb.Length, ExpressionRange.EmptyRanges),
                    New ExpressionRange(CharacterExpr, tr, innerStartPos, innerEndPos, ExpressionRange.EmptyRanges)
                )
                Return New ExpressionRange(Me, tr, startPos, tr.Position,
                    New ExpressionRange() {innerExpr}
                )
            End If

            snap.Restore()
            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
