Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 数値値を表します。
    ''' num-val = "%" (bin-val / dec-val / hex-val)
    ''' bin-val = "b" 1*BIT [ 1*("." 1*BIT) / ("-" 1*BIT) ]
    ''' dec-val = "d" 1*DIGIT [ 1*("." 1*DIGIT) / ("-" 1*DIGIT) ]
    ''' hex-val = "x" 1*HEXDIG [ 1*("." 1*HEXDIG) / ("-" 1*HEXDIG) ]
    ''' </summary>
    NotInheritable Class NumValExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 数値値にマッチすれば
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
            Dim mths As New List(Of ExpressionRange)()

            ' 開始文字は "%"
            If tr.Peek() = AscW("%"c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 続く文字は b / d / x のいずれか
            Dim expr = ExpressionRange.Invalid
            Select Case tr.Peek()
                Case AscW("b"c)
                    expr = MatchValue(tr, {
                        AscW("0"c), AscW("1"c)
                    })
                Case AscW("d"c)
                    expr = MatchValue(tr, {
                        AscW("0"c), AscW("1"c), AscW("2"c), AscW("3"c), AscW("4"c),
                        AscW("5"c), AscW("6"c), AscW("7"c), AscW("8"c), AscW("9"c)
                    })
                Case AscW("x"c)
                    expr = MatchValue(tr, {
                        AscW("0"c), AscW("1"c), AscW("2"c), AscW("3"c), AscW("4"c),
                        AscW("5"c), AscW("6"c), AscW("7"c), AscW("8"c), AscW("9"c),
                        AscW("A"c), AscW("B"c), AscW("C"c), AscW("D"c), AscW("E"c), AscW("F"c),
                        AscW("a"c), AscW("b"c), AscW("c"c), AscW("d"c), AscW("e"c), AscW("f"c)
                    })
                Case Else
                    snap.Restore()
                    Return ExpressionRange.Invalid
            End Select

            ' マッチ結果を返す
            If expr.Enable Then
                mths.Add(expr)
                Return New ExpressionRange(Me, tr, startPos, tr.Position, mths.ToArray())
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If
        End Function

        ''' <summary>
        ''' 指定された数値セットに基づいてマッチングを行います。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="numbers">マッチング対象の数値セット。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Private Function MatchValue(tr As IPositionAdjustReader, numbers() As Integer) As ExpressionRange
            Dim startPos = tr.Position
            Dim mths As New List(Of ExpressionRange)()

            ' b / d / x を読み飛ばす
            tr.Read()

            ' マッチ用の数値セットを作成
            Dim numSet = New Boolean(255) {}
            For Each n As Integer In numbers
                numSet(n) = True
            Next

            Dim enabled = False
            Dim b As Integer = 0

            ' 先頭の数字を取得
            Dim fst = tr.Position
            Do While True
                b = tr.Peek()
                If b >= 0 AndAlso numSet(b) Then
                    tr.Read()
                    enabled = True
                Else
                    Exit Do
                End If
            Loop

            ' 連結または範囲指定の取得
            If enabled Then
                mths.Add(New ExpressionRange(Me, tr, fst, tr.Position, ExpressionRange.EmptyRanges))

                Select Case tr.Peek()
                    Case AscW("."c)
                        ' 連結取得
                        tr.Read()

                        ' 続く数字を取得
                        Do While True
                            Dim nst = tr.Position
                            Dim nena = False
                            Do While True
                                b = tr.Peek()
                                If b >= 0 AndAlso numSet(b) Then
                                    tr.Read()
                                    enabled = True
                                    nena = True
                                Else
                                    Exit Do
                                End If
                            Loop
                            If nena Then
                                mths.Add(New ExpressionRange(ABNFNumValRangeExpr, tr, nst, tr.Position, ExpressionRange.EmptyRanges))
                            End If

                            ' 次のピリオドを確認
                            If tr.Peek() = AscW("."c) Then
                                tr.Read()
                            Else
                                Exit Do
                            End If
                        Loop

                    Case AscW("-"c)
                        ' 範囲取得
                        tr.Read()

                        ' 続く数字を取得
                        Dim nst = tr.Position
                        Dim nena = False
                        Do While True
                            b = tr.Peek()
                            If b >= 0 AndAlso numSet(b) Then
                                tr.Read()
                                enabled = True
                                nena = True
                            Else
                                Exit Do
                            End If
                        Loop
                        If nena Then
                            mths.Add(New ExpressionRange(ABNFNumValConcatExpr, tr, nst, tr.Position, ExpressionRange.EmptyRanges))
                        End If
                End Select
            End If

            ' マッチ結果を返す
            Return If(
                enabled,
                New ExpressionRange(Me, tr, startPos, tr.Position, mths.ToArray()),
                ExpressionRange.Invalid
            )
        End Function

        ''' <summary>
        ''' 範囲指定を表します。
        ''' </summary>
        Public NotInheritable Class Range
            Implements IExpression

            ''' <summary>
            ''' マッチ判定を行います（未使用）。
            ''' </summary>
            ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
            ''' <returns>無効値。</returns>
            Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
                Return ExpressionRange.Invalid
            End Function

        End Class

        ''' <summary>
        ''' 連結を表します。
        ''' </summary>
        Public NotInheritable Class Concat
            Implements IExpression

            ''' <summary>
            ''' マッチ判定を行います（未使用）。
            ''' </summary>
            ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
            ''' <returns>無効値。</returns>
            Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
                Return ExpressionRange.Invalid
            End Function

        End Class

    End Class

End Namespace
