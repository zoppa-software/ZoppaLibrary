Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 反復式。
    ''' [1*DIGIT / (*DIGIT "*" *DIGIT)] (rulename / group / option / char-val / num-val / prose-val)
    ''' </summary>
    NotInheritable Class RepetitionExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 反復式にマッチすれば
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

            ' 反復回数を取得
            Dim repeatRange = MatchRepeat(tr)
            If repeatRange.Enable Then
                ranges.Add(repeatRange)
            End If

            ' 続く式を取得
            Dim expr = SelectExpression(tr)
            If expr.Enable Then
                ranges.Add(expr)
                Return New ExpressionRange(Me, tr, startPos, tr.Position, ranges.ToArray())
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If
        End Function

        ''' <summary>
        ''' 反復回数部分にマッチします。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>反復回数。</returns>
        Private Function MatchRepeat(tr As IPositionAdjustReader) As ExpressionRange
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim ranges = {ExpressionRange.Invalid, ExpressionRange.Invalid}
            Dim enable = False

            ' 先頭の数字を取得
            Dim first = tr.Position
            Dim fena = False
            Do While True
                Dim ch = tr.Peek()
                If ch >= AscW("0"c) AndAlso ch <= AscW("9"c) Then
                    tr.Read()
                    enable = True
                    fena = True
                Else
                    Exit Do
                End If
            Loop
            If fena Then
                ranges(0) = New ExpressionRange(Me, tr, first, tr.Position, ExpressionRange.EmptyRanges)
            End If

            ' "*" があれば読み進める
            If tr.Peek() = AscW("*"c) Then
                tr.Read()

                Dim nest = tr.Position
                Dim nena = False
                enable = True

                ' 続く数字を取得
                Do While True
                    Dim ch = tr.Peek()
                    If ch >= AscW("0"c) AndAlso ch <= AscW("9"c) Then
                        tr.Read()
                        nena = True
                    Else
                        Exit Do
                    End If
                Loop
                If nena Then
                    ranges(1) = New ExpressionRange(Me, tr, nest, tr.Position, ExpressionRange.EmptyRanges)
                End If
            End If

            ' マッチ結果を返す
            If enable Then
                Return New ExpressionRange(Me, tr, startPos, tr.Position, ranges)
            Else
                ' マッチ失敗
                snap.Restore()
                Return ExpressionRange.Invalid
            End If
        End Function

        ''' <summary>
        ''' 続く式を選択してマッチします。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>マッチした式。</returns>
        Private Function SelectExpression(tr As IPositionAdjustReader) As ExpressionRange
            Dim expr = ABNFRuleNameExpr.Match(tr)
            If expr.Enable Then
                Return expr
            End If

            expr = ABNFGroupExpr().Match(tr)
            If expr.Enable Then
                Return expr
            End If

            expr = ABNFOptionExpr().Match(tr)
            If expr.Enable Then
                Return expr
            End If

            expr = ABNFCharValExpr.Match(tr)
            If expr.Enable Then
                Return expr
            End If

            expr = ABNFNumValExpr.Match(tr)
            If expr.Enable Then
                Return expr
            End If

            expr = ABNFProseValExpr.Match(tr)
            If expr.Enable Then
                Return expr
            End If

            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
