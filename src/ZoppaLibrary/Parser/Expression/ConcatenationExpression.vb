Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' カンマ区切りの式を表します。
    ''' concatenation = ( S , factor , S , "," ? ) + ;
    ''' </summary>
    Public NotInheritable Class ConcatenationExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' カンマ区切りの繰り返し記号付きの式にマッチすれば
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
            Dim enable = False

            Dim mths As New List(Of ExpressionRange)()

            ' 1つ以上のブロックにマッチするか試みる
            Do While True
                ' ブロックにマッチするか試みる
                Dim mth = Me.BlockMatch(tr)
                If Not mth.Enable Then
                    Exit Do
                End If

                mths.Add(mth)
                enable = True

                ' カンマがあれば読み進める
                If tr.Peek() = AscW(",") Then
                    tr.Read()
                Else
                    Exit Do
                End If
            Loop

            ' マッチした範囲を返す
            If enable Then
                Return New ExpressionRange(Me, tr, startPos, tr.Position, mths.ToArray())
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If
        End Function

        ''' <summary>
        ''' 1つのブロックにマッチするか試みます。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Private Function BlockMatch(tr As IPositionAdjustReader) As ExpressionRange
            Dim snap = tr.MemoryPosition()

            ' 空白を読み進める
            SpaceExpr.Match(tr)

            ' 式にマッチするか試みる
            Dim mth = FactorExpr.Match(tr)
            If Not mth.Enable Then
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白を読み進める
            SpaceExpr.Match(tr)

            Return mth
        End Function

    End Class

End Namespace
