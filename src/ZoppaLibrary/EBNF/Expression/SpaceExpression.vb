Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 空白文字（スペース、タブ、改行など）にマッチする式を表します。
    ''' S = { " " | "\n" | "\t" | "\r" | "\f" | "\b" } ;
    ''' </summary>
    NotInheritable Class SpaceExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が空白文字であれば
        ''' その1文字を読み進め、マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim startPos = tr.Position
            Do While True
                Dim c = tr.Peek()
                Select Case c
                    Case AscW(" "c), AscW(vbLf), AscW(vbTab), AscW(vbCr), AscW(vbFormFeed), AscW(vbBack)
                        tr.Read()
                    Case Else
                        Exit Do
                End Select
            Loop
            Return New ExpressionRange(Me, tr, startPos, tr.Position, ExpressionRange.EmptyRanges)
        End Function

    End Class

End Namespace
