Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 1文字の数字 (0-9) にマッチする式を表します。
    ''' digit = "0" | "1" | "2" | "3" | "4" | "5" | "6" | "7" | "8" | "9" ;
    ''' </summary>
    NotInheritable Class DigitExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が数字であれば
        ''' その1文字を読み進め、マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim c = tr.Peek()
            If c >= AscW("0"c) AndAlso c <= AscW("9"c) Then
                tr.Read()
                Return New ExpressionRange(Me, tr, tr.Position - 1, tr.Position, ExpressionRange.EmptyRanges)
            End If
            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
