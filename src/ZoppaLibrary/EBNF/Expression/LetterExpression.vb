Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 1文字の英字 (A-Z, a-z) にマッチする式を表します。
    ''' letter = "A" | "B" | "C" | "D" | "E" | "F" | "G"
    '''       | "H" | "I" | "J" | "K" | "L" | "M" | "N"
    '''       | "O" | "P" | "Q" | "R" | "S" | "T" | "U"
    '''       | "V" | "W" | "X" | "Y" | "Z" | "a" | "b"
    '''       | "c" | "d" | "e" | "f" | "g" | "h" | "i"
    '''       | "j" | "k" | "l" | "m" | "n" | "o" | "p"
    '''       | "q" | "r" | "s" | "t" | "u" | "v" | "w"
    '''       | "x" | "y" | "z" ;
    ''' </summary>
    NotInheritable Class LetterExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が英字であれば
        ''' その1文字を読み進め、マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim c = tr.Peek()
            If (c >= AscW("A"c) AndAlso c <= AscW("Z"c)) OrElse
               (c >= AscW("a"c) AndAlso c <= AscW("z"c)) Then
                tr.Read()
                Return New ExpressionRange(Me, tr, tr.Position - 1, tr.Position, ExpressionRange.EmptyRanges)
            End If
            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
