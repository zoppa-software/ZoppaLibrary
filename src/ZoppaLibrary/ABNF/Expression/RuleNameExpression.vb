Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ルール名を表します。
    ''' rulename = ALPHA *(ALPHA / DIGIT / "-" / "_")
    ''' </summary>
    NotInheritable Class RuleNameExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' ルール名にマッチすれば
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

            ' 最初の文字はALPHA
            Dim firstChar = tr.Peek()
            If (firstChar >= AscW("A") AndAlso firstChar <= AscW("Z")) OrElse
               (firstChar >= AscW("a") AndAlso firstChar <= AscW("z")) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 続く文字はALPHA / DIGIT / "-" / "_" 0回以上
            Do While True
                Dim ch = tr.Peek()
                If (ch >= AscW("A"c) AndAlso ch <= AscW("Z"c)) OrElse
                   (ch >= AscW("a"c) AndAlso ch <= AscW("z"c)) OrElse
                   (ch >= AscW("0"c) AndAlso ch <= AscW("9"c)) OrElse
                   ch = AscW("-"c) OrElse
                   ch = AscW("_"c) Then
                    tr.Read()
                Else
                    Exit Do
                End If
            Loop

            ' マッチした範囲を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, ExpressionRange.EmptyRanges)
        End Function

    End Class

End Namespace
