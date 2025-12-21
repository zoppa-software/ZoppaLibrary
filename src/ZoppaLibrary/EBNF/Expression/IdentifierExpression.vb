Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 識別子を表す式。
    ''' identifier = letter , { letter | digit | "_" } ;
    ''' </summary>
    NotInheritable Class IdentifierExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が識別子にマッチすれば
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

            ' 最初の文字は letter にマッチする必要がある
            Dim range = LetterExpr.Match(tr)
            If Not range.Enable Then
                Return ExpressionRange.Invalid
            End If

            ' 続く文字は letter, digit, "_" のいずれかにマッチする
            Do While True
                Dim currentPos = tr.Position

                range = LetterExpr.Match(tr)
                If Not range.Enable Then
                    range = DigitExpr.Match(tr)
                    If Not range.Enable Then
                        Dim c = tr.Peek()
                        If c = AscW("_"c) Then
                            tr.Read()
                        Else
                            ' どれにもマッチしなかった場合はループを抜ける
                            Exit Do
                        End If
                    End If
                End If
            Loop

            ' マッチした範囲を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, ExpressionRange.EmptyRanges)
        End Function

    End Class

End Namespace
