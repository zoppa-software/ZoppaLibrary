Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' コメントを表します。
    ''' comment = ";" *(WSP / VCHAR) CRLF
    ''' </summary>
    NotInheritable Class CommentExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' コメントにマッチすれば
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

            ' 開始文字は";"
            If tr.Peek() = AscW(";"c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 改行文字が現れるまで読み飛ばす
            Do While tr.Peek() <> -1
                Dim ch = tr.Peek()
                If ch <> &HD AndAlso ch <> &HA Then
                    tr.Read()
                Else
                    Exit Do
                End If
            Loop

            ' 改行を読み飛ばす
            Dim endPos = tr.Position
            ABNFCrLfExpr().Match(tr)

            Return New ExpressionRange(Me, tr, startPos, endPos, ExpressionRange.EmptyRanges)
        End Function

    End Class

End Namespace