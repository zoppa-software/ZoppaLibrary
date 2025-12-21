Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' コメントと改行を表します。
    ''' c-nl = comment CRLF
    ''' </summary>
    NotInheritable Class CommentNewLineExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' コメントと改行にマッチすれば
        ''' マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim snap = tr.MemoryPosition()

            ' コメントを読み飛ばす
            Dim commentRange = ABNFCommentExpr().Match(tr)
            If commentRange.Enable Then
                Dim crlfRange = ABNFCrLfExpr().Match(tr)
                If crlfRange.Enable Then
                    Return commentRange
                End If
            End If

            ' 失敗
            snap.Restore()
            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
