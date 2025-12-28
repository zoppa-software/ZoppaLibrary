Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' コメントと空白を表します。
    ''' c-wsp = WSP / (c-nl WSP)
    ''' </summary>
    NotInheritable Class CommentWhiteSpaceExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' コメントと空白にマッチすれば
        ''' マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            ' コメントと空白の組み合わせを試す
            ABNFSpaceExpr().Match(tr)
            Dim comment = ABNFCommentNlExpr().Match(tr)
            ABNFSpaceExpr().Match(tr)
            If comment.Enable Then
                Return comment
            End If
            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
