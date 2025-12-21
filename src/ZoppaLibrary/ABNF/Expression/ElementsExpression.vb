Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 要素式。
    ''' elements = alternation *WSP
    ''' </summary>
    NotInheritable Class ElementsExpression
        Implements IExpression

        ''' <summary>
        ''' マッチングを行います。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <returns>マッチ結果。
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim alter = ABNFAlterExpr().Match(tr)
            ABNFSpaceExpr().Match(tr)
            Return alter
        End Function

    End Class

End Namespace
