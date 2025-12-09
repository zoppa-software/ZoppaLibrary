Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 文法全体にマッチする式を表します。
    ''' grammar = ( S , rule , S ) * ;
    ''' </summary>
    NotInheritable Class GrammarExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 文法全体にマッチすればマッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position

            Dim mths As New List(Of ExpressionRange)()

            Do While tr.Peek() <> -1
                ' 空白を読み進める
                SpaceExpr.Match(tr)

                ' ルール式、コメント式にマッチするか試みる
                Dim ruleRange = RuleExpr.Match(tr)
                If ruleRange.Enable Then
                    mths.Add(ruleRange)
                ElseIf Not CommentExpr.Match(tr).Enable Then
                    Return ExpressionRange.Invalid
                End If
#If DEBUG Then
                Debug.WriteLine($"Matched rule: {tr.Substring(ruleRange.[Start], ruleRange.[End] - ruleRange.[Start])}")
#End If

                ' 空白を読み進める
                SpaceExpr.Match(tr)
            Loop
            Return New ExpressionRange(Me, tr, startPos, tr.Position, mths.ToArray())
        End Function

    End Class

End Namespace
