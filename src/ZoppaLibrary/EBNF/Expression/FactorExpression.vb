Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 繰り返し記号付きの式を表します。
    ''' factor = term , S , "?"
    '''        | term , S , "*"
    '''        | term , S , "+"
    '''        | term , S , "-" , S , term
    '''        | term , S ;
    ''' </summary>
    NotInheritable Class FactorExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 繰り返し記号付きの式にマッチすれば
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

            Dim mths As New List(Of ExpressionRange)()

            ' 式にマッチするか試みる
            Dim lmth = TermExpr.Match(tr)
            If lmth.Enable Then
                mths.Add(lmth)
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白を読み進める
            SpaceExpr.Match(tr)

            ' 繰り返し記号を確認する
            Dim c = tr.Peek()
            If c = AscW("*"c) OrElse c = AscW("+"c) OrElse c = AscW("?"c) Then
                ' 繰り返し記号があれば読み進める
                tr.Read()
                mths.Add(New ExpressionRange(New CharacterExpression(), tr, tr.Position - 1, tr.Position, ExpressionRange.EmptyRanges))
            ElseIf c = AscW("-"c) Then
                tr.Read()
                mths.Add(New ExpressionRange(New CharacterExpression(), tr, tr.Position - 1, tr.Position, ExpressionRange.EmptyRanges))

                ' 空白を読み進める
                SpaceExpr.Match(tr)

                ' 式にマッチするか試みる
                Dim rmth = TermExpr.Match(tr)
                If rmth.Enable Then
                    mths.Add(rmth)
                Else
                    snap.Restore()
                    Return ExpressionRange.Invalid
                End If
            End If

            ' マッチした範囲を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, mths.ToArray())
        End Function

    End Class

End Namespace
