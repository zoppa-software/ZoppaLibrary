Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' ルール式を表します。
    ''' rule = lhs , S , "=" , S , rhs , S , terminator ;
    ''' </summary>
    NotInheritable Class RuleExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' ルール式にマッチすれば
        ''' マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' rule = lhs , S , "=" , S , rhs , S , terminator ;
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position

            ' 左辺式にマッチするか試みる
            Dim lmth = LhsExpr.Match(tr)
            If Not lmth.Enable Then
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白を読み進める
            SpaceExpr.Match(tr)

            ' 等号記号を確認する
            Dim c = tr.Peek()
            If c = AscW("="c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白を読み進める
            SpaceExpr.Match(tr)

            ' 右辺式にマッチするか試みる
            Dim rmth = RhsExpr.Match(tr)
            If Not rmth.Enable Then
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白を読み進める
            SpaceExpr.Match(tr)

            ' 終端にマッチするか試みる
            If Not TerminatorExpr.Match(tr).Enable Then
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' マッチした範囲を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, New ExpressionRange() {lmth, rmth})
        End Function

    End Class

End Namespace
