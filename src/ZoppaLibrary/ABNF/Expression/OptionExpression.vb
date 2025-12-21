Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 選択式。
    ''' option = "[" *c-wsp alternation *c-wsp "]"
    ''' </summary>
    NotInheritable Class OptionExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 選択式にマッチすれば
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

            ' '['
            Dim firstBracket = tr.Peek()
            If firstBracket = AscW("["c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白読み捨て
            ABNFCommentWspExpr.Match(tr)

            ' 選択式をマッチングする
            Dim alterRange = ABNFAlterExpr().Match(tr)
            If alterRange.Enable Then
                mths.Add(alterRange)
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白読み捨て
            ABNFCommentWspExpr.Match(tr)

            ' ']'
            Dim endBracket = tr.Peek()
            If endBracket = AscW("]"c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' オプション式のマッチ結果を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, mths.ToArray())
        End Function

    End Class

End Namespace
