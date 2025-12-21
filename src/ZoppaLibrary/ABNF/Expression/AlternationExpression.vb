Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 選択式。
    ''' alternation = concatenation *(*c-wsp "/" *c-wsp concatenation)
    ''' </summary>
    NotInheritable Class AlternationExpression
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
            Dim ranges As New List(Of ExpressionRange)()

            ' 最初の式を取得
            Dim concatRange = ABNFConcatExpr().Match(tr)
            If concatRange.Enable Then
                ranges.Add(concatRange)
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 以降の選択する式を取得
            Do While tr.Peek() <> -1
                Dim nextSnap = tr.MemoryPosition()

                ' コメントまたは空白
                ABNFCommentWspExpr().Match(tr)

                ' '/' がなければ終了する
                If tr.Peek() = AscW("/") Then
                    tr.Read()
                Else
                    nextSnap.Restore()
                    Exit Do
                End If

                ' コメントまたは空白
                ABNFCommentWspExpr().Match(tr)

                ' 次の式を取得
                concatRange = ABNFConcatExpr().Match(tr)
                If concatRange.Enable Then
                    ranges.Add(concatRange)
                Else
                    nextSnap.Restore()
                    Exit Do
                End If
            Loop

            ' マッチした範囲を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, ranges.ToArray())
        End Function

    End Class

End Namespace
