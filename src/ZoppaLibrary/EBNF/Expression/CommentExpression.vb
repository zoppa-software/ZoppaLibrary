Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' コメント式を表します。
    ''' comment = "(*" , ( character | S ) * , "*)" ;
    ''' </summary>
    NotInheritable Class CommentExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' コメント式にマッチすればマッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim snap = tr.MemoryPosition()

            Dim startPos = tr.Position

            ' 開始の引用符を確認する
            Dim start1Char = tr.Peek()
            If start1Char = AscW("("c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            Dim start2Char = tr.Peek()
            If start2Char = AscW("*"c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' コメントの内容を読み進める
            Do While True
                Dim c = tr.Peek()
                If c = AscW("*"c) Then
                    ' 終了の引用符の可能性を確認する
                    Dim snapInner = tr.MemoryPosition()
                    tr.Read()
                    Dim nc = tr.Peek()
                    If nc = AscW(")"c) Then
                        ' 終了の引用符にマッチした場合は読み進める
                        tr.Read()
                        Return New ExpressionRange(Me, tr, startPos, tr.Position, ExpressionRange.EmptyRanges)
                    Else
                        ' 終了の引用符にマッチしなかった場合は元に戻す
                        snapInner.Restore()
                    End If
                ElseIf c = -1 Then
                    Exit Do
                End If

                ' 文字または空白を読み進める
                tr.Read()
            Loop

            ' 入力の終端に達した場合は終了
            snap.Restore()
            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
