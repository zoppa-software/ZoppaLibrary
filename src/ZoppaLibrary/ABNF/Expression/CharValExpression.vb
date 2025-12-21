Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 文字値を表します。
    ''' char-val = DQUOTE *(%x20-21 / %x23-7E) DQUOTE
    ''' </summary>
    NotInheritable Class CharValExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 文字値にマッチすれば
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

            ' 開始文字はDQUOTE
            If tr.Peek() = AscW(""""c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            Dim inStart = tr.Position

            ' 続く文字は%20-21 / %23-7E 0回以上
            Do While True
                Dim ch = tr.Peek()
                If (ch >= &H20 AndAlso ch <= &H21) OrElse
                   (ch >= &H23 AndAlso ch <= &H7E) Then
                    tr.Read()
                Else
                    Exit Do
                End If
            Loop

            Dim inEnd = tr.Position

            ' 終了文字はDQUOTE
            If tr.Peek() = AscW(""""c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' マッチした範囲を返す
            ranges.Add(New ExpressionRange(Me, tr, inStart, inEnd, ExpressionRange.EmptyRanges))
            Return New ExpressionRange(Me, tr, startPos, tr.Position, ranges)
        End Function

    End Class

End Namespace
