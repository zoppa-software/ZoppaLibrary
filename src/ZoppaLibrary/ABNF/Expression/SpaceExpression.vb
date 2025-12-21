Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 空白を表します。
    ''' WSP = SP / HTAB
    ''' </summary>
    NotInheritable Class SpaceExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 空白にマッチすれば
        ''' マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim startPos = tr.Position
            Dim matched = False

            ' SP または HTAB にマッチするか試みる
            Do While True
                Dim ch = tr.Peek()
                If ch = AscW(" "c) OrElse ch = AscW(vbTab) Then
                    tr.Read()
                    matched = True
                Else
                    Exit Do
                End If
            Loop

            ' マッチ結果を返す
            If matched Then
                Return New ExpressionRange(Me, tr, startPos, tr.Position, ExpressionRange.EmptyRanges)
            Else
                Return ExpressionRange.Invalid
            End If
        End Function

    End Class

End Namespace
