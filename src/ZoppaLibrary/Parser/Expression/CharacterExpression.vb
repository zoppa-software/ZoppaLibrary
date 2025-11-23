Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 1文字の任意の文字にマッチする式を表します。
    ''' character = letter | digit | symbol | "_" | " " ;
    ''' </summary>
    Public NotInheritable Class CharacterExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が任意の1文字であれば
        ''' その1文字を読み進め、マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim snap = tr.MemoryPosition()

            ' それぞれの式でマッチを試みる
            For Each expr In New IExpression() {LetterExpr(), DigitExpr(), SymbolExpr()}
                Dim range = expr.Match(tr)
                If range.Enable Then
                    Return range
                End If
            Next

            ' "_" または " " にマッチするか確認
            Dim c = tr.Peek()
            If c = AscW("_"c) OrElse c = AscW(" "c) OrElse c = AscW("\"c) Then
                tr.Read()
                Return New ExpressionRange(Me, tr, tr.Position - 1, tr.Position, ExpressionRange.EmptyRanges)
            End If

            ' いずれにもマッチしなかった場合は元の位置に戻す
            snap.Restore()
            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
