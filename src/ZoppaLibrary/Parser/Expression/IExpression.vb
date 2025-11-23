Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 式を表すインターフェイス。
    ''' </summary>
    Public Interface IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が識別子にマッチすれば
        ''' マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Function Match(tr As IPositionAdjustReader) As ExpressionRange

    End Interface

End Namespace
