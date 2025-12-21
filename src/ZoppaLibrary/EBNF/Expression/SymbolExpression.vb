Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' シンボルを表す式。
    ''' symbol = "[" | "]" | "{" | "}" | "(" | ")" | "<" | ">"
    '''        | "'" | '"' | "=" | "|" | "." | "," | ";" | "-" 
    '''        | "+" | "*" | "?" | "\n" | "\t" | "\r" | "\f" | "\b" ;
    ''' </summary>
    NotInheritable Class SymbolExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字がシンボルであれば
        ''' その1文字を読み進め、マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim c = tr.Peek()
            Select Case c
                Case AscW("["c), AscW("]"c), AscW("{"c), AscW("}"c), AscW("("c), AscW(")"c),
                     AscW("<"c), AscW(">"c), AscW("'"c), AscW(""""c), AscW("="c), AscW("|"c),
                     AscW("."c), AscW(","c), AscW(";"c), AscW("-"c), AscW("+"c), AscW("*"c),
                     AscW("?"c), AscW(vbLf), AscW(vbTab), AscW(vbCr), AscW(vbFormFeed), AscW(vbBack), AscW("\"c)
                    tr.Read()
                    Return New ExpressionRange(Me, tr, tr.Position - 1, tr.Position, ExpressionRange.EmptyRanges)
                Case Else
                    Return ExpressionRange.Invalid
            End Select
        End Function

    End Class

End Namespace
