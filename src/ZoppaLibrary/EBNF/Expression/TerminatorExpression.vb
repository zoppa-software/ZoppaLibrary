Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 終端記号（セミコロン ';' またはピリオド '.'）を表す式。
    ''' </summary>
    ''' <remarks>
    ''' マッチした場合に対応する式の範囲を返します。
    ''' terminator = ";" | "." ;
    ''' </remarks>
    NotInheritable Class TerminatorExpression
        Implements IExpression

        ''' <summary>
        ''' カスタムテキストリーダーの現在位置から1文字読み取り、
        ''' その文字が終端記号(';', '.')であればマッチした範囲を返します。
        ''' </summary>
        ''' <param name="tr">解析対象の <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は終端文字を含む <see cref="ExpressionRange"/>、
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </returns>
        ''' <remarks>
        ''' IExpression.Match の実装です。読み取り位置は <see cref="IPositionAdjustReader.Position"/> で扱われます。
        ''' </remarks>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim c = tr.Peek()
            If c = AscW(";"c) OrElse c = AscW("."c) Then
                tr.Read()
                Return New ExpressionRange(Me, tr, tr.Position - 1, tr.Position, ExpressionRange.EmptyRanges)
            End If
            Return ExpressionRange.Invalid
        End Function

    End Class

End Namespace
