Option Explicit On
Option Strict On

Imports System.Text

Namespace Parser

    ''' <summary>
    ''' コメント式を表します。
    ''' specialSeq = "?" , ( character | S ) * , "?" ;
    ''' </summary>
    Public NotInheritable Class SpecialSeqExpression
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
            If start1Char = AscW("?"c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' エスケープシーケンスを格納するバッファ
            Dim sb As New StringBuilder()
            Dim innerStartPos = tr.Position
            Dim esc = False

            ' コメントの内容を読み進める
            Do While True
                Dim c = tr.Peek()
                If c = AscW("?"c) Then
                    ' 終了の引用符にマッチした場合は読み進める
                    tr.Read()

                    ' エスケープ処理に対応
                    Dim innerExpr = If(esc,
                        New ExpressionRange(IdentifierExpr, New PositionAdjustString(sb.ToString()), 0, sb.Length, ExpressionRange.EmptyRanges),
                        New ExpressionRange(IdentifierExpr, tr, innerStartPos, tr.Position - 1, ExpressionRange.EmptyRanges)
                    )
                    Return New ExpressionRange(Me, tr, startPos, tr.Position,
                        New ExpressionRange() {innerExpr}
                    )

                ElseIf c = AscW("\"c) Then
                    ' エスケープシーケンスを読み進める
                    tr.Read()
                    Dim ec = tr.Peek()
                    If ec = AscW("\"c) OrElse ec = AscW("?"c) Then
                        sb.Append(ChrW(ec))
                        esc = True
                        tr.Read()
                    End If

                ElseIf c = -1 Then
                    Exit Do

                Else
                    sb.Append(ChrW(c))
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

