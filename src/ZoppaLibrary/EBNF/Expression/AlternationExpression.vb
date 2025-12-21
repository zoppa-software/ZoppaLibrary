Option Explicit On
Option Strict On

Imports ZoppaLibrary.ABNF.NumValExpression
Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 縦棒区切りのカンマ区切りの式を表します。
    ''' alternation = ( S , concatenation , S , "|" ? ) + ;
    ''' </summary>
    NotInheritable Class AlternationExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 縦棒区切りのカンマ区切りの式にマッチすれば
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

            ' 最初のブロックにマッチするか試みる
            Dim mth = Me.BlockMatch(tr)
            If mth.Enable Then
                mths.Add(mth)
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If


            '// 最初の式を取得
            'ExpressionRange concatRange = ExpressionDefines.getConcatExpr().match(accesser);
            'If (concatRange.isEnable()) Then {
            '    ranges.add(concatRange);
            '}
            'Else {
            '    mark.restore();
            '    Return ExpressionRange.getInvalid();
            '}   

            '// 以降の選択する式を取得
            'While (accesser.peek()!= -1) {
            '    IByteAccesser.IPosition nextmark = accesser.mark();

            '    // コメントまたは空白
            '    ExpressionDefines.getCommentWspExpr().match(accesser);

            '    // '/' がなければ終了する
            '    int b = accesser.peek();
            '    If (b == '/') {
            '        accesser.read();
            '    }
            '    Else {
            '        nextmark.restore();
            '        break;
            '    }

            '    // コメントまたは空白
            '    ExpressionDefines.getCommentWspExpr().match(accesser);

            '    // 次の式を取得
            '    ExpressionRange nextRange = ExpressionDefines.getConcatExpr().match(accesser);
            '    If (nextRange.isEnable()) Then {
            '        ranges.add(nextRange);
            '    }
            '    Else {
            '        nextmark.restore();
            '        break;
            '    }
            '}

            '// マッチ結果を返す
            'Return New ExpressionRange(this, accesser.span(start, accesser.getPosition()), ranges);



            ' 1つ以上のブロックにマッチするか試みる
            Do While True
                Dim nextSnap = tr.MemoryPosition()

                ' 縦棒があれば読み進める
                If tr.Peek() = AscW("|") Then
                    tr.Read()
                Else
                    Exit Do
                End If

                ' ブロックにマッチするか試みる
                mth = Me.BlockMatch(tr)
                If mth.Enable Then
                    mths.Add(mth)
                Else
                    nextSnap.Restore()
                    Exit Do
                End If
            Loop

            ' マッチした範囲を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, mths.ToArray())
        End Function

        ''' <summary>
        ''' ブロックにマッチするか試みます。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Private Function BlockMatch(tr As IPositionAdjustReader) As ExpressionRange
            Dim snap = tr.MemoryPosition()

            ' 空白を読み進める
            SpaceExpr.Match(tr)

            ' 式にマッチするか試みる
            Dim mth = ConcatenationExpr.Match(tr)
            If Not mth.Enable Then
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白を読み進める
            SpaceExpr.Match(tr)

            Return mth
        End Function

    End Class

End Namespace
