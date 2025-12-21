Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 終端記号、識別子、または括弧で囲まれた式にマッチする式を表します。
    ''' term = "(" , S , rhs , S , ")"
    '''      | "[" , S , rhs , S , "]"
    '''      | "{" , S , rhs , S , "}"
    '''      | terminal
    '''      | identifier ;
    ''' </summary>
    NotInheritable Class TermExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' 終端記号、識別子、または括弧で囲まれた式にマッチすれば
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

            Select Case tr.Peek()
                Case AscW("("c)
                    Return MatchBracketTerm(tr, AscW(")"c), startPos, snap)
                Case AscW("["c)
                    Return MatchBracketTerm(tr, AscW("]"c), startPos, snap)
                Case AscW("{"c)
                    Return MatchBracketTerm(tr, AscW("}"c), startPos, snap)
                Case Else
                    ' 終端、識別子のいずれかにマッチを試みる
                    For Each expr In New IExpression() {TerminalExpr(), IdentifierExpr(), SpecialSeqExpr()}
                        Dim mth = expr.Match(tr)
                        If mth.Enable Then
                            Return New ExpressionRange(Me, tr, startPos, tr.Position, New ExpressionRange() {mth})
                        End If
                    Next
            End Select

            ' いずれにもマッチしなかった場合は元の位置に戻す
            snap.Restore()
            Return ExpressionRange.Invalid
        End Function

        ''' <summary>
        ''' 括弧で囲まれた式にマッチします。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="bracketChar">終了括弧の文字コード。</param>
        ''' <param name="startPos">マッチ開始位置。</param>
        ''' <param name="snap">スナップショット。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Private Function MatchBracketTerm(tr As IPositionAdjustReader,
                                          bracketChar As Integer,
                                          startPos As Integer,
                                          snap As IPositionAdjustReader.IPosition) As ExpressionRange
            ' 開始の括弧を読み進める
            tr.Read()

            ' 空白を確認
            SpaceExpr.Match(tr)

            ' 式のマッチ判定
            Dim mth = RhsExpr.Match(tr)
            If Not mth.Enable Then
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 空白を確認
            SpaceExpr.Match(tr)

            ' 終了の括弧を読み進める
            Dim nc = tr.Peek()
            If nc = bracketChar Then
                ' 括弧にマッチした場合は読み進める
                tr.Read()
            Else
                ' 括弧がマッチしなかった場合は終了
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' マッチした範囲を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, New ExpressionRange() {mth})
        End Function

    End Class

End Namespace
