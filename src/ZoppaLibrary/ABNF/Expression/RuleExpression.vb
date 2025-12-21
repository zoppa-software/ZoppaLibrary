Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ルール名式。
    ''' rule = rulename defined-as elements c-nl
    ''' </summary>
    NotInheritable Class RuleExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' ルール名式にマッチすれば
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

            ' ルール名を取得
            Dim ruleNameRange = ABNFRuleNameExpr.Match(tr)
            If ruleNameRange.Enable Then
                ranges.Add(ruleNameRange)
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' defined-as
            ' 空白読み捨て
            ABNFCommentWspExpr.Match(tr)

            ' '=' を取得
            Dim equalChar = tr.Peek()
            If equalChar = AscW("="c) Then
                tr.Read()
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' '/' を取得
            Dim orChar = tr.Peek()
            If orChar = AscW("/"c) Then
                tr.Read()
                ranges.Add(New ExpressionRange(ABNFAlterExpr, tr, tr.Position - 1, tr.Position, ranges.ToArray()))
            End If

            ' 空白読み捨て
            ABNFCommentWspExpr.Match(tr)

            ' 右辺の定義を取得
            Dim elementsRange = ABNFElementsExpr.Match(tr)
            If elementsRange.Enable Then
                ranges.Add(elementsRange)
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' コメントと改行をスキップする
            ABNFCommentNlExpr.Match(tr)

            Return New ExpressionRange(Me, tr, startPos, tr.Position, ranges.ToArray())
        End Function

    End Class

End Namespace
