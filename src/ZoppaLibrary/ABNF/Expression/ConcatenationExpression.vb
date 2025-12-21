Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 連結式。
    ''' concatenation = repetition *(1*c-wsp repetition)
    ''' </summary>
    NotInheritable Class ConcatenationExpression
        Implements IExpression

        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim snap = tr.MemoryPosition()
            Dim startPos = tr.Position
            Dim mths As New List(Of ExpressionRange)()

            ' 最初の式を取得
            Dim repeatRange = ABNFRepeatExpr.Match(tr)
            If repeatRange.Enable Then
                mths.Add(repeatRange)
            Else
                snap.Restore()
                Return ExpressionRange.Invalid
            End If

            ' 以降の選択する式を取得
            Do While tr.Peek() <> -1
                Dim nextSnap = tr.MemoryPosition()

                ' コメントまたは空白
                Dim comExpr = ABNFCommentWspExpr.Match(tr)
                If Not comExpr.Enable Then
                    nextSnap.Restore()
                    Exit Do
                End If

                ' 次の式を取得
                Dim nextRange = ABNFRepeatExpr.Match(tr)
                If nextRange.Enable Then
                    mths.Add(nextRange)
                Else
                    nextSnap.Restore()
                    Exit Do
                End If
            Loop

            ' マッチ結果を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, mths)
        End Function

    End Class

End Namespace
