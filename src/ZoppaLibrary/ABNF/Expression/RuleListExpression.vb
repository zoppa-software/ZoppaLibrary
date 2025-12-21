Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ルールリスト式。
    ''' rulelist = 1*( rule / (*WSP c-nl) )
    ''' </summary>
    NotInheritable Class RuleListExpression
        Implements IExpression

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字が
        ''' ルールリスト式にマッチすれば
        ''' マッチした範囲を <see cref="ExpressionRange"/> として返します。
        ''' マッチしない場合は <see cref="ExpressionRange.Invalid"/> を返します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <returns>
        ''' マッチした場合は開始位置と終了位置を持つ <see cref="ExpressionRange"/>。失敗時は <see cref="ExpressionRange.Invalid"/>.
        ''' </returns>
        Public Function Match(tr As IPositionAdjustReader) As ExpressionRange Implements IExpression.Match
            Dim startPos = tr.Position
            Dim ranges As New List(Of ExpressionRange)()
            Dim lists As New SortedDictionary(Of String, ExpressionRange)()
            Dim prevPos = -1

            Do While tr.Peek() <> -1
                ' ルール式をマッチングする
                Dim ruleRange = ABNFRuleExpr.Match(tr)
                If ruleRange.Enable Then
                    Dim key = ruleRange.SubRanges(0).ToString()
                    If Not lists.ContainsKey(key) Then
                        ranges.Add(ruleRange)
                        lists.Add(key, ruleRange)
                    ElseIf ruleRange.SubRanges.Count > 2 AndAlso
                           lists(key).SubRanges.Count = 2 Then
                        lists(key).GetRange(1).AddSubRanges(ruleRange.GetRange(2).SubRanges)
                    End If
                Else
                    ' SP / HTAB をスキップする
                    ABNFSpaceExpr.Match(tr)

                    ' コメント式をマッチングする
                    Dim commentRange = ABNFCommentNlExpr.Match(tr)
                    If commentRange.Enable Then
                        ranges.Add(commentRange)
                    End If
                End If

                ' 改行のみが続く場合はスキップする
                ABNFCrLfExpr.Match(tr)

                ' 進捗チェック
                If prevPos = tr.Position Then
                    Throw New ABNFException(
                        String.Format("解析エラー：'{0}'", tr.ToString(prevPos))
                    )
                End If
                prevPos = tr.Position
            Loop

            ' マッチ結果を返す
            Return New ExpressionRange(Me, tr, startPos, tr.Position, ranges.ToArray())
        End Function

    End Class

End Namespace
