Option Explicit On
Option Strict On

Namespace ABNF

    ''' <summary>
    ''' 式のインスタンスを格納するモジュール。
    ''' </summary>
    Module ExpressionInstancesModule

        ''' <summary>
        ''' 選択式のインスタンス。
        ''' </summary>
        Private _alterExpr As New Lazy(Of AlternationExpression)(
            Function() New AlternationExpression()
        )

        ''' <summary>
        ''' 選択式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>選択式のインスタンス。</returns>
        Public Function ABNFAlterExpr() As AlternationExpression
            Return _alterExpr.Value
        End Function

        ''' <summary>
        ''' 文字列式のインスタンス。
        ''' </summary>
        Private _charValExpr As New Lazy(Of CharValExpression)(
            Function() New CharValExpression()
        )

        ''' <summary>
        ''' 文字列式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>文字列式のインスタンス。</returns>
        Public Function ABNFCharValExpr() As CharValExpression
            Return _charValExpr.Value
        End Function

        ''' <summary>
        ''' コメント式のインスタンス。
        ''' </summary>
        Private _commentExpr As New Lazy(Of CommentExpression)(
            Function() New CommentExpression()
        )

        ''' <summary>
        ''' コメント式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>コメント式のインスタンス。</returns>
        Public Function ABNFCommentExpr() As CommentExpression
            Return _commentExpr.Value
        End Function

        ''' <summary>
        ''' コメントと改行の式のインスタンス。
        ''' </summary>
        Private _commentNlExpr As New Lazy(Of CommentNewLineExpression)(
            Function() New CommentNewLineExpression()
        )

        ''' <summary>
        ''' コメントと改行の式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>コメントと改行の式のインスタンス。</returns>
        Public Function ABNFCommentNlExpr() As CommentNewLineExpression
            Return _commentNlExpr.Value
        End Function

        ''' <summary>
        ''' コメントと空白の式のインスタンス。
        ''' </summary>
        Private _commentWspExpr As New Lazy(Of CommentWhiteSpaceExpression)(
            Function() New CommentWhiteSpaceExpression()
        )

        ''' <summary>
        ''' コメントと空白の式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>コメントと空白の式のインスタンス。</returns>
        Public Function ABNFCommentWspExpr() As CommentWhiteSpaceExpression
            Return _commentWspExpr.Value
        End Function

        ''' <summary>
        ''' 連結式のインスタンス。
        ''' </summary>
        Private _concatExpr As New Lazy(Of ConcatenationExpression)(
            Function() New ConcatenationExpression()
        )

        ''' <summary>
        ''' 連結式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>連結式を表す式のインスタンス。</returns>
        Public Function ABNFConcatExpr() As ConcatenationExpression
            Return _concatExpr.Value
        End Function

        ''' <summary>
        ''' 改行式のインスタンス。
        ''' </summary>
        Private _crlfExpr As New Lazy(Of CrLfExpression)(
            Function() New CrLfExpression()
        )

        ''' <summary>
        ''' 改行式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>改行式を表す式のインスタンス。</returns>
        Public Function ABNFCrLfExpr() As CrLfExpression
            Return _crlfExpr.Value
        End Function

        ''' <summary>
        ''' 要素式のインスタンス。
        ''' </summary>
        Private _elementsExpr As New Lazy(Of ElementsExpression)(
            Function() New ElementsExpression()
        )

        ''' <summary>
        ''' 要素式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>要素式のインスタンス。</returns>
        Public Function ABNFElementsExpr() As ElementsExpression
            Return _elementsExpr.Value
        End Function

        ''' <summary>
        ''' グループ式のインスタンス。
        ''' </summary>
        Private _groupExpr As New Lazy(Of GroupExpression)(
            Function() New GroupExpression()
        )

        ''' <summary>
        ''' グループ式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>グループ式を表す式のインスタンス。</returns>
        Public Function ABNFGroupExpr() As GroupExpression
            Return _groupExpr.Value
        End Function

        ''' <summary>
        ''' 数値式のインスタンス。
        ''' </summary>
        Private _numValExpr As New Lazy(Of NumValExpression)(
            Function() New NumValExpression()
        )

        ''' <summary>
        ''' 数値式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>数値式を表す式のインスタンス。</returns>
        Public Function ABNFNumValExpr() As NumValExpression
            Return _numValExpr.Value
        End Function

        ''' <summary>
        ''' 数値範囲式のインスタンス。
        ''' </summary>
        Private _numValRangeExpr As New Lazy(Of NumValExpression.Range)(
            Function() New NumValExpression.Range()
        )

        ''' <summary>
        ''' 数値範囲式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>数値範囲式を表す式のインスタンス。</returns>
        Public Function ABNFNumValRangeExpr() As NumValExpression.Range
            Return _numValRangeExpr.Value
        End Function

        ''' <summary>
        ''' 数値連結式のインスタンス。
        ''' </summary>
        Private _numValConcatExpr As New Lazy(Of NumValExpression.Concat)(
            Function() New NumValExpression.Concat()
        )

        ''' <summary>
        ''' 数値連結式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>数値連結式を表す式のインスタンス。</returns>
        Public Function ABNFNumValConcatExpr() As NumValExpression.Concat
            Return _numValConcatExpr.Value
        End Function

        ''' <summary>
        ''' オプション式のインスタンス。
        ''' </summary>
        Private _optionExpr As New Lazy(Of OptionExpression)(
            Function() New OptionExpression()
        )

        ''' <summary>
        ''' オプション式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>オプション式を表す式のインスタンス。</returns>
        Public Function ABNFOptionExpr() As OptionExpression
            Return _optionExpr.Value
        End Function

        ''' <summary>
        ''' 散文式のインスタンス。
        ''' </summary>
        Private _proseValExpr As New Lazy(Of ProseValExpression)(
            Function() New ProseValExpression()
        )

        ''' <summary>
        ''' 散文式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>散文式を表す式のインスタンス。</returns>
        Public Function ABNFProseValExpr() As ProseValExpression
            Return _proseValExpr.Value
        End Function

        ''' <summary>
        ''' 反復式のインスタンス。
        ''' </summary>
        Private _repeatExpr As New Lazy(Of RepetitionExpression)(
            Function() New RepetitionExpression()
        )

        ''' <summary>
        ''' 反復式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>反復式を表す式を表す式のインスタンス。</returns>
        Public Function ABNFRepeatExpr() As RepetitionExpression
            Return _repeatExpr.Value
        End Function

        ''' <summary>
        ''' ルールを表す式のインスタンス。
        ''' </summary>
        Private _ruleExpr As New Lazy(Of RuleExpression)(
            Function() New RuleExpression()
        )

        ''' <summary>
        ''' ルールを表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>ルールを表す式を表す式のインスタンス。</returns>
        Public Function ABNFRuleExpr() As RuleExpression
            Return _ruleExpr.Value
        End Function

        ''' <summary>
        ''' ルール名を表す式のインスタンス。
        ''' </summary>
        Private _ruleNameExpr As New Lazy(Of RuleNameExpression)(
            Function() New RuleNameExpression()
        )

        ''' <summary>
        ''' ルール名を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>ルール名を表す式を表す式のインスタンス。</returns>
        Public Function ABNFRuleNameExpr() As RuleNameExpression
            Return _ruleNameExpr.Value
        End Function

        ''' <summary>
        ''' 空白式のインスタンス。
        ''' </summary>
        Private _spaceExpr As New Lazy(Of SpaceExpression)(
            Function() New SpaceExpression()
        )

        ''' <summary>
        ''' 空白式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>空白式を表す式のインスタンス。</returns>
        Public Function ABNFSpaceExpr() As SpaceExpression
            Return _spaceExpr.Value
        End Function

    End Module

End Namespace
