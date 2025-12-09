Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 式のインスタンスを格納するモジュール。
    ''' </summary>
    Module ExpressionInstancesModule

        ''' <summary>
        ''' 1文字の任意の文字にマッチする式のインスタンス。
        ''' </summary>
        Private _characterExpr As New Lazy(Of CharacterExpression)(
            Function() New CharacterExpression()
        )

        ''' <summary>
        ''' 1文字の任意の文字にマッチする式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>1文字の任意の文字にマッチする式のインスタンス。</returns>
        Public Function CharacterExpr() As CharacterExpression
            Return _characterExpr.Value
        End Function

        ''' <summary>
        ''' 空白文字式を表す式のインスタンス。
        ''' </summary>
        Private _sExpr As New Lazy(Of SpaceExpression)(
            Function() New SpaceExpression()
        )

        ''' <summary>
        ''' 空白文字式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>空白文字式を表す式のインスタンス。</returns>
        Public Function SpaceExpr() As SpaceExpression
            Return _sExpr.Value
        End Function

        ''' <summary>
        ''' 左辺式を表す式のインスタンス。
        ''' </summary>
        Private _lhsExpr As New Lazy(Of LhsExpression)(
            Function() New LhsExpression()
        )

        ''' <summary>
        ''' 左辺式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>左辺式を表す式のインスタンス。</returns>
        Public Function LhsExpr() As LhsExpression
            Return _lhsExpr.Value
        End Function

        ''' <summary>
        ''' 右辺式を表す式のインスタンス。
        ''' </summary>
        Private _rhsExpr As New Lazy(Of RhsExpression)(
            Function() New RhsExpression()
        )

        ''' <summary>
        ''' 右辺式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>右辺式を表す式のインスタンス。</returns>
        Public Function RhsExpr() As RhsExpression
            Return _rhsExpr.Value
        End Function

        ''' <summary>
        ''' 終端記号式を表す式のインスタンス。
        ''' </summary>
        Private _terminalExpr As New Lazy(Of TerminalExpression)(
            Function() New TerminalExpression()
        )

        ''' <summary>
        ''' 終端記号式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>終端記号式を表す式のインスタンス。</returns>
        Public Function TerminalExpr() As TerminalExpression
            Return _terminalExpr.Value
        End Function

        ''' <summary>
        ''' 識別子式を表す式のインスタンス。
        ''' </summary>
        Private _identifierExpr As New Lazy(Of IdentifierExpression)(
            Function() New IdentifierExpression()
        )

        ''' <summary>
        ''' 識別子式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>識別子式を表す式のインスタンス。</returns>
        Public Function IdentifierExpr() As IdentifierExpression
            Return _identifierExpr.Value
        End Function

        ''' <summary>
        ''' 終端記号を表す式のインスタンス。
        ''' </summary>
        Private _terminatorExpr As New Lazy(Of TerminatorExpression)(
            Function() New TerminatorExpression()
        )

        ''' <summary>
        ''' 終端記号を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>終端記号を表す式のインスタンス。</returns>
        Public Function TerminatorExpr() As TerminatorExpression
            Return _terminatorExpr.Value
        End Function

        ''' <summary>
        ''' 縦棒区切りのカンマ区切りの式を表す式のインスタンス。
        ''' </summary>
        Private _alternatExpr As New Lazy(Of AlternationExpression)(
            Function() New AlternationExpression()
        )

        ''' <summary>
        ''' 縦棒区切りのカンマ区切りの式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>縦棒区切りのカンマ区切りの式を表す式のインスタンス。</returns>
        Public Function AlternatExpr() As AlternationExpression
            Return _alternatExpr.Value
        End Function

        ''' <summary>
        ''' 1文字の英字 (a-z, A-Z) にマッチする式のインスタンス。
        ''' </summary>
        Private _letterExpr As New Lazy(Of LetterExpression)(
            Function() New LetterExpression()
        )

        ''' <summary>
        ''' 1文字の英字 (a-z, A-Z) にマッチする式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>1文字の英字 (a-z, A-Z) にマッチする式のインスタンス。</returns>
        Public Function LetterExpr() As LetterExpression
            Return _letterExpr.Value
        End Function

        ''' <summary>
        ''' 1文字の数字 (0-9) にマッチする式のインスタンス。
        ''' </summary>
        Private _digitExpr As New Lazy(Of DigitExpression)(
            Function() New DigitExpression()
        )

        ''' <summary>
        ''' 1文字の数字 (0-9) にマッチする式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>1文字の数字 (0-9) にマッチする式のインスタンス。</returns>
        Public Function DigitExpr() As DigitExpression
            Return _digitExpr.Value
        End Function

        ''' <summary>
        ''' ルール式を表す式のインスタンス。
        ''' </summary>
        Private _ruleExpr As New Lazy(Of RuleExpression)(
            Function() New RuleExpression()
        )

        ''' <summary>
        ''' ルール式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>ルール式を表す式のインスタンス。</returns>
        Public Function RuleExpr() As RuleExpression
            Return _ruleExpr.Value
        End Function

        ''' <summary>
        ''' 終端記号、識別子、または括弧で囲まれた式を表す式のインスタンス。
        ''' </summary>
        Private _termExpr As New Lazy(Of TermExpression)(
            Function() New TermExpression()
        )

        ''' <summary>
        ''' 終端記号、識別子、または括弧で囲まれた式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>終端記号、識別子、または括弧で囲まれた式を表す式のインスタンス。</returns>
        Public Function TermExpr() As TermExpression
            Return _termExpr.Value
        End Function

        ''' <summary>
        ''' 繰り返し記号付きの式を表す式のインスタンス。
        ''' </summary>
        Private _factorExpr As New Lazy(Of FactorExpression)(
            Function() New FactorExpression()
        )

        ''' <summary>
        ''' 繰り返し記号付きの式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>繰り返し記号付きの式を表す式のインスタンス。</returns>
        Public Function FactorExpr() As FactorExpression
            Return _factorExpr.Value
        End Function

        ''' <summary>
        ''' 1文字の記号にマッチする式のインスタンス。
        ''' </summary>
        Private _symbolExpr As New Lazy(Of SymbolExpression)(
            Function() New SymbolExpression()
        )

        ''' <summary>
        ''' 1文字の記号にマッチする式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>1文字の記号にマッチする式のインスタンス。</returns>
        Public Function SymbolExpr() As SymbolExpression
            Return _symbolExpr.Value
        End Function

        ''' <summary>
        ''' カンマ区切りの式を表す式のインスタンス。
        ''' </summary>
        Private _concatExpr As New Lazy(Of ConcatenationExpression)(
            Function() New ConcatenationExpression()
        )

        ''' <summary>
        ''' カンマ区切りの式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>カンマ区切りの式を表す式のインスタンス。</returns>
        Public Function ConcatenationExpr() As ConcatenationExpression
            Return _concatExpr.Value
        End Function

        ''' <summary>
        ''' コメント式を表す式のインスタンス。
        ''' </summary>
        Private _commentExpr As New Lazy(Of CommentExpression)(
            Function() New CommentExpression()
        )

        ''' <summary>
        ''' コメント式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>コメント式を表す式のインスタンス。</returns>
        Public Function CommentExpr() As CommentExpression
            Return _commentExpr.Value
        End Function

        ''' <summary>
        ''' 特殊文式を表す式のインスタンス。
        ''' </summary>
        Private _specialSeqExpr As New Lazy(Of SpecialSeqExpression)(
            Function() New SpecialSeqExpression()
        )

        ''' <summary>
        ''' 特殊文式を表す式のインスタンスを取得します。
        ''' </summary>
        ''' <returns>特殊文式を表す式のインスタンス。</returns>
        Public Function SpecialSeqExpr() As SpecialSeqExpression
            Return _specialSeqExpr.Value
        End Function

    End Module

End Namespace
