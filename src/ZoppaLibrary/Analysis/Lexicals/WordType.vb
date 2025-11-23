Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 単語の種類を定義する列挙型です。
    ''' この列挙型は、プログラム内で使用されるキーワードや記号を表します。
    ''' </summary>
    Public Enum WordType

        ''' <summary>論理積。</summary>
        AndOperator

        ''' <summary>代入。</summary>
        Assign

        ''' <summary>バックスラッシュ。</summary>
        Backslash

        ''' <summary>コロン。</summary>
        Colon

        ''' <summary>カンマ。</summary>
        Comma

        ''' <summary>日時リテラル。</summary>
        DateTimeLiteral

        ''' <summary>スラッシュ。</summary>
        Divide

        ''' <summary>ダラー。</summary>
        Dollar

        ''' <summary>等しい。</summary>
        Equal

        ''' <summary>偽リテラル。</summary>
        FalseLiteral

        ''' <summary>大なりイコール。</summary>
        GreaterEqual

        ''' <summary>大なり。</summary>
        GreaterThan

        ''' <summary>ハッシュ。</summary>
        Hash

        ''' <summary>識別子。</summary>
        ''' <remarks>
        ''' 変数名や関数名など、識別子として使用される名前を表します。
        ''' </remarks>
        Identifier

        ''' <summary>インキーワード。</summary>
        InKeyword

        ''' <summary>左ブラケット。</summary>
        LeftBracket

        ''' <summary>左括弧。</summary>
        LeftParen

        ''' <summary>小なりイコール。</summary>
        LessEqual

        ''' <summary>小なり。</summary>
        LessThan

        ''' <summary>マイナス。</summary>
        Minus

        ''' <summary>アスタリスク。</summary>
        Multiply

        ''' <summary>否定。</summary>
        [Not]

        ''' <summary>等しくない。</summary>
        NotEqual

        ''' <summary>ヌルリテラル。</summary>
        NullLiteral

        ''' <summary>数値。</summary>
        ''' <remarks>
        ''' 数値リテラルを表します。整数や浮動小数点数など、数値として解釈される値を含みます。
        ''' </remarks>
        Number

        ''' <summary>論理和。</summary>
        OrOperator

        ''' <summary>ピリオド。</summary>
        Period

        ''' <summary>プラス。</summary>
        ''' <remarks>
        ''' 数値の加算や文字列の連結など、加算操作を表します。
        ''' </remarks>
        Plus

        ''' <summary>クエスチョン。</summary>
        Question

        ''' <summary>右ブラケット。</summary>
        RightBracket

        ''' <summary>右括弧。</summary>
        RightParen

        ''' <summary>セミコロン。</summary>
        Semicolon

        ''' <summary>
        ''' 文字列リテラル。
        ''' 文字列を表すリテラルで、通常は引用符で囲まれたテキストを含みます。
        ''' </summary>
        StringLiteral

        ''' <summary>時間リテラル。</summary>
        TimeSpanLiteral

        ''' <summary>真リテラル。</summary>
        TrueLiteral

        ''' <summary>排他的論理和。</summary>
        XorOperator

    End Enum

End Namespace
