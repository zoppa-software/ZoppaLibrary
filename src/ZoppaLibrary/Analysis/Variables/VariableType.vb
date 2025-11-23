Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 変数の型を定義する列挙型です。
    ''' 変数は数値、文字列、または真偽値のいずれかを表すことができます。
    ''' </summary>
    Public Enum VariableType

        ''' <summary>式。</summary>
        Expr

        ''' <summary>数値。</summary>
        Number

        ''' <summary>文字列。</summary>
        Str

        ''' <summary>真偽値。</summary>
        Bool

        ''' <summary>配列。</summary>
        Array

        ''' <summary>日時。</summary>
        [Date]

        ''' <summary>時間。</summary>
        [Time]

        ''' <summary>オブジェクト。</summary>
        Obj

    End Enum

End Namespace