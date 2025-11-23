Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 値の型を定義する列挙型です。
    ''' 値は数値、文字列、真偽値、または配列のいずれかを表すことができます。
    ''' </summary>
    Public Enum ValueType

        ''' <summary>未定義値。</summary>
        Null

        ''' <summary>数値。</summary>
        Number

        ''' <summary>文字列。</summary>
        Str

        ''' <summary>真偽値。</summary>
        Bool

        ''' <summary>配列。</summary>
        Array

        ''' <summary>日付。</summary>
        DateTime

        ''' <summary>時間。</summary>
        TimeSpan

        ''' <summary>オブジェクト。</summary>
        Obj

    End Enum

End Namespace