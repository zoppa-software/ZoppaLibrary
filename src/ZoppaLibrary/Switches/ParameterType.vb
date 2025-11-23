Option Strict On
Option Explicit On

Namespace Switches

    ''' <summary>パラメータの型を表す列挙体。</summary>
    ''' <remarks>
    ''' 1:文字列、2:整数、4:倍精度浮動小数点数、8:URI、16:JSON、32:配列。
    ''' </remarks>
    Public Enum ParameterType As Integer

        ''' <summary>パラメータなし。</summary>
        None = 0

        ''' <summary>文字列パラメータ。</summary>
        Str = 1

        ''' <summary>整数パラメータ。</summary>
        Int = 2

        ''' <summary>倍精度浮動小数点数パラメータ。</summary>
        ''' <remarks>倍精度浮動小数点数は、64ビットの浮動小数点数を表します。</remarks>
        Dbl = 4

        ''' <summary>URIパラメータ。</summary>
        URI = 8

        ''' <summary>JSONパラメータ。</summary>
        JOSN = 16

        ''' <summary>配列パラメータ。</summary>
        ''' <remarks>配列は、任意の型の要素を持つことができます。</remarks>
        Array = 32

    End Enum

End Namespace