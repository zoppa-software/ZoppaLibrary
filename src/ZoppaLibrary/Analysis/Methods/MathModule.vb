Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 解析用の数学関数を提供します。
    ''' </summary>
    Module MathModule

        ''' <summary>
        ''' 絶対値を返します。
        ''' </summary>
        ''' <param name="x">値。</param>
        ''' <returns>計算結果。</returns>
        Function Abs(x As IValue) As IValue
            Return Math.Abs(x.Number).ToNumberValue()
        End Function

        ''' <summary>
        ''' 指数関数を返します。
        ''' </summary>
        ''' <param name="x">値。</param>
        ''' <returns>計算結果。</returns>
        Function Pow(x As IValue, y As IValue) As IValue
            Return Math.Pow(x.Number, y.Number).ToNumberValue()
        End Function

        ''' <summary>
        ''' 四捨五入した値を返します。
        ''' </summary>
        ''' <param name="d">値。</param>
        ''' <param name="decimals">小数点以下の桁数。</param>
        ''' <returns>計算結果。</returns>
        Function Round(d As IValue, decimals As IValue) As IValue
            Return Math.Round(d.Number, CInt(decimals.Number)).ToNumberValue()
        End Function

    End Module

End Namespace
