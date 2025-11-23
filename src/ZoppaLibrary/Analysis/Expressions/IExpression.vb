Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 式を表すインターフェイスです。
    ''' このインターフェイスは、式の型を定義し、式の解析や評価に使用されます。
    ''' </summary>
    Public Interface IExpression

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        ReadOnly Property Type As ExpressionType

        ''' <summary>
        ''' 式の値を取得します。
        ''' このメソッドは、式を評価し、その結果を返します。
        ''' </summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>式の値。</returns>
        ''' <remarks>
        ''' このメソッドは、式の評価を行い、結果としてValue型の値を返します。
        ''' </remarks>
        Function GetValue(venv As AnalysisEnvironment) As IValue

    End Interface

End Namespace
