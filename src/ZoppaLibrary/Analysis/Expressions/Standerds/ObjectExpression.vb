Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' オブジェクト式を表す構造体です。
    ''' この式は、任意のオブジェクトを値として持つことができます。
    ''' </summary>
    ''' <remarks>
    ''' この式は、特定の型に依存せず、任意のオブジェクトを扱うことができます。
    ''' </remarks>
    NotInheritable Class ObjectExpression
        Implements IExpression

        ' 値
        Private ReadOnly _value As Object

        ''' <summary>真偽値式のコンストラクタ。</summary>
        ''' <param name="value">真偽値の値。</param>
        Public Sub New(value As Object)
            _value = value
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.ObjectExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return New ObjectValue(Me._value)
        End Function

    End Class

End Namespace
