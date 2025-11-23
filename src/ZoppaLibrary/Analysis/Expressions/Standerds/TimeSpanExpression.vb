Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' TimeSpan式を表す構造体です。
    ''' この構造体は、TimeSpanの値を保持し、式の型を提供します。
    ''' </summary>
    NotInheritable Class TimeSpanExpression
        Implements IExpression

        ' 値
        Private ReadOnly _value As TimeSpan

        ''' <summary>TimeSpan式のコンストラクタ。</summary>
        ''' <param name="value">TimeSpanの値。</param>
        Public Sub New(value As TimeSpan)
            _value = value
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.TimeSpanExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>TimeSpanの値。</returns>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return New TimeSpanValue(_value)
        End Function

    End Class

End Namespace
