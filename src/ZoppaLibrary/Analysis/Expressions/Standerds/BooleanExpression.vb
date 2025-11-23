Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 真偽値式を表す構造体です。
    ''' この構造体は、真偽値の値を保持し、式の型を提供します。
    ''' </summary>
    NotInheritable Class BooleanExpression
        Implements IExpression

        ' 真値
        Private Shared ReadOnly _trueInstance As New Lazy(Of BooleanExpression)(Function() New BooleanExpression(True))

        ' 偽値
        Private Shared ReadOnly _falseInstance As New Lazy(Of BooleanExpression)(Function() New BooleanExpression(False))

        ''' <summary>真の値を取得します。</summary>
        ''' <returns>真の値。</returns>
        Public Shared ReadOnly Property TrueValue As BooleanExpression
            Get
                Return _trueInstance.Value
            End Get
        End Property

        ''' <summary>偽の値を取得します。</summary>
        ''' <returns>偽の値。</returns>
        Public Shared ReadOnly Property FalseValue As BooleanExpression
            Get
                Return _falseInstance.Value
            End Get
        End Property

        ' 値
        Private ReadOnly _value As Boolean

        ''' <summary>真偽値式のコンストラクタ。</summary>
        ''' <param name="value">真偽値の値。</param>
        Private Sub New(value As Boolean)
            _value = value
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.BooleanExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return If(_value, BooleanValue.TrueValue, BooleanValue.FalseValue)
        End Function

    End Class

End Namespace
