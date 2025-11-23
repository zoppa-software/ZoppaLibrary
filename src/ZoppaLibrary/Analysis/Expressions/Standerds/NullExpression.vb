Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' Null式を表す構造体です。
    ''' この構造体は、Null値を保持し、式の型を提供します。
    ''' Null式は、値が存在しないことを示すために使用されます。
    ''' </summary>
    ''' <remarks>
    ''' Null式は、値が存在しないことを明示的に示すために使用されます。
    ''' </remarks>
    NotInheritable Class NullExpression
        Implements IExpression

        ''' <summary>NullExpressionのインスタンスを取得します。</summary>
        ''' <returns>NullExpressionのインスタンス。</returns>
        ''' <remarks>
        ''' このプロパティは、NullExpressionの唯一のインスタンスを返します。
        ''' </remarks>
        Private Shared ReadOnly _instance As New Lazy(Of NullExpression)(Function() New NullExpression())

        ''' <summary>NullExpressionのインスタンスを取得します。</summary>
        ''' <returns>NullExpressionのインスタンス。</returns>
        ''' <remarks>
        ''' このプロパティは、NullExpressionの唯一のインスタンスを返します。
        ''' </remarks>
        Public Shared ReadOnly Property Value As NullExpression
            Get
                Return _instance.Value
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <remarks>
        ''' このコンストラクタは、NullExpressionのインスタンスを初期化します。
        ''' </remarks>

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.NullExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return NullValue.Value
        End Function

    End Class

End Namespace
