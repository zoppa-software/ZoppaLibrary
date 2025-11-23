Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' default式を表す構造体です。
    ''' この構造体は、マッチ式と式を保持し、式の型を提供します。
    ''' </summary>
    ''' <remarks>
    ''' case式は、特定の条件に基づいて値を選択するために使用されます。
    ''' </remarks>
    NotInheritable Class SelectDefaultExpression
        Implements IExpression

        ' 式
        Private ReadOnly _expression As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="expression">式。</param>
        Public Sub New(expression As IExpression)
            If expression Is Nothing Then
                Throw New ArgumentNullException(NameOf(expression))
            End If
            Me._expression = expression
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.SelectDefaultExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>式の値。</returns>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return _expression.GetValue(venv)
        End Function

    End Class

End Namespace
