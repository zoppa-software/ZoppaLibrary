Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' ()括弧式を表す構造体です。
    ''' この構造体は、式を括弧で囲むことで優先順位を明示的に指定します。
    ''' </summary>
    ''' <remarks>
    ''' ()括弧式は、式の評価順序を制御するために使用されます。
    ''' 例: (x + y) * z
    ''' </remarks>
    NotInheritable Class ParenExpression
        Implements IExpression

        ''' <summary>対象となる式。</summary>
        Private ReadOnly _expression As IExpression

        ''' <summary>()括弧式のコンストラクタ。</summary>
        ''' <param name="expression">対象となる式。</param>
        Public Sub New(expression As IExpression)
            _expression = expression
        End Sub

        ''' <summary>()括弧式の種類を取得します。</summary>
        ''' <returns>()括弧式の種類。</returns> 
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.ParenExpression
            End Get
        End Property

        ''' <summary>
        ''' 式の値を取得します。
        ''' 単項演算子は、1つの式に対して適用されます。
        ''' </summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>演算結果の値。</returns>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return _expression.GetValue(venv)
        End Function

    End Class

End Namespace
