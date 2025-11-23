Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 配列フィールド式を表す構造体です。
    ''' この構造体は、配列の要素を保持し、式の型を提供します。
    ''' </summary>
    ''' <remarks>
    ''' 配列フィールド式は、複数の式を配列として扱うために使用されます。
    ''' 例: [x, y, z]
    ''' </remarks>
    NotInheritable Class ArrayFieldExpression
        Implements IExpression

        ' 配列
        Private ReadOnly _items() As IExpression

        ''' <summary>配列フィールド式のコンストラクタ。</summary>
        ''' <param name="items">配列の要素となる式の配列。</param>
        ''' <remarks>
        ''' 配列フィールド式は、複数の式を配列として扱うために使用されます。
        ''' </summary>
        Public Sub New(items() As IExpression)
            If items Is Nothing Then
                Throw New ArgumentNullException(NameOf(items))
            End If
            _items = items
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.ArrayFieldExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Dim _value(_items.Length - 1) As IValue
            For i As Integer = 0 To _items.Length - 1
                _value(i) = _items(i).GetValue(venv)
            Next
            Return New ArrayValue(_value)
        End Function

    End Class

End Namespace