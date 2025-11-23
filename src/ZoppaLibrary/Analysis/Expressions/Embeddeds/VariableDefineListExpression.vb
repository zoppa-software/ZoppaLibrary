Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 変数定義式リストを表す構造体です。
    ''' 変数定義式の配列を持ち、これらを一括で定義します。
    ''' </summary>
    ''' <remarks>
    ''' 変数定義式リストは、複数の変数定義式をまとめて扱うために使用されます。
    ''' 各変数定義式は、変数名とその値を表す式を持ちます。
    ''' </remarks>
    NotInheritable Class VariableDefineListExpression
        Implements IExpression

        ' 変数リスト
        Private ReadOnly _vardefines As VariableDefineExpression()

        ''' <summary>変数定義式リストを初期化します。</summary>
        ''' <param name="vardefines">変数定義式の配列。</param>
        ''' <remarks>各変数定義式は、変数名とその値を表す式を持ちます。</remarks>
        Public Sub New(vardefines As VariableDefineExpression())
            If vardefines Is Nothing Then
                Throw New ArgumentNullException(NameOf(vardefines))
            End If
            _vardefines = vardefines
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.VariableDefineExpression
            End Get
        End Property

        ''' <summary>変数定義式から変数を定義します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>空の文字列値。</returns>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            For Each vardef In _vardefines
                vardef.GetValue(venv)
            Next
            Return StringValue.Empty
        End Function

    End Class

End Namespace