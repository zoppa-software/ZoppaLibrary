Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 仮想改行式を表す構造体です。
    ''' この構造体は、改行を表現します。
    ''' </summary>
    NotInheritable Class VlBrExpression
        Implements IExpression

        ' インスタンス
        Private Shared ReadOnly _instance As New Lazy(Of VlBrExpression)(Function() New VlBrExpression())

        ''' <summary>
        ''' 改行式のインスタンスを取得します。
        ''' このインスタンスは、常に同じ空の式を返します。
        ''' </summary>
        Public Shared ReadOnly Property Instance As VlBrExpression
            Get
                Return _instance.Value
            End Get
        End Property

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.VlBrExpression
            End Get
        End Property

        ''' <summary>
        ''' 式の値を取得します。
        ''' 空の式は何も値を持たないため、空の文字列値を返します。
        ''' </summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>空の文字列値。</returns>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return StringValue.Empty
        End Function

    End Class

End Namespace

