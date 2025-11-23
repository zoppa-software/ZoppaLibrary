Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' Else式を表す構造体です。
    ''' この構造体は、内側の式を保持し、式の型を提供します。
    ''' </summary>
    ''' <remarks>
    ''' Else式は、条件が偽の場合に実行される処理を定義するために使用されます。
    ''' </remarks>
    NotInheritable Class ElseExpression
        Implements IExpression

        ' 内側の式
        Private ReadOnly _innerExpr As IExpression

        '''' <summary>コンストラクタ。</summary>
        ''' <param name="innerExpr">内側の式。</param>
        ''' <remarks>
        ''' このコンストラクタは、Else式の内側の式を初期化します。
        ''' </remarks>
        Public Sub New(innerExpr As IExpression)
            If innerExpr Is Nothing Then
                Throw New ArgumentNullException(NameOf(innerExpr))
            End If
            Me._innerExpr = innerExpr
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.ElseExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>内側の式の値。</returns>
        ''' <remarks>
        ''' このメソッドは、Else式の内側の式を評価し、その結果を返します。
        ''' </remarks>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return _innerExpr.GetValue(venv)
        End Function

    End Class

End Namespace
