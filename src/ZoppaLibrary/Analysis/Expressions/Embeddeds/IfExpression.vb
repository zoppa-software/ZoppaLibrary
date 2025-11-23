Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' If式を表す構造体です。
    ''' この構造体は、条件と内側の式を保持し、式の型を提供します。
    ''' </summary>
    ''' <remarks>
    ''' If式は、条件に基づいて異なる処理を実行するために使用されます。
    ''' </remarks>
    NotInheritable Class IfExpression
        Implements IExpression

        ' Ifの条件
        Private ReadOnly _condition As IExpression

        ' 内側の式
        Private ReadOnly _innerExpr As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="condition">Ifの条件式。</param>
        ''' <param name="innerExpr">内側の式。</param>
        ''' <remarks>
        ''' このコンストラクタは、If式の条件と内側の式を初期化します。
        ''' </remarks>
        Public Sub New(condition As IExpression, innerExpr As IExpression)
            If condition Is Nothing Then
                Throw New ArgumentNullException(NameOf(condition))
            End If
            If innerExpr Is Nothing Then
                Throw New ArgumentNullException(NameOf(innerExpr))
            End If
            Me._condition = condition
            Me._innerExpr = innerExpr
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.IfExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>条件が真の場合は内側の式の値、偽の場合はNull。</returns>
        ''' <remarks>
        ''' このメソッドは、Ifの条件を評価し、条件が真であれば内側の式の値を返します。
        ''' 偽の場合はNullを返します。
        ''' </remarks>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Dim isTrue = _condition.GetValue(venv).Bool
            If isTrue Then
                Return _innerExpr.GetValue(venv)
            Else
                Return Nothing ' 偽の場合はNullを返す
            End If
        End Function

    End Class

End Namespace
