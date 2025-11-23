Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' If式を表す構造体です。
    ''' この構造体は、複数のIf式を保持し、式の型を提供します。
    ''' </summary>
    ''' <remarks>
    ''' If式は、条件に基づいて異なる処理を実行するために使用されます。
    ''' </remarks>
    NotInheritable Class IfStatementExpression
        Implements IExpression

        ' If式のリスト
        Private ReadOnly _ifExprs As IExpression()

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="ifExprs">If式のリスト。</param>
        Public Sub New(ifExprs As IExpression())
            If ifExprs Is Nothing Then
                Throw New ArgumentNullException(NameOf(ifExprs))
            End If
            _ifExprs = ifExprs
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.IfStatementExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>リストの値。</returns>
        ''' <remarks>
        ''' このメソッドは、リスト内の各式を評価し、その結果を返します。
        ''' </remarks>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            For Each expr In _ifExprs
                Using venv.GetScope()
                    ' 各If式を評価
                    Dim value = expr.GetValue(venv)

                    ' 条件が真の場合はその値を返す
                    If value IsNot Nothing Then
                        Return value
                    End If
                End Using
            Next

            ' すべての条件が偽の場合は空の文字列を返す
            Return StringValue.Empty
        End Function

    End Class

End Namespace
