Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 非エスケープ埋込テキスト式を表す構造体です。
    ''' この構造体は、埋込式を保持し、式の型を提供します。
    ''' </summary>
    NotInheritable Class NoEscapeUnfoldExpression
        Implements IExpression

        ' 埋込式
        Private ReadOnly _expr As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="expr">埋込式。</param>
        Public Sub New(expr As IExpression)
            If expr Is Nothing Then
                Throw New ArgumentNullException(NameOf(expr))
            End If
            _expr = expr
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.NoEscapeUnfoldExpression
            End Get
        End Property

        ''' <summary>
        ''' 式の値を取得します。
        ''' 埋込式は、埋込テキストを評価して返します。
        ''' </summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>展開された埋込テキストの値。</returns>
        ''' <exception cref="InvalidOperationException">不正な操作が行われた場合にスローされます。</exception>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return _expr.GetValue(venv)
        End Function

    End Class

End Namespace