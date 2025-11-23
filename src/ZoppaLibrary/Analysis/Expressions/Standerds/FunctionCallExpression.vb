Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 関数呼び出し式を表す構造体です。
    ''' この構造体は、関数名と引数リストを保持し、式の型を提供します。
    ''' </summary>
    NotInheritable Class FunctionCallExpression
        Implements IExpression

        ' 関数名
        Private ReadOnly _name As U8String

        ' 引数リスト
        Private ReadOnly _parameter As IExpression()

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="name">関数名。</param>
        ''' <param name="parameter">関数の引数リスト。</param>
        ''' <remarks>
        ''' このコンストラクタは、関数呼び出し式を初期化します。
        ''' </remarks>
        Public Sub New(name As U8String, parameter() As IExpression)
            Me._name = name
            Me._parameter = If(parameter, New IExpression() {})
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.FunctionCallExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>関数呼び出しの結果としての値。</returns>
        ''' <remarks>
        ''' このメソッドは、関数名と引数リストを使用して関数を呼び出し、その結果を返します。
        ''' </remarks>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Dim prms = _parameter.Select(Function(p) p.GetValue(venv))
            Return venv.CallFunction(_name, prms.ToArray())
        End Function

    End Class

End Namespace
