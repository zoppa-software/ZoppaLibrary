Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 変数代入式を表す構造体です。
    ''' 変数名と代入する値を表す式を持ちます。
    ''' </summary>
    ''' <remarks>
    ''' 変数代入式は、変数名と代入する値を定義するために使用されます。
    ''' 変数名はU8String型で指定され、値はIExpression型で表されます。
    ''' </remarks>
    NotInheritable Class SetVariableExpression
        Implements IExpression

        ' 変数名
        Private ReadOnly _name As U8String

        ' 変数値
        Private ReadOnly _value As IExpression

        ''' <summary>変数代入式を初期化します。</summary>
        ''' <param name="name">変数名。</param>
        ''' <param name="value">変数の値を表す式。</param>
        ''' <remarks>変数名は、U8String型で指定されます。</remarks>
        Public Sub New(name As U8String, value As IExpression)
            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value))
            End If
            _name = name
            _value = value
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.SetVariableExpression
            End Get
        End Property

        ''' <summary>変数定義式から変数を定義します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>空の文字列値。</returns>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            If venv.Contains(_name) Then
                venv.RegisterExpr(_name, _value)
            Else
                ' 変数が存在しない場合は、エラーをスローします。
                Throw New KeyNotFoundException($"'{_name}'変数が存在しません")
            End If
            Return StringValue.Empty
        End Function

    End Class

End Namespace