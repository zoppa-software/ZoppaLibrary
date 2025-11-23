Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 文字列式を表す構造体です。
    ''' この構造体は、文字列の値を保持し、式の型を提供します。
    ''' </summary>
    NotInheritable Class StringExpression
        Implements IExpression

        ' 値
        Private ReadOnly _value As U8String

        ''' <summary>文字列式のコンストラクタ。</summary>
        ''' <param name="value">文字列の値。</param>
        Public Sub New(value As U8String)
            _value = value
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.StringExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return New StringValue(_value)
        End Function

    End Class

End Namespace
