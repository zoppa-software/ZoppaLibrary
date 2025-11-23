Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 日付時刻式を表す構造体です。
    ''' この構造体は、日付と時刻の値を保持し、式の型を提供します。
    ''' </summary>
    NotInheritable Class DateTimeExpression
        Implements IExpression

        ' 値
        Private ReadOnly _value As DateTime

        ''' <summary>日付時刻式のコンストラクタ。</summary>
        ''' <param name="value">日付時刻の値。</param>
        Public Sub New(value As DateTime)
            _value = value
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.DateTimeExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>日付時刻の値。</returns>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Return New DateTimeValue(_value)
        End Function

    End Class

End Namespace
