Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' リスト式を表す構造体です。
    ''' この構造体は、複数の式をリストとして保持し、式の型を提供します。
    ''' </summary>
    ''' <remarks>
    ''' リスト式は、複数の式をまとめて扱うために使用されます。
    ''' 各式は個別に評価され、その結果が式として返されます。
    ''' </remarks>
    NotInheritable Class ListExpression
        Implements IExpression

        ' 各式のリスト
        Private ReadOnly _expressions As IExpression()

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="expressions">リスト内の式のリスト。</param>
        Public Sub New(expressions As IExpression())
            If expressions Is Nothing Then
                Throw New ArgumentNullException(NameOf(expressions))
            End If
            _expressions = expressions
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.ListExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>リストの値。</returns>
        ''' <remarks>
        ''' このメソッドは、リスト内の各式を評価し、その結果を返します。
        ''' </remarks>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Dim result As New List(Of Byte)()
            For Each expr In _expressions
                Dim getstr = If(expr.GetValue(venv)?.Str, U8String.Empty)
                If getstr.Length > 0 Then
                    result.AddRange(getstr.GetByteEnumerator())
                End If
            Next

            ' 文字列に変換して返す
            Return New StringValue(U8String.NewStringChangeOwner(result.ToArray()))
        End Function

    End Class

End Namespace
