Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 埋込式のトリムステートメントを表すクラス。
    ''' 指定された文字列をトリムします。
    ''' </summary>
    ''' <remarks>
    ''' このクラスは、埋込式の一部として使用され、指定された文字列から特定の文字列をトリムします。
    ''' </remarks>
    NotInheritable Class TrimStatementExpression
        Implements IExpression

        ''' <summary>トリムする文字列の配列。</summary>
        ''' <remarks>
        ''' この配列には、トリム対象の文字列が含まれます。
        ''' </remarks>
        Private ReadOnly _trimStrs As IExpression()

        ''' <summary>トリム対象のコンテンツ式。</summary>
        ''' <remarks>
        ''' この式は、トリム対象の文字列を提供します。
        ''' </remarks>
        Private ReadOnly _contentsExpr As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="trimStrs">トリムする文字列の配列。</param>
        ''' <param name="contentsExpr">トリム対象のコンテンツ式。</param>
        ''' <exception cref="ArgumentNullException">引数がnullの場合にスローされます。</exception>
        Public Sub New(trimStrs() As IExpression, contentsExpr As IExpression)
            If trimStrs Is Nothing Then
                Throw New ArgumentNullException(NameOf(trimStrs))
            End If
            If contentsExpr Is Nothing Then
                Throw New ArgumentNullException(NameOf(contentsExpr))
            End If
            Me._trimStrs = trimStrs
            Me._contentsExpr = contentsExpr
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.TrimExpression
            End Get
        End Property

        ''' <summary>
        ''' 式の値を取得します。
        ''' 指定された文字列からトリム対象の文字列を削除し、結果を返します。
        ''' </summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>トリム後の文字列の値。</returns>
        ''' <exception cref="ArgumentNullException">引数がnullの場合にスローされます。</exception>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            ' トリムした文字列を取得します。
            Dim contentsValue = _contentsExpr.GetValue(venv).Str.Trim()

            ' 前後で一致する文字列をトリムします。
            Dim st As Integer = 0
            Dim ed As Integer = 0
            For Each trimStrExpr In _trimStrs
                Dim trimStrValue = trimStrExpr.GetValue(venv).Str

                ' 前要素の一致でトリム
                If contentsValue.StartWith(trimStrValue) AndAlso st < trimStrValue.ByteLength Then
                    st = trimStrValue.ByteLength
                End If

                ' 後要素の一致でトリム
                If contentsValue.EndWith(trimStrValue) AndAlso ed < trimStrValue.ByteLength Then
                    ed = trimStrValue.ByteLength
                End If
            Next

            ' トリムされた文字列を返します。
            If st > 0 OrElse ed > 0 Then
                ' スライスを使用してトリムされた文字列を取得します。
                Dim trimmedValue = U8String.NewSlice(contentsValue, st, contentsValue.ByteLength - st - ed).Trim()
                Return New StringValue(trimmedValue)
            End If
            Return New StringValue(contentsValue)
        End Function

    End Class

End Namespace
