Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 埋込式の削除ステートメントを表すクラス。
    ''' 指定された文字列ならば削除します。
    ''' </summary>
    ''' <remarks>
    ''' このクラスは、埋込式の一部として使用され、指定された文字列ならば削除します。
    ''' </remarks>
    NotInheritable Class RemStatementExpression
        Implements IExpression

        ''' <summary>削除する文字列の配列。</summary>
        ''' <remarks>
        ''' この配列には、削除対象の文字列が含まれます。
        ''' </remarks>
        Private ReadOnly _remStrs As IExpression()

        ''' <summary>削除対象のコンテンツ式。</summary>
        ''' <remarks>
        ''' この式は、削除対象の文字列を提供します。
        ''' </remarks>
        Private ReadOnly _contentsExpr As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="remStrs">削除する文字列の配列。</param>
        ''' <param name="contentsExpr">削除対象のコンテンツ式。</param>
        ''' <exception cref="ArgumentNullException">引数がnullの場合にスローされます。</exception>
        Public Sub New(remStrs() As IExpression, contentsExpr As IExpression)
            If remStrs Is Nothing Then
                Throw New ArgumentNullException(NameOf(remStrs))
            End If
            If contentsExpr Is Nothing Then
                Throw New ArgumentNullException(NameOf(contentsExpr))
            End If
            Me._remStrs = remStrs
            Me._contentsExpr = contentsExpr
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.RemoveExpression
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
            Dim contentsValue = _contentsExpr.GetValue(venv).Str
            Dim trimed = contentsValue.Trim()

            ' 削除文字列と一致する場合は、削除します。
            For Each remStrExpr In _remStrs
                Dim remStrValue = remStrExpr.GetValue(venv).Str

                If trimed = remStrValue Then
                    Return StringValue.Empty
                End If
            Next

            ' 削除されていない場合は、元の文字列を返します。
            Return New StringValue(contentsValue)
        End Function

    End Class

End Namespace

