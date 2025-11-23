Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' select式を表す構造体です。
    ''' この構造体は、複数のcase式を保持し、式の型を提供します。
    ''' </summary>
    ''' <remarks>
    ''' select式は、条件に基づいて異なる処理を実行するために使用されます。
    ''' </remarks>
    NotInheritable Class SelectStatementExpression
        Implements IExpression

        ' select式
        Private ReadOnly _selectExprs As IExpression

        ' case式のリスト
        Private ReadOnly _caseExprs As IExpression()

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="caseExprs">If式のリスト。</param>
        Public Sub New(selectExprs As IExpression, caseExprs As IExpression())
            If selectExprs Is Nothing Then
                Throw New ArgumentNullException(NameOf(selectExprs))
            End If
            If caseExprs Is Nothing Then
                Throw New ArgumentNullException(NameOf(caseExprs))
            End If
            _selectExprs = selectExprs
            _caseExprs = caseExprs
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.SelectStatementExpression
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

            ' select式の値を取得
            Dim selectExpr = DirectCast(_selectExprs, SelectExpression)
            Dim matchValue = selectExpr.MatchExpr.GetValue(venv)
            Dim ss = If(_selectExprs.GetValue(venv)?.Str, U8String.Empty)
            If ss.Length > 0 Then
                result.AddRange(ss.GetByteEnumerator())
            End If

            ' case式のリストをループして、条件に一致するものを探す
            For Each expr In _caseExprs
                Select Case expr.Type
                    Case ExpressionType.SelectCaseExpression
                        ' 値が一致する場合はその値を返す
                        Dim caseExpr = DirectCast(expr, SelectCaseExpression)
                        If BinaryExpression.CompareValues(matchValue, caseExpr.MatchExpr.GetValue(venv)) Then
                            Using venv.GetScope()
                                Return caseExpr.GetValue(venv)
                            End Using
                        End If

                    Case ExpressionType.SelectDefaultExpression
                        ' デフォルトの値を取得
                        Using venv.GetScope()
                            Return expr.GetValue(venv)
                        End Using

                    Case Else
                        ' その他の式は無視する
                End Select
            Next

            ' すべての条件が偽の場合は空の文字列を返す
            Return StringValue.Empty
        End Function

    End Class

End Namespace
