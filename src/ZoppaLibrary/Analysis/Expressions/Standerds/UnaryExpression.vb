Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 単項演算子式を表す構造体です。
    ''' この構造体は、単項演算子の種類と対象となる式を保持し、式の評価を行います。
    ''' </summary>
    ''' <remarks>
    ''' 単項演算子は、1つの式に対して適用される演算子です。
    ''' 例: -x, +x, Not x
    ''' </remarks>
    NotInheritable Class UnaryExpression
        Implements IExpression

        ''' <summary>単項演算子の種類。</summary>
        Private ReadOnly _wordType As WordType

        ''' <summary>対象となる式。</summary>
        ''' <remarks>単項演算子は、1つの式に対して適用されます。</remarks>
        Private ReadOnly _expression As IExpression

        ''' <summary>単項演算子式のコンストラクタ。</summary>
        ''' <param name="wordType">単項演算子の種類。</param>
        ''' <param name="expression">対象となる式。</param>
        Public Sub New(wordType As WordType, expression As IExpression)
            If expression Is Nothing Then
                Throw New ArgumentNullException(NameOf(expression))
            End If
            _wordType = wordType
            _expression = expression
        End Sub

        ''' <summary>単項演算子の種類を取得します。</summary>
        ''' <returns>単項演算子の種類。</returns> 
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.UnaryExpression
            End Get
        End Property

        ''' <summary>
        ''' 式の値を取得します。
        ''' 単項演算子は、1つの式に対して適用されます。
        ''' </summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>演算結果の値。</returns>
        ''' <exception cref="InvalidOperationException">不正な操作が行われた場合にスローされます。</exception>
        ''' <exception cref="NotSupportedException">サポートされていない単項演算子が使用された場合にスローされます。</exception>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Dim fact = _expression.GetValue(venv)
            Select Case _wordType
                Case WordType.Not
                    If fact.Type = ValueType.Bool Then
                        Return If(fact.Bool, BooleanValue.FalseValue, BooleanValue.TrueValue)
                    Else
                        Throw New InvalidOperationException("単項演算子 Not は真偽値にのみ適用できます。")
                    End If
                Case WordType.Minus
                    If fact.Type = ValueType.Number Then
                        Return New NumberValue(-fact.Number)
                    Else
                        Throw New InvalidOperationException("単項演算子 - は数値にのみ適用できます。")
                    End If
                Case WordType.Plus
                    If fact.Type = ValueType.Number Then
                        Return New NumberValue(fact.Number)
                    Else
                        Throw New InvalidOperationException("単項演算子 + は数値にのみ適用できます。")
                    End If
                Case Else
                    Throw New NotSupportedException($"単項演算子は '{_wordType}' をサポートしていません。")
            End Select
        End Function

    End Class

End Namespace
