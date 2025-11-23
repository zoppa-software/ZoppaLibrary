Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 二項演算子式を表す構造体です。
    ''' この構造体は、二項演算子の種類と左辺および右辺の式を保持し、式の評価を行います。
    ''' </summary>
    ''' <remarks>
    ''' 二項演算子は、2つの式に対して適用される演算子です。
    ''' 例: x + y, x - y, x * y, x / y
    ''' </remarks>
    NotInheritable Class BinaryExpression
        Implements IExpression

        ''' <summary>許容誤差。</summary>
        Private Const Epsilon As Double = 0.0000000001

        ''' <summary>左辺の式。</summary>
        Private ReadOnly _left As IExpression

        ''' <summary>右辺の式。</summary>
        Private ReadOnly _right As IExpression

        ''' <summary>演算子の種類。</summary>
        Private ReadOnly _wordType As WordType

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="wordType">演算子の種類。</param>
        ''' <param name="left">左辺の式。</param>
        ''' <param name="right">右辺の式。</param>
        Public Sub New(wordType As WordType, left As IExpression, right As IExpression)
            If left Is Nothing Then
                Throw New ArgumentNullException(NameOf(left))
            End If
            If right Is Nothing Then
                Throw New ArgumentNullException(NameOf(right))
            End If
            _left = left
            _right = right
            _wordType = wordType
        End Sub

        ''' <summary>演算子の種類を取得します。</summary>
        ''' <returns>演算子の種類。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.BinaryExpression
            End Get
        End Property

        ''' <summary>
        ''' 式の値を取得します。
        ''' 二項演算子は、左辺と右辺の式に対して適用されます。
        ''' </summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>演算結果の値。</returns>
        ''' <exception cref="InvalidOperationException">不正な操作が行われた場合にスローされます。</exception>
        ''' <exception cref="NotSupportedException">サポートされていない演算子が使用された場合にスローされます。</exception>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            Dim lv = _left.GetValue(venv)
            Dim rv = _right.GetValue(venv)
            Select Case _wordType
                Case WordType.Plus
                    If lv.Type = ValueType.Number AndAlso rv.Type = ValueType.Number Then
                        Return New NumberValue(lv.Number + rv.Number)
                    ElseIf lv.Type = ValueType.Str OrElse rv.Type = ValueType.Str Then
                        Return New StringValue(lv.Str.Concat(rv.Str))
                    ElseIf lv.Type = ValueType.DateTime AndAlso rv.Type = ValueType.TimeSpan Then
                        Return New DateTimeValue(lv.ToDate.Add(rv.ToTimeSpan))
                    ElseIf lv.Type = ValueType.TimeSpan AndAlso rv.Type = ValueType.DateTime Then
                        Return New DateTimeValue(rv.ToDate.Add(lv.ToTimeSpan))
                    ElseIf lv.Type = ValueType.TimeSpan AndAlso rv.Type = ValueType.TimeSpan Then
                        Return New TimeSpanValue(lv.ToTimeSpan.Add(rv.ToTimeSpan))
                    Else
                        Throw New InvalidOperationException("加算は数値または文字列に対してのみ適用できます。")
                    End If
                Case WordType.Minus
                    If lv.Type = ValueType.Number AndAlso rv.Type = ValueType.Number Then
                        Return New NumberValue(lv.Number - rv.Number)
                    ElseIf lv.Type = ValueType.DateTime AndAlso rv.Type = ValueType.TimeSpan Then
                        Return New DateTimeValue(lv.ToDate.Subtract(rv.ToTimeSpan))
                    ElseIf lv.Type = ValueType.TimeSpan AndAlso rv.Type = ValueType.TimeSpan Then
                        Return New TimeSpanValue(lv.ToTimeSpan.Subtract(rv.ToTimeSpan))
                    Else
                        Throw New InvalidOperationException("減算は数値に対してのみ適用できます。")
                    End If
                Case WordType.Multiply
                    If lv.Type = ValueType.Number AndAlso rv.Type = ValueType.Number Then
                        Return New NumberValue(lv.Number * rv.Number)
                    Else
                        Throw New InvalidOperationException("乗算は数値に対してのみ適用できます。")
                    End If
                Case WordType.Divide
                    If lv.Type = ValueType.Number AndAlso rv.Type = ValueType.Number Then
                        If rv.Number = 0 Then
                            Throw New DivideByZeroException("ゼロ除算は許可されていません。")
                        End If
                        Return New NumberValue(lv.Number / rv.Number)
                    Else
                        Throw New InvalidOperationException("除算は数値に対してのみ適用できます。")
                    End If
                Case WordType.Equal
                    Return If(CompareValues(lv, rv), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case WordType.NotEqual
                    Return If(CompareValues(lv, rv), BooleanValue.FalseValue, BooleanValue.TrueValue)
                Case WordType.GreaterThan
                    Return If(GreaterThan(lv, rv), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case WordType.GreaterEqual
                    Return If(GreaterEqual(lv, rv), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case WordType.LessThan
                    Return If(LessThan(lv, rv), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case WordType.LessEqual
                    Return If(LessEqual(lv, rv), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case WordType.AndOperator
                    Return If(AndOperator(lv, rv), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case WordType.OrOperator
                    Return If(OrOperator(lv, rv), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case WordType.XorOperator
                    Return If(XorOperator(lv, rv), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case Else
                    Throw New NotSupportedException($"サポートされていない演算子: {_wordType}")
            End Select
        End Function

        ''' <summary>左辺と右辺の値が等しいか比較します。</summary>
        ''' <param name="left">左辺の値。</param>
        ''' <param name="right">右辺の値。</param>
        ''' <returns>比較結果。等しい場合はTrue、それ以外はFalse。</returns>
        ''' <exception cref="NotSupportedException">サポートされていない値の型が使用された場合にスローされます。</exception>
        Public Shared Function CompareValues(left As IValue, right As IValue) As Boolean
            Select Case left.Type
                Case ValueType.Null
                    Return (right.Type = ValueType.Null OrElse right.Obj Is Nothing)
                Case ValueType.Number
                    Return Math.Abs(left.Number - right.Number) < Epsilon
                Case ValueType.Str
                    Return left.Str.Equals(right.Str)
                Case ValueType.Bool
                    Return left.Bool = right.Bool
                Case ValueType.DateTime
                    Return left.ToDate = right.ToDate
                Case ValueType.TimeSpan
                    Return left.ToTimeSpan = right.ToTimeSpan
                Case ValueType.Array
                    If left.Array.Length = right.Array.Length Then
                        For i As Integer = 0 To left.Array.Length - 1
                            If Not CompareValues(left.Array(i), right.Array(i)) Then
                                Return False
                            End If
                        Next
                        Return True
                    Else
                        Return False
                    End If
                Case ValueType.Obj
                    Return left.Obj.Equals(right.Obj)
                Case Else
                    Throw New NotSupportedException($"サポートされていない値の型: {left.Type}")
            End Select
        End Function

        ''' <summary>
        ''' 左辺の値が右辺の値より大きいかどうかを比較します。
        ''' このメソッドは、数値、文字列、日付、時間に対してのみ適用できます。
        ''' </summary>
        ''' <param name="left">左辺の値。</param>
        ''' <param name="right">右辺の値。</param>
        ''' <returns>比較結果。より大きければ真。</returns>
        Private Shared Function GreaterThan(left As IValue, right As IValue) As Boolean
            If left.Type = ValueType.Number Then
                Return left.Number > right.Number AndAlso Math.Abs(left.Number - right.Number) >= Epsilon
            ElseIf left.Type = ValueType.Str Then
                Return left.Str.CompareTo(right.Str) > 0
            ElseIf left.Type = ValueType.DateTime Then
                Return left.ToDate.CompareTo(right.ToDate) > 0
            ElseIf left.Type = ValueType.TimeSpan Then
                Return left.ToTimeSpan.CompareTo(right.ToTimeSpan) > 0
            Else
                Throw New InvalidOperationException("大なり演算子は数値、文字列、日付、時間に対してのみ適用できます。")
            End If
        End Function

        ''' <summary>
        ''' 左辺の値が右辺の値以上かどうかを比較します。
        ''' このメソッドは、数値、文字列、日付、時間に対してのみ適用できます。
        ''' </summary>
        ''' <param name="left">左辺の値。</param>
        ''' <param name="right">右辺の値。</param>
        ''' <returns>比較結果。より大きいか等しい場合は真。</returns>
        Private Shared Function GreaterEqual(left As IValue, right As IValue) As Boolean
            If left.Type = ValueType.Number Then
                Return left.Number > right.Number OrElse Math.Abs(left.Number - right.Number) < Epsilon
            ElseIf left.Type = ValueType.Str Then
                Return left.Str.CompareTo(right.Str) >= 0
            ElseIf left.Type = ValueType.DateTime Then
                Return left.ToDate.CompareTo(right.ToDate) >= 0
            ElseIf left.Type = ValueType.TimeSpan Then
                Return left.ToTimeSpan.CompareTo(right.ToTimeSpan) >= 0
            Else
                Throw New InvalidOperationException("以上演算子は数値、文字列、日付、時間に対してのみ適用できます。")
            End If
        End Function

        ''' <summary>
        ''' 左辺の値が右辺の値より小さいかどうかを比較します。
        ''' このメソッドは、数値、文字列、日付、時間に対してのみ適用できます。
        ''' </summary>
        ''' <param name="left">左辺の値。</param>
        ''' <param name="right">右辺の値。</param>
        ''' <returns>比較結果。より小さい場合は真。</returns>
        Private Shared Function LessThan(left As IValue, right As IValue) As Boolean
            If left.Type = ValueType.Number Then
                Return left.Number < right.Number AndAlso Math.Abs(left.Number - right.Number) >= Epsilon
            ElseIf left.Type = ValueType.Str Then
                Return left.Str.CompareTo(right.Str) < 0
            ElseIf left.Type = ValueType.DateTime Then
                Return left.ToDate.CompareTo(right.ToDate) < 0
            ElseIf left.Type = ValueType.TimeSpan Then
                Return left.ToTimeSpan.CompareTo(right.ToTimeSpan) < 0
            Else
                Throw New InvalidOperationException("小なり演算子は数値、文字列、日付、時間に対してのみ適用できます。")
            End If
        End Function

        ''' <summary>
        ''' 左辺の値が右辺の値以下かどうかを比較します。
        ''' このメソッドは、数値、文字列、日付、時間に対してのみ適用できます。
        ''' </summary>
        ''' <param name="left">左辺の値。</param>
        ''' <param name="right">右辺の値。</param>
        ''' <returns>比較結果。より小さいか等しい場合は真。</returns>
        Private Shared Function LessEqual(left As IValue, right As IValue) As Boolean
            If left.Type = ValueType.Number Then
                Return left.Number < right.Number OrElse Math.Abs(left.Number - right.Number) < Epsilon
            ElseIf left.Type = ValueType.Str Then
                Return left.Str.CompareTo(right.Str) <= 0
            ElseIf left.Type = ValueType.DateTime Then
                Return left.ToDate.CompareTo(right.ToDate) <= 0
            ElseIf left.Type = ValueType.TimeSpan Then
                Return left.ToTimeSpan.CompareTo(right.ToTimeSpan) <= 0
            Else
                Throw New InvalidOperationException("以上演算子は数値、文字列、日付、時間に対してのみ適用できます。")
            End If
        End Function

        ''' <summary>
        ''' 左辺と右辺の値に対して論理積演算子を適用します。
        ''' このメソッドは、真偽値に対してのみ適用されます。
        ''' </summary>
        ''' <param name="left">左辺の値。</param>
        ''' <param name="right">右辺の値。</param>
        ''' <returns>論理積の結果。</returns>
        Private Shared Function AndOperator(left As IValue, right As IValue) As Boolean
            If left.Type = ValueType.Bool Then
                Return left.Bool AndAlso right.Bool
            Else
                Throw New InvalidOperationException("論理積演算子は真偽値に対してのみ適用できます。")
            End If
        End Function

        ''' <summary>
        ''' 左辺と右辺の値に対して論理和演算子を適用します。
        ''' このメソッドは、真偽値に対してのみ適用されます。
        ''' </summary>
        ''' <param name="left">左辺の値。</param>
        ''' <param name="right">右辺の値。</param>
        ''' <returns>論理和の結果。</returns>
        Private Shared Function OrOperator(left As IValue, right As IValue) As Boolean
            If left.Type = ValueType.Bool Then
                Return left.Bool OrElse right.Bool
            Else
                Throw New InvalidOperationException("論理和演算子は真偽値に対してのみ適用できます。")
            End If
        End Function

        ''' <summary>
        ''' 左辺と右辺の値に対して排他的論理和演算子を適用します。
        ''' このメソッドは、真偽値に対してのみ適用されます。
        ''' </summary>
        ''' <param name="left">左辺の値。</param>
        ''' <param name="right">右辺の値。</param>
        ''' <returns>排他的論理和の結果。</returns>
        Private Shared Function XorOperator(left As IValue, right As IValue) As Boolean
            If left.Type = ValueType.Bool Then
                Return left.Bool Xor right.Bool
            Else
                Throw New InvalidOperationException("排他的論理和演算子は真偽値に対してのみ適用できます。")
            End If
        End Function

    End Class

End Namespace
