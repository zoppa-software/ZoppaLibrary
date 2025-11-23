Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 変数の値を解析するための拡張メソッドを提供するモジュールです。
    ''' このモジュールは、変数の型に応じて値を真偽値、数値、文字列に変換する機能を提供します。
    ''' </summary>
    Public Module AnalysisVariable

        ''' <summary>
        ''' 変数の値を真偽値に変換します。
        ''' 変数の型に応じて、適切な変換を行います。
        ''' 文字列の場合は、"true" または "false" のキーワードを使用して変換します。
        ''' 数値の場合は、0以外を真とし、0を偽とします。
        ''' 
        ''' 例:
        ''' - Bool(VariableBool(True)) => True
        ''' - Bool(VariableNumber(1)) => True
        ''' - Bool(VariableStr("true")) => True
        ''' - Bool(VariableStr("false")) => False
        ''' </summary>
        ''' <param name="value">変数。</param>
        ''' <returns>真偽値。</returns>
        <Extension()>
        Public Function Bool(value As IVariable) As Boolean
            Select Case value.Type
                Case VariableType.Bool
                    Return DirectCast(value, BooleanVariable).Value
                Case VariableType.Number
                    Dim nv = DirectCast(value, NumberVariable).Value
                    Return nv <> 0
                Case VariableType.Str
                    Dim sv = DirectCast(value, StringVariable).Value
                    Select Case sv
                        Case LexicalModule.TrueKeyword
                            Return True
                        Case LexicalModule.FalseKeyword
                            Return False
                        Case Else
                            Throw New InvalidCastException("文字列を真偽値に変換できません。")
                    End Select
                Case Else
                    Throw New InvalidCastException("真偽値に変換できませんでした。")
            End Select
        End Function

        ''' <summary>
        ''' 変数の値を文字列に変換します。
        ''' 変数の型に応じて、適切な変換を行います。
        ''' 
        ''' 例:
        ''' - Str(VariableBool(True)) => "true"
        ''' - Str(VariableNumber(1)) => "1"
        ''' - Str(VariableStr("example")) => "example"
        ''' </summary>
        ''' <param name="value">変数。</param>
        ''' <returns>文字列。</returns>
        <Extension()>
        Public Function Number(value As IVariable) As Double
            Select Case value.Type
                Case VariableType.Bool
                    Dim bv = DirectCast(value, BooleanVariable).Value
                    Return If(bv, 1.0, 0.0)
                Case VariableType.Number
                    Return DirectCast(value, NumberVariable).Value
                Case VariableType.Str
                    Return ParserModule.ParseNumber(DirectCast(value, StringVariable).Value)
                Case Else
                    Throw New InvalidCastException("数値に変換できませんでした。")
            End Select
        End Function

        ''' <summary>
        ''' 変数の値を文字列に変換します。
        ''' 変数の型に応じて、適切な変換を行います。
        ''' 
        ''' 例:
        ''' - Str(VariableBool(True)) => "true"
        ''' - Str(VariableNumber(1)) => "1"
        ''' - Str(VariableStr("example")) => "example"
        ''' - Str(VariableArray([1, 2, 3])) => "1,2,3"
        ''' </summary>
        ''' <param name="value">変数。</param>
        ''' <param name="venv"> 解析環境。</param>
        ''' <returns>文字列。</returns>
        <Extension()>
        Public Function Str(value As IVariable, venv As AnalysisEnvironment) As U8String
            Select Case value.Type
                Case VariableType.Bool
                    Dim bv = DirectCast(value, BooleanVariable).Value
                    Return If(bv, LexicalModule.TrueKeyword, LexicalModule.FalseKeyword)

                Case VariableType.Number
                    Return U8String.NewString(DirectCast(value, NumberVariable).Value.ToString())

                Case VariableType.Str
                    Return DirectCast(value, StringVariable).Value

                Case VariableType.Array
                    Dim av = DirectCast(value, ArrayVariable).Value
                    Dim res As New List(Of Byte)()
                    For i As Integer = 0 To av.Length - 1
                        If i > 0 Then
                            res.Add(CByte(44)) ' カンマのASCIIコード
                        End If
                        res.AddRange(av(i).GetValue(venv).Str.GetByteEnumerator())
                    Next
                    Return U8String.NewStringChangeOwner(res.ToArray())

                Case VariableType.Obj
                    Dim obj = DirectCast(value, ObjectVariable).Value
                    If TypeOf obj Is U8String Then
                        Return DirectCast(obj, U8String)
                    ElseIf TypeOf obj Is String Then
                        Return U8String.NewString(obj.ToString())
                    ElseIf TypeOf obj Is IValue Then
                        Return U8String.NewString(obj.ToString())
                    Else
                        Throw New InvalidCastException("オブジェクトを文字列に変換できません。")
                    End If

                Case VariableType.Date
                    Return U8String.NewString(DirectCast(value, DateTimeVariable).Value.ToString("yyyy-MM-dd HH:mm:ss"))

                Case VariableType.Time
                    Return U8String.NewString(DirectCast(value, TimeSpanVariable).Value.ToString("hh\:mm\:ss"))

                Case Else
                    Throw New InvalidCastException("文字列に変換できませんでした。")
            End Select
        End Function

        ''' <summary>
        ''' 変数の値を日付に変換します。
        ''' 変数の型に応じて、適切な変換を行います。
        ''' </summary>
        ''' <param name="value">変数。</param>
        ''' <returns>日付。</returns>
        <Extension()>
        Public Function ToDate(value As IVariable) As DateTime
            Select Case value.Type
                Case VariableType.Date
                    Return DirectCast(value, DateTimeVariable).Value
                Case Else
                    Throw New InvalidCastException("日付に変換できませんでした。")
            End Select
        End Function

        ''' <summary>
        ''' 変数の値を時間に変換します。
        ''' 変数の型に応じて、適切な変換を行います。
        ''' </summary>
        ''' <param name="value">変数。</param>
        ''' <returns>時間。</returns>
        <Extension()>
        Public Function ToTime(value As IVariable) As TimeSpan
            Select Case value.Type
                Case VariableType.Time
                    Return DirectCast(value, TimeSpanVariable).Value
                Case Else
                    Throw New InvalidCastException("時間に変換できませんでした。")
            End Select
        End Function

        ''' <summary>
        ''' オブジェクトを変数に変換します。
        ''' オブジェクトの型に応じて、適切な変数型に変換します。
        ''' </summary>
        ''' <param name="obj">変換するオブジェクト。</param>
        ''' <returns>変数。</returns>
        Function ConvertToVariable(obj As Object) As IVariable
            If obj Is Nothing Then
                Return New ObjectVariable(Nothing)　' NullValueを返す
            End If
            Select Case obj.GetType()
                Case GetType(Boolean)
                    Return If(CBool(obj), BooleanVariable.TrueValue, BooleanVariable.FalseValue)
                Case GetType(Integer)
                    Return New NumberVariable(CInt(obj))
                Case GetType(Double)
                    Return New NumberVariable(CDbl(obj))
                Case GetType(String)
                    Return New StringVariable(U8String.NewString(CStr(obj)))
                Case GetType(U8String)
                    Return New StringVariable(DirectCast(obj, U8String))
                Case GetType(IVariable)
                    Return DirectCast(obj, IVariable)
                Case GetType(DateTime)
                    Return New DateTimeVariable(DirectCast(obj, DateTime))
                Case GetType(TimeSpan)
                    Return New TimeSpanVariable(DirectCast(obj, TimeSpan))
                Case Else
                    If obj.GetType().IsEnum Then
                        Return New NumberVariable(CInt(obj))
                    ElseIf obj.GetType().IsArray Then
                        Dim arr = CType(obj, Array)
                        Dim items = New IExpression(arr.Length - 1) {}
                        For i As Integer = 0 To arr.Length - 1
                            items(i) = ConvertToExpression(arr.GetValue(i))
                        Next
                        Return New ArrayVariable(items.ToArray())
                    Else
                        Return New ObjectVariable(obj)
                    End If
            End Select
        End Function

        ''' <summary>
        ''' オブジェクトを式に変換します。
        ''' オブジェクトの型に応じて、適切な式型に変換します。
        ''' </summary>
        ''' <param name="obj">変換するオブジェクト。</param>
        ''' <returns>式。</returns>
        Function ConvertToExpression(obj As Object) As IExpression
            If obj Is Nothing Then
                Return New ObjectExpression(Nothing) ' NullValueを返す
            End If
            Select Case obj.GetType()
                Case GetType(Boolean)
                    Return If(CBool(obj), BooleanExpression.TrueValue, BooleanExpression.FalseValue)
                Case GetType(Double)
                    Return New NumberExpression(CDbl(obj))
                Case GetType(String)
                    Return New StringExpression(U8String.NewString(CStr(obj)))
                Case GetType(U8String)
                    Return New StringExpression(DirectCast(obj, U8String))
                Case GetType(IExpression)
                    Return DirectCast(obj, IExpression)
                Case GetType(DateTime)
                    Return New DateTimeExpression(DirectCast(obj, DateTime))
                Case GetType(TimeSpan)
                    Return New TimeSpanExpression(DirectCast(obj, TimeSpan))
                Case Else
                    If obj.GetType().IsArray Then
                        Dim arr = CType(obj, Array)
                        Dim items = New IExpression(arr.Length - 1) {}
                        For i As Integer = 0 To arr.Length - 1
                            items(i) = ConvertToExpression(arr.GetValue(i))
                        Next
                        Return New ArrayFieldExpression(items.ToArray())
                    Else
                        Return New ObjectExpression(obj)
                    End If
            End Select
        End Function

        ''' <summary>
        ''' 変数を値に変換します。
        ''' 変数の型に応じて、適切な値型に変換します。
        ''' </summary>
        ''' <param name="value">変数。</param>
        ''' <param name="venv">解析環境。</param>
        ''' <returns>値。</returns>
        <Extension()>
        Function ToValue(value As IVariable, venv As AnalysisEnvironment) As IValue
            Select Case value.Type
                Case VariableType.Expr
                    Return DirectCast(value, ExprVariable).Value.GetValue(venv)
                Case VariableType.Bool
                    Return If(DirectCast(value, BooleanVariable).Value,
                        BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case VariableType.Number
                    Return New NumberValue(DirectCast(value, NumberVariable).Value)
                Case VariableType.Str
                    Return New StringValue(DirectCast(value, StringVariable).Value)
                Case VariableType.Array
                    Dim arr = DirectCast(value, ArrayVariable).Value.Select(Function(i) i.GetValue(venv)).ToArray()
                    Return New ArrayValue(arr)
                Case VariableType.Obj
                    Return New ObjectValue(DirectCast(value, ObjectVariable).Value)
                Case VariableType.Date
                    Return New DateTimeValue(DirectCast(value, DateTimeVariable).Value)
                Case VariableType.Time
                    Return New TimeSpanValue(DirectCast(value, TimeSpanVariable).Value)
                Case Else
                    Throw New InvalidOperationException("サポートされていない変数の型です。")
            End Select
        End Function

        ''' <summary>
        ''' 値を変数に変換します。
        ''' 値の型に応じて、適切な変数型に変換します。
        ''' </summary>
        ''' <param name="value">値。</param>
        ''' <returns>変数。</returns>
        <Extension()>
        Function ToVariable(value As IValue) As IVariable
            Select Case value.Type
                Case ValueType.Str
                    Return New StringVariable(value.Str)
                Case ValueType.Number
                    Return New NumberVariable(value.Number)
                Case ValueType.Bool
                    Return If(value.Bool, BooleanVariable.TrueValue, BooleanVariable.FalseValue)
                Case ValueType.Obj
                    Return New ObjectVariable(value.Obj)
                Case ValueType.DateTime
                    Return New DateTimeVariable(value.ToDate)
                Case ValueType.TimeSpan
                    Return New TimeSpanVariable(value.ToTimeSpan)
                Case ValueType.Array
                    Dim vars As New List(Of IExpression)()
                    For Each item In value.Array
                        vars.Add(item.ToExpression())
                    Next
                    Return New ArrayVariable(vars.ToArray)
                Case Else
                    Throw New AnalysisException("値から変数に変換できません。")
            End Select
        End Function

        ''' <summary>
        ''' 値を式に変換します。
        ''' 値の型に応じて、適切な式型に変換します。
        ''' </summary>
        ''' <param name="value">値。</param>
        ''' <returns>式。</returns>
        <Extension()>
        Function ToExpression(value As IValue) As IExpression
            Select Case value.Type
                Case ValueType.Str
                    Return New StringExpression(value.Str)
                Case ValueType.Number
                    Return New NumberExpression(value.Number)
                Case ValueType.Bool
                    Return If(value.Bool, BooleanExpression.TrueValue, BooleanExpression.FalseValue)
                Case ValueType.Obj
                    Return New ObjectExpression(value.Obj)
                Case ValueType.DateTime
                    Return New DateTimeExpression(value.ToDate)
                Case ValueType.TimeSpan
                    Return New TimeSpanExpression(value.ToTimeSpan)
                Case ValueType.Array
                    Dim vars As New List(Of IExpression)()
                    For Each item In value.Array
                        vars.Add(item.ToExpression())
                    Next
                    Return New ArrayFieldExpression(vars.ToArray())
                Case Else
                    Throw New AnalysisException("値から変数に変換できません。")
            End Select
        End Function

    End Module

End Namespace
