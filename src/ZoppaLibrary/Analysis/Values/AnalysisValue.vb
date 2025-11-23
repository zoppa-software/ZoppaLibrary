Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 分析値を表すモジュールです。
    ''' このモジュールは、分析に関連する値の型や操作を定義します。
    ''' </summary>
    ''' <remarks>
    ''' このモジュールは、分析のための値を定義し、他の分析関連の構造体やクラスで使用されます。
    ''' </remarks>
    Public Module AnalysisValue

        ''' <summary>U8StringをIValueに変換します。</summary>
        ''' <param name="value">変換するU8String値。</param>
        ''' <returns>IValue型のStringValue。</returns>
        <Extension()>
        Public Function ToStringValue(value As U8String) As IValue
            Return New StringValue(value)
        End Function

        ''' <summary>文字列をIValueに変換します。</summary>
        ''' <param name="value">変換するU8String値。</param>
        ''' <returns>IValue型のStringValue。</returns>
        <Extension()>
        Public Function ToStringValue(value As String) As IValue
            Return New StringValue(U8String.NewString(value))
        End Function

        ''' <summary>数値をIValueに変換します。</summary>
        ''' <param name="value">変換する数値。</param>
        ''' <returns>IValue型のNumberValue。</returns>
        <Extension()>
        Public Function ToNumberValue(value As Double) As IValue
            Return New NumberValue(value)
        End Function

        ''' <summary>整数をIValueに変換します。</summary>
        ''' <param name="value">変換する整数。</param>
        ''' <returns>IValue型のNumberValue。</returns>
        <Extension()>
        Public Function ToNumberValue(value As Integer) As IValue
            Return New NumberValue(value)
        End Function

        ''' <summary>真偽値をIValueに変換します。</summary>
        ''' <param name="value">変換する真偽値。</param>
        ''' <returns>IValue型のBooleanValue。</returns>
        <Extension()>
        Public Function ToBooleanValue(value As Boolean) As IValue
            Return If(value, BooleanValue.TrueValue, BooleanValue.FalseValue)
        End Function

        ''' <summary>オブジェクトの配列をIValueに変換します。</summary>
        ''' <param name="values">変換するオブジェクトの配列。</param>
        ''' <returns>IValue型のArrayValue。</returns>
        <Extension()>
        Public Function ToArrayValue(Of T)(values As T()) As IValue
            If values Is Nothing Then
                Throw New ArgumentNullException(NameOf(values))
            End If
            Dim items = New IValue(values.Length - 1) {}
            For i As Integer = 0 To values.Length - 1
                items(i) = ConvertToValue(values(i))
            Next
            Return New ArrayValue(items)
        End Function

        ''' <summary>オブジェクトをIValueに変換します。</summary>
        ''' <param name="value">変換するオブジェクト。</param>
        ''' <returns>IValue型のObjectValue。</returns>
        <Extension()>
        Public Function ToObjectValue(value As Object) As IValue
            Return New ObjectValue(value)
        End Function

        ''' <summary>日付時刻をIValueに変換します。</summary>
        ''' <param name="value">変換する日付時刻。</param>
        ''' <returns>IValue型のDateTimeValue。</returns>
        <Extension()>
        Public Function ToDateTimeValue(value As DateTime) As IValue
            Return New DateTimeValue(value)
        End Function

        ''' <summary>時間をIValueに変換します。</summary>
        ''' <param name="value">変換する時間。</param>
        ''' <returns>IValue型のTimeSpanValue。</returns>
        <Extension()>
        Public Function ToTimeSpanValue(value As TimeSpan) As IValue
            Return New TimeSpanValue(value)
        End Function

        ''' <summary>
        ''' オブジェクトをIValueに変換します。
        ''' オブジェクトの型に応じて、適切なIValueを返します。
        ''' 
        ''' 例:
        ''' - ToValue(True) => BooleanValue
        ''' - ToValue(1.0) => NumberValue
        ''' - ToValue("example") => StringValue
        ''' - ToValue([1, 2, 3]) => ArrayValue
        ''' </summary>
        ''' <param name="obj">オブジェクト。</param>
        ''' <returns>IValue。</returns>
        Public Function ConvertToValue(obj As Object) As IValue
            If obj Is Nothing Then
                Return NullValue.Value ' NullValueを返す
            End If
            Select Case obj.GetType()
                Case GetType(Boolean)
                    Return If(CBool(obj), BooleanValue.TrueValue, BooleanValue.FalseValue)
                Case GetType(Integer)
                    Return New NumberValue(CInt(obj))
                Case GetType(Double)
                    Return New NumberValue(CDbl(obj))
                Case GetType(String)
                    Return New StringValue(U8String.NewString(CStr(obj)))
                Case GetType(U8String)
                    Return New StringValue(DirectCast(obj, U8String))
                Case GetType(IValue)
                    Return DirectCast(obj, IValue)
                Case GetType(IValue())
                    Dim tmpo = DirectCast(obj, IValue())
                    Return New ArrayValue(DirectCast(tmpo.Clone(), IValue()))
                Case GetType(DateTime)
                    Return New DateTimeValue(DirectCast(obj, DateTime))
                Case GetType(TimeSpan)
                    Return New TimeSpanValue(DirectCast(obj, TimeSpan))
                Case Else
                    If obj.GetType().IsEnum Then
                        Return New NumberValue(CInt(obj))
                    ElseIf obj.GetType().IsArray Then
                        Dim arr = CType(obj, Array)
                        Dim items = New IValue(arr.Length - 1) {}
                        For i As Integer = 0 To arr.Length - 1
                            items(i) = ConvertToValue(arr.GetValue(i))
                        Next
                        Return New ArrayValue(items)
                    Else
                        Return New ObjectValue(obj)
                    End If
            End Select
        End Function

        ''' <summary>
        ''' IValueを数値に変換します。
        ''' IValueの型に応じて、適切な数値を返します。
        ''' </summary>
        ''' <param name="val">変換するIValue。</param>
        ''' <returns>数値。</returns>
        ''' <exception cref="InvalidOperationException">数値に変換できない場合にスローされます。</exception>
        <Extension()>
        Public Function Number(val As IValue) As Double
            Select Case val.Type
                Case ValueType.Null
                    Return 0 ' Nullは0として扱う
                Case ValueType.Number
                    Return DirectCast(val, NumberValue).Value
                Case ValueType.Bool
                    Return If(DirectCast(val, BooleanValue).Value, 1, 0)
                Case ValueType.Str
                    Dim o = DirectCast(val, StringValue).Value
                    Return ParserModule.ParseNumber(o)
                Case ValueType.Obj
                    Dim o = DirectCast(val, ObjectValue).Value
                    If TypeOf o Is Double Then
                        Return CDbl(o)
                    ElseIf TypeOf o Is Integer Then
                        Return CDbl(CInt(o))
                    Else
                        Throw New InvalidOperationException("オブジェクト値は数値になりません。")
                    End If
                Case Else
                    Throw New InvalidOperationException("数値に変換することができません")
            End Select
        End Function

        ''' <summary>
        ''' IValueを文字列に変換します。
        ''' IValueの型に応じて、適切な文字列を返します。
        ''' </summary>
        ''' <param name="val">変換するIValue。</param>
        ''' <returns>文字列。</returns>
        ''' <exception cref="InvalidOperationException">文字列に変換できない場合にスローされます。</exception>
        <Extension()>
        Public Function Str(val As IValue) As U8String
            Select Case val.Type
                Case ValueType.Null
                    Return U8String.Empty ' Nullは空文字列として扱う
                Case ValueType.Str
                    Return DirectCast(val, StringValue).Value
                Case ValueType.Number
                    Return U8String.NewString(DirectCast(val, NumberValue).Value.ToString())
                Case ValueType.Bool
                    Return If(DirectCast(val, BooleanValue).Value, LexicalModule.TrueKeyword, LexicalModule.FalseKeyword)
                Case ValueType.Obj
                    Dim o = DirectCast(val, ObjectValue).Value
                    If TypeOf o Is U8String Then
                        Return CType(o, U8String)
                    ElseIf TypeOf o Is String Then
                        Return U8String.NewString(CType(o, String))
                    Else
                        Return U8String.NewString(o.ToString())
                    End If
                Case ValueType.Array
                    Dim o = DirectCast(val, ArrayValue).Value
                    Dim res As New List(Of Byte)()
                    For i As Integer = 0 To o.Length - 1
                        If i > 0 Then
                            res.Add(CByte(44)) ' カンマのASCIIコード
                        End If
                        res.AddRange(o(i).Str.GetByteEnumerator())
                    Next
                    Return U8String.NewStringChangeOwner(res.ToArray())
                Case ValueType.DateTime
                    Return ParseU8StringFromDateTime(DirectCast(val, DateTimeValue).Value)
                Case ValueType.TimeSpan
                    Return ParseU8StringFromTimeSpan(DirectCast(val, TimeSpanValue).Value)
                Case Else
                    Throw New InvalidOperationException("文字列に変換することができません")
            End Select
        End Function

        ''' <summary>
        ''' DateTimeをU8Stringに変換します。
        ''' 日付時刻をISO 8601形式の文字列に変換します。
        ''' </summary>
        ''' <param name="dt">変換する日付時刻。</param>
        ''' <returns>ISO 8601形式のU8String。</returns>
        ''' <remarks>例: "2023-10-01T12:34:56.789"</remarks>
        <Extension()>
        Public Function ParseU8StringFromDateTime(dt As DateTime) As U8String
            Dim buf As New List(Of Byte)(32)
            Dim stack As New Stack(Of Byte)(4)

            For Each item In New(dt As Integer, figre As Integer, sprit As Byte)() {
                    (dt.Year, 4, &H2D),
                    (dt.Month, 2, &H2D),
                    (dt.Day, 2, &H54),
                    (dt.Hour, 2, &H3A),
                    (dt.Minute, 2, &H3A),
                    (dt.Second, 2, &H2E),
                    (dt.Millisecond, 3, 0)
                }
                Dim qv = item.dt
                Dim rv As Integer
                stack.Clear()
                While qv > 0
                    qv = Math.DivRem(qv, 10, rv)
                    stack.Push(CByte(rv + &H30)) ' ASCIIコードに変換
                End While

                While stack.Count < item.figre
                    stack.Push(&H30) ' ゼロパディング
                End While

                buf.AddRange(stack)
                If item.sprit <> 0 Then
                    buf.Add(item.sprit) ' スプリット文字を追加
                End If
            Next

            Return U8String.NewStringChangeOwner(buf.ToArray())
        End Function

        ''' <summary>
        ''' TimeSpanをU8Stringに変換します。
        ''' 時間を"DD:HH:MM:SS.MMM"形式の文字列に変換します。
        ''' </summary>
        ''' <param name="ts">変換する時間。</param>
        ''' <returns>"DD:HH:MM:SS.MMM"形式のU8String。</returns>
        ''' <remarks>例: "01:12:34:56.789"</remarks>
        <Extension()>
        Public Function ParseU8StringFromTimeSpan(ts As TimeSpan) As U8String
            Dim buf As New List(Of Byte)(32)
            Dim stack As New Stack(Of Byte)(4)

            Dim list As New List(Of (ts As Integer, figre As Integer, sprit As Byte))()
            If ts.Days > 0 Then
                list.Add((ts.Days, 1, &H54)) ' 日
            End If
            list.Add((ts.Hours, 2, &H3A)) ' 時
            list.Add((ts.Minutes, 2, &H3A)) ' 分
            If ts.Milliseconds > 0 Then
                list.Add((ts.Seconds, 2, &H2E)) ' 秒
                list.Add((ts.Milliseconds, 3, 0)) ' ミリ秒
            Else
                list.Add((ts.Seconds, 2, 0)) ' 秒
            End If

            For Each item In list
                Dim qv = item.ts
                Dim rv As Integer
                stack.Clear()
                While qv > 0
                    qv = Math.DivRem(qv, 10, rv)
                    stack.Push(CByte(rv + &H30)) ' ASCIIコードに変換
                End While
                While stack.Count < item.figre
                    stack.Push(&H30) ' ゼロパディング
                End While
                buf.AddRange(stack)
                If item.sprit <> 0 Then
                    buf.Add(item.sprit) ' スプリット文字を追加
                End If
            Next
            Return U8String.NewStringChangeOwner(buf.ToArray())
        End Function

        ''' <summary>
        ''' IValueを真偽値に変換します。
        ''' IValueの型に応じて、適切な真偽値を返します。
        ''' </summary>
        ''' <param name="val">変換するIValue。</param>
        ''' <returns>真偽値。</returns>
        ''' <exception cref="InvalidOperationException">真偽値に変換できない場合にスローされます。</exception>
        <Extension()>
        Public Function Bool(val As IValue) As Boolean
            Select Case val.Type
                Case ValueType.Null
                    Return False ' Nullは偽とみなす
                Case ValueType.Bool
                    Return DirectCast(val, BooleanValue).Value
                Case ValueType.Str
                    Dim o = DirectCast(val, StringValue).Value
                    If o = TrueKeyword Then
                        Return True
                    ElseIf o = FalseKeyword Then
                        Return False
                    Else
                        ' 文字列が真偽値として解釈できない場合は例外を投げる
                        Throw New InvalidOperationException("文字列を真偽値として解釈できません。")
                    End If
                Case ValueType.Number
                    Return DirectCast(val, NumberValue).Value <> 0
                Case ValueType.Obj
                    Dim o = DirectCast(val, ObjectValue).Value
                    If TypeOf o Is Boolean Then
                        Return CType(o, Boolean)
                    Else
                        Throw New InvalidOperationException("オブジェクト値を真偽値として解釈できません。")
                    End If
                Case Else
                    Throw New InvalidOperationException("真偽値に変換することができません")
            End Select
        End Function

        ''' <summary>
        ''' IValueを配列値に変換します。
        ''' IValueの型に応じて、適切な配列値を返します。
        ''' </summary>
        ''' <param name="val">変換するIValue。</param>
        ''' <returns>配列値。</returns>
        <Extension()>
        Public Function Array(val As IValue) As IValue()
            Select Case val.Type
                Case ValueType.Null
                    Return New IValue() {}
                Case ValueType.Array
                    Return DirectCast(val, ArrayValue).Value
                Case ValueType.Obj
                    Dim o = DirectCast(val, ObjectValue).Value
                    If TypeOf o Is IValue() Then
                        Return DirectCast(o, IValue())
                    ElseIf TypeOf o Is IEnumerable Then
                        Dim items As New List(Of IValue)()
                        For Each item In CType(o, IEnumerable)
                            items.Add(ConvertToValue(item))
                        Next
                        Return items.ToArray()
                    Else
                        Return New IValue() {val}
                    End If
                Case Else
                    Return New IValue() {val}
            End Select
        End Function

        ''' <summary>
        ''' IValueをオブジェクト値に変換します。
        ''' IValueの型に応じて、適切なオブジェクト値を返します。
        ''' </summary>
        ''' <param name="val">変換するIValue。</param>
        ''' <returns>配列オブジェクト値。</returns>
        <Extension()>
        Public Function Obj(val As IValue) As Object
            Select Case val.Type
                Case ValueType.Null
                    Return Nothing
                Case ValueType.Array
                    Return DirectCast(val, ArrayValue).Value
                Case ValueType.Bool
                    Return DirectCast(val, BooleanValue).Value
                Case ValueType.Number
                    Return DirectCast(val, NumberValue).Value
                Case ValueType.Str
                    Return DirectCast(val, StringValue).Value
                Case ValueType.Obj
                    Return DirectCast(val, ObjectValue).Value
                Case ValueType.DateTime
                    Return DirectCast(val, DateTimeValue).Value
                Case ValueType.TimeSpan
                    Return DirectCast(val, TimeSpanValue).Value
                Case Else
                    Throw New InvalidOperationException("オブジェクト値に変換することができません")
            End Select
        End Function

        ''' <summary>
        ''' IValueを日付時刻に変換します。
        ''' IValueの型に応じて、適切な日付時刻を返します。
        ''' </summary>
        ''' <param name="val">変換するIValue。</param>
        ''' <returns>日付時刻。</returns>
        ''' <exception cref="InvalidOperationException">日付に変換できない場合にスローされます。</exception>
        <Extension()>
        Public Function ToDate(val As IValue) As DateTime
            Select Case val.Type
                Case ValueType.Null
                    Return DateTime.MinValue ' Nullは最小値として扱う
                Case ValueType.DateTime
                    Return DirectCast(val, DateTimeValue).Value
                Case ValueType.Str
                    Try
                        Return ParseDateTime(DirectCast(val, StringValue).Value)
                    Catch ex As Exception
                        Throw New InvalidOperationException("日付に変換できません", ex)
                    End Try
                Case Else
                    Throw New InvalidOperationException("日付に変換することができません")
            End Select
        End Function

        ''' <summary>
        ''' IValueを時間に変換します。
        ''' IValueの型に応じて、適切な時間を返します。
        ''' </summary>
        ''' <param name="val">変換するIValue。</param>
        ''' <returns>時間。</returns>
        ''' <exception cref="InvalidOperationException">時間に変換できない場合にスローされます。</exception>
        <Extension()>
        Public Function ToTimeSpan(val As IValue) As TimeSpan
            Select Case val.Type
                Case ValueType.Null
                    Return TimeSpan.Zero ' Nullはゼロとして扱う
                Case ValueType.TimeSpan
                    Return DirectCast(val, TimeSpanValue).Value
                Case ValueType.Str
                    Try
                        Return ParseTimeSpan(DirectCast(val, StringValue).Value)
                    Catch ex As Exception
                        Throw New InvalidOperationException("時間に変換できません", ex)
                    End Try
                Case Else
                    Throw New InvalidOperationException("時間に変換することができません")
            End Select
        End Function

    End Module

End Namespace
