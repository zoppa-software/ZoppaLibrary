Imports System
Imports Xunit
Imports ZoppaLibrary
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.Strings
Imports ZoppaLibrary.Analysis.AnalysisValue

Public Class ParseFactorTest

    <Fact>
    Public Sub TestParseNumverFactor()
        Dim result1 = Analysis.ParserModule.Parse("123.456")
        Assert.Equal(123.456, result1.Expression.GetValue(Nothing).Number)

        Dim result2 = Analysis.ParserModule.Parse("0.001")
        Assert.Equal(0.001, result2.Expression.GetValue(Nothing).Number)

        Dim result3 = Analysis.ParserModule.Parse("1000")
        Assert.Equal(1000.0, result3.Expression.GetValue(Nothing).Number)

        Dim result4 = Analysis.ParserModule.Parse("0")
        Assert.Equal(0.0, result4.Expression.GetValue(Nothing).Number)

        Dim result5 = Analysis.ParserModule.Parse("123_456.789")
        Assert.Equal(123456.789, result5.Expression.GetValue(Nothing).Number)
    End Sub

    <Fact>
    Public Sub TestParseStringFactor()
        Dim result1 = Analysis.ParserModule.Parse("""Hello, World!""")
        Assert.True(result1.Expression.GetValue(Nothing).Str.Equals("Hello, World!"))

        Dim result2 = Analysis.ParserModule.Parse("""12345""")
        Assert.True(result2.Expression.GetValue(Nothing).Str.Equals("12345"))

        Dim result3 = Analysis.ParserModule.Parse("""Test String with spaces""")
        Assert.True(result3.Expression.GetValue(Nothing).Str.Equals("Test String with spaces"))

        Dim result4 = Analysis.ParserModule.Parse("""Special characters !@#$%^&*()""")
        Assert.True(result4.Expression.GetValue(Nothing).Str.Equals("Special characters !@#$%^&*()"))

        Dim result5 = Analysis.ParserModule.Parse("""\""""")
        Assert.True(result5.Expression.GetValue(Nothing).Str.Equals(""""))

        Dim result6 = Analysis.ParserModule.Parse("'\''")
        Assert.True(result6.Expression.GetValue(Nothing).Str.Equals("'"))
    End Sub

    <Fact>
    Public Sub TestParseBoolFactor()
        Dim result1 = Analysis.ParserModule.Parse("true")
        Assert.True(result1.Expression.GetValue(Nothing).Bool)
        Dim result2 = Analysis.ParserModule.Parse("false")
        Assert.False(result2.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestParseUnaryFactor()
        Dim result1 = Analysis.ParserModule.Parse("-123.456")
        Assert.Equal(-123.456, result1.Expression.GetValue(Nothing).Number)
        Dim result2 = Analysis.ParserModule.Parse("+123.456")
        Assert.Equal(123.456, result2.Expression.GetValue(Nothing).Number)
        Dim result3 = Analysis.ParserModule.Parse("not true")
        Assert.False(result3.Expression.GetValue(Nothing).Bool)
        Dim result4 = Analysis.ParserModule.Parse("not false")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestParseParenFactor()
        Dim result1 = Analysis.ParserModule.Parse("(123.456)")
        Assert.Equal(123.456, result1.Expression.GetValue(Nothing).Number)
        Dim result2 = Analysis.ParserModule.Parse("(true)")
        Assert.True(result2.Expression.GetValue(Nothing).Bool)
        Dim result3 = Analysis.ParserModule.Parse("(-123.456)")
        Assert.Equal(-123.456, result3.Expression.GetValue(Nothing).Number)
        Dim result4 = Analysis.ParserModule.Parse("(not false)")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestParseArrayFactor()
        Dim result1 = Analysis.ParserModule.Parse("[1, 2, 3]")
        Dim arrayValue1 = result1.Expression.GetValue(Nothing).Array
        Assert.Equal(3, arrayValue1.Length)
        Assert.Equal(1.0, arrayValue1(0).Number)
        Assert.Equal(2.0, arrayValue1(1).Number)
        Assert.Equal(3.0, arrayValue1(2).Number)
        Dim result2 = Analysis.ParserModule.Parse("[true, false, true]")
        Dim arrayValue2 = result2.Expression.GetValue(Nothing).Array
        Assert.Equal(3, arrayValue2.Length)
        Assert.True(arrayValue2(0).Bool)
        Assert.False(arrayValue2(1).Bool)
        Assert.True(arrayValue2(2).Bool)
        Dim result3 = Analysis.ParserModule.Parse("['a', 'b', 'c']")
        Dim arrayValue3 = result3.Expression.GetValue(Nothing).Array
        Assert.Equal(3, arrayValue3.Length)
        Assert.True(arrayValue3(0).Str.Equals("a"))
        Assert.True(arrayValue3(1).Str.Equals("b"))
        Assert.True(arrayValue3(2).Str.Equals("c"))

        Assert.Throws(Of Analysis.AnalysisException)(
            Sub()
                Analysis.ParserModule.Parse("[1, , 3,]")
            End Sub
        )
    End Sub

    <Fact>
    Public Sub TestFunctionCall()
        Dim venv As New Analysis.AnalysisEnvironment()
        venv.AddFunction(
            U8String.NewString("test"),
            Function(name)
                Return U8String.NewString("こんにちは! ").Concat(name.Str).ToStringValue()
            End Function
        )
        Dim result1 = Analysis.ParserModule.Parse("test('崇')")
        Assert.True(result1.Expression.GetValue(venv).Str.Equals("こんにちは! 崇"))
    End Sub

    <Fact>
    Public Sub TestVariableFactor()
        Dim venv As New Analysis.AnalysisEnvironment()

        venv.RegisterNumber("x", 42)
        Dim result1 = Analysis.ParserModule.Parse("x")
        Assert.Equal(42.0, result1.Expression.GetValue(venv).Number)

        venv.RegisterBool("y", True)
        Dim result2 = Analysis.ParserModule.Parse("y")
        Assert.True(result2.Expression.GetValue(venv).Bool)

        venv.RegisterStr("z", "Hello")
        Dim result3 = Analysis.ParserModule.Parse("z")
        Assert.True(result3.Expression.GetValue(venv).Str.Equals("Hello"))
    End Sub

    <Fact>
    Public Sub TestArrayAccessFactor()
        Dim venv As New Analysis.AnalysisEnvironment()
        venv.RegisterArray("arr", 1, 2, 3)
        Dim result1 = Analysis.ParserModule.Parse("arr[0]")
        Assert.Equal(1.0, result1.Expression.GetValue(venv).Number)
        Dim result2 = Analysis.ParserModule.Parse("arr[1]")
        Assert.Equal(2.0, result2.Expression.GetValue(venv).Number)
        Dim result3 = Analysis.ParserModule.Parse("arr[2]")
        Assert.Equal(3.0, result3.Expression.GetValue(venv).Number)
    End Sub

    Class Test
        Public Property Name As String
        Public Property Age As Integer
        Public Property IsActive As Boolean
        Public Property Scores As Integer()
    End Class

    <Fact>
    Public Sub TestFieldAccessFactor()
        Dim venv As New Analysis.AnalysisEnvironment()
        Dim testObj As New Test With {
            .Name = "崇",
            .Age = 49,
            .IsActive = True,
            .Scores = New Integer() {90, 80, 70}
        }
        venv.RegisterObject("testObj", testObj)
        Dim result1 = Analysis.ParserModule.Parse("testObj.Name")
        Assert.True(result1.Expression.GetValue(venv).Str.Equals("崇"))
        Dim result2 = Analysis.ParserModule.Parse("testObj.Age")
        Assert.Equal(49, result2.Expression.GetValue(venv).Number)
        Dim result3 = Analysis.ParserModule.Parse("testObj.IsActive")
        Assert.True(result3.Expression.GetValue(venv).Bool)
        Dim result4 = Analysis.ParserModule.Parse("testObj.Scores[0]")
        Assert.Equal(90.0, result4.Expression.GetValue(venv).Number)
    End Sub

    <Fact>
    Public Sub TestCalc()
        ' 数値の乗算解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("123.456 * 2")
        Assert.Equal(246.912, result1.Expression.GetValue(Nothing).Number)
        Dim result2 = Analysis.ParserModule.Parse("0.001 * 1000")
        Assert.Equal(1.0, result2.Expression.GetValue(Nothing).Number)
        Dim result3 = Analysis.ParserModule.Parse("1000 * 0")
        Assert.Equal(0.0, result3.Expression.GetValue(Nothing).Number)
        Dim result4 = Analysis.ParserModule.Parse("123_456.789 * 2")
        Assert.Equal(246913.578, result4.Expression.GetValue(Nothing).Number)

        ' 数値の除算解析をテスト
        Dim result5 = Analysis.ParserModule.Parse("123.456 / 2")
        Assert.Equal(61.728, result5.Expression.GetValue(Nothing).Number)
        Dim result6 = Analysis.ParserModule.Parse("0.001 / 1000")
        Assert.Equal(0.000001, result6.Expression.GetValue(Nothing).Number)
        Dim result7 = Analysis.ParserModule.Parse("1000 / 2")
        Assert.Equal(500.0, result7.Expression.GetValue(Nothing).Number)
        Dim result8 = Analysis.ParserModule.Parse("123_456.789 / 3")
        Assert.Equal(41152.263, result8.Expression.GetValue(Nothing).Number)
        Dim result9 = Analysis.ParserModule.Parse("1000 / 0")
        Assert.Throws(Of DivideByZeroException)(Function() result9.Expression.GetValue(Nothing))

        ' 数値の加算解析をテスト
        Dim result10 = Analysis.ParserModule.Parse("123.456 + 2")
        Assert.Equal(125.456, result10.Expression.GetValue(Nothing).Number)
        Dim result11 = Analysis.ParserModule.Parse("0.001 + 1000")
        Assert.Equal(1000.001, result11.Expression.GetValue(Nothing).Number)
        Dim result12 = Analysis.ParserModule.Parse("1000 + 0")
        Assert.Equal(1000.0, result12.Expression.GetValue(Nothing).Number)
        Dim result13 = Analysis.ParserModule.Parse("123_456.789 + 2")
        Assert.Equal(123458.789, result13.Expression.GetValue(Nothing).Number)

        ' 文字列の加算解析をテスト
        Dim resultString1 = Analysis.ParserModule.Parse("""Hello"" + "" World!""")
        Assert.True(resultString1.Expression.GetValue(Nothing).Str.Equals("Hello World!"))

        ' 数値の減算解析をテスト
        Dim result14 = Analysis.ParserModule.Parse("123.456 - 2")
        Assert.Equal(121.456, result14.Expression.GetValue(Nothing).Number)
        Dim result15 = Analysis.ParserModule.Parse("0.001 - 1000")
        Assert.Equal(-999.999, result15.Expression.GetValue(Nothing).Number)
        Dim result16 = Analysis.ParserModule.Parse("1000 - 0")
        Assert.Equal(1000.0, result16.Expression.GetValue(Nothing).Number)
        Dim result17 = Analysis.ParserModule.Parse("123_456.789 - -2")
        Assert.Equal(123458.789, result17.Expression.GetValue(Nothing).Number)
    End Sub

    <Fact>
    Public Sub TestEquals()
        ' 等価演算子の解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("123.456 == 123.456")
        Assert.True(result1.Expression.GetValue(Nothing).Bool)
        Dim result2 = Analysis.ParserModule.Parse("123.456 == 123.457")
        Assert.False(result2.Expression.GetValue(Nothing).Bool)
        Dim result3 = Analysis.ParserModule.Parse("0.001 == 0.001")
        Assert.True(result3.Expression.GetValue(Nothing).Bool)
        Dim result4 = Analysis.ParserModule.Parse("1000 == 1000")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)

        ' 文字列の等価演算子の解析をテスト
        Dim resultString1 = Analysis.ParserModule.Parse("""Hello"" == ""Hello""")
        Assert.True(resultString1.Expression.GetValue(Nothing).Bool)
        Dim resultString2 = Analysis.ParserModule.Parse("""Hello"" == ""World!""")
        Assert.False(resultString2.Expression.GetValue(Nothing).Bool)
        Dim resultString3 = Analysis.ParserModule.Parse("""Hello"" == ""Hello World!""")
        Assert.False(resultString3.Expression.GetValue(Nothing).Bool)

        ' 真偽値の等価演算子の解析をテスト
        Dim resultBool1 = Analysis.ParserModule.Parse("true == true")
        Assert.True(resultBool1.Expression.GetValue(Nothing).Bool)
        Dim resultBool2 = Analysis.ParserModule.Parse("true == false")
        Assert.False(resultBool2.Expression.GetValue(Nothing).Bool)
        Dim resultBool3 = Analysis.ParserModule.Parse("false == false")
        Assert.True(resultBool3.Expression.GetValue(Nothing).Bool)

        ' 配列の等価演算子の解析をテスト
        Dim resultArray1 = Analysis.ParserModule.Parse("[1, 2, 3] == [1, 2, 3]")
        Assert.True(resultArray1.Expression.GetValue(Nothing).Bool)
        Dim resultArray2 = Analysis.ParserModule.Parse("[1, 2, 3] == [4, 5, 6]")
        Assert.False(resultArray2.Expression.GetValue(Nothing).Bool)
        Dim resultArray3 = Analysis.ParserModule.Parse("[1, 2, 3] == [1, 2, 3, 4]")
        Assert.False(resultArray3.Expression.GetValue(Nothing).Bool)
        Dim resultArray4 = Analysis.ParserModule.Parse("[1, 2, 3] == [1, 2]")
        Assert.False(resultArray4.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestNotEquals()
        ' 非等価演算子の解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("123.456 <> 123.457")
        Assert.True(result1.Expression.GetValue(Nothing).Bool)
        Dim result2 = Analysis.ParserModule.Parse("123.456 <> 123.456")
        Assert.False(result2.Expression.GetValue(Nothing).Bool)
        Dim result3 = Analysis.ParserModule.Parse("0.001 <> 0.002")
        Assert.True(result3.Expression.GetValue(Nothing).Bool)
        Dim result4 = Analysis.ParserModule.Parse("1000 <> 999")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)
        ' 文字列の非等価演算子の解析をテスト
        Dim resultString1 = Analysis.ParserModule.Parse("""Hello"" <> ""World!""")
        Assert.True(resultString1.Expression.GetValue(Nothing).Bool)
        Dim resultString2 = Analysis.ParserModule.Parse("""Hello"" <> ""Hello""")
        Assert.False(resultString2.Expression.GetValue(Nothing).Bool)
        ' 真偽値の非等価演算子の解析をテスト
        Dim resultBool1 = Analysis.ParserModule.Parse("true <> false")
        Assert.True(resultBool1.Expression.GetValue(Nothing).Bool)
        Dim resultBool2 = Analysis.ParserModule.Parse("true <> true")
        Assert.False(resultBool2.Expression.GetValue(Nothing).Bool)
        ' 配列の非等価演算子の解析をテスト
        Dim resultArray1 = Analysis.ParserModule.Parse("[1, 2, 3] <> [4, 5, 6]")
        Assert.True(resultArray1.Expression.GetValue(Nothing).Bool)
        Dim resultArray2 = Analysis.ParserModule.Parse("[1, 2, 3] <> [1, 2, 3]")
        Assert.False(resultArray2.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestGreaterThan()
        ' 大なり演算子の解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("123.456 > 123.455")
        Assert.True(result1.Expression.GetValue(Nothing).Bool)
        Dim result2 = Analysis.ParserModule.Parse("123.456 > 123.456")
        Assert.False(result2.Expression.GetValue(Nothing).Bool)
        Dim result3 = Analysis.ParserModule.Parse("0.001 > 0.000999")
        Assert.True(result3.Expression.GetValue(Nothing).Bool)
        Dim result4 = Analysis.ParserModule.Parse("1000 > 999")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)
        ' 文字列の大なり演算子の解析をテスト
        Dim resultString1 = Analysis.ParserModule.Parse("""Hello"" > ""Hello""")
        Assert.False(resultString1.Expression.GetValue(Nothing).Bool)
        Dim resultString2 = Analysis.ParserModule.Parse("""Hello"" > ""World!""")
        Assert.False(resultString2.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestLessThan()
        ' 小なり演算子の解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("123.456 < 123.457")
        Assert.True(result1.Expression.GetValue(Nothing).Bool)
        Dim result2 = Analysis.ParserModule.Parse("123.456 < 123.456")
        Assert.False(result2.Expression.GetValue(Nothing).Bool)
        Dim result3 = Analysis.ParserModule.Parse("0.001 < 0.001001")
        Assert.True(result3.Expression.GetValue(Nothing).Bool)
        Dim result4 = Analysis.ParserModule.Parse("999 < 1000")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)
        ' 文字列の小なり演算子の解析をテスト
        Dim resultString1 = Analysis.ParserModule.Parse("""Hello"" < ""World!""")
        Assert.True(resultString1.Expression.GetValue(Nothing).Bool)
        Dim resultString2 = Analysis.ParserModule.Parse("""World!"" < ""Hello""")
        Assert.False(resultString2.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestGreaterThanOrEqual()
        ' 大なりイコール演算子の解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("123.456 >= 123.456")
        Assert.True(result1.Expression.GetValue(Nothing).Bool)
        Dim result2 = Analysis.ParserModule.Parse("123.456 >= 123.455")
        Assert.True(result2.Expression.GetValue(Nothing).Bool)
        Dim result3 = Analysis.ParserModule.Parse("0.001 >= 0.001")
        Assert.True(result3.Expression.GetValue(Nothing).Bool)
        Dim result4 = Analysis.ParserModule.Parse("1000 >= 999")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestLessThanOrEqual()
        ' 小なりイコール演算子の解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("123.456 <= 123.456")
        Assert.True(result1.Expression.GetValue(Nothing).Bool)
        Dim result2 = Analysis.ParserModule.Parse("123.456 <= 123.457")
        Assert.True(result2.Expression.GetValue(Nothing).Bool)
        Dim result3 = Analysis.ParserModule.Parse("0.001 <= 0.001")
        Assert.True(result3.Expression.GetValue(Nothing).Bool)
        Dim result4 = Analysis.ParserModule.Parse("999 <= 1000")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)

        Dim result5 = Analysis.ParserModule.Parse("'ABC' <= 'ABG'")
        Assert.True(result5.Expression.GetValue(Nothing).Bool)
        Dim result6 = Analysis.ParserModule.Parse("'ABG' <= 'ABC'")
        Assert.False(result6.Expression.GetValue(Nothing).Bool)
        Dim result7 = Analysis.ParserModule.Parse("'ABC' <= 'ABC'")
        Assert.True(result7.Expression.GetValue(Nothing).Bool)

        Dim result8 = Analysis.ParserModule.Parse("00:01:30 <= 00:02:00")
        Assert.True(result8.Expression.GetValue(Nothing).Bool)

        Dim result9 = Analysis.ParserModule.Parse("2025-07-13 <= 2025-07-14")
        Assert.True(result9.Expression.GetValue(Nothing).Bool)

        Dim result10 = Analysis.ParserModule.Parse("true <= 2025-07-13")
        Assert.Throws(Of InvalidOperationException)(
            Function() result10.Expression.GetValue(Nothing)
        )
    End Sub

    <Fact>
    Public Sub TestLogicalOperators()
        ' 論理演算子の解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("true and false")
        Assert.False(result1.Expression.GetValue(Nothing).Bool)
        Dim result2 = Analysis.ParserModule.Parse("true or false")
        Assert.True(result2.Expression.GetValue(Nothing).Bool)
        Dim result3 = Analysis.ParserModule.Parse("not true")
        Assert.False(result3.Expression.GetValue(Nothing).Bool)
        Dim result4 = Analysis.ParserModule.Parse("true xor false")
        Assert.True(result4.Expression.GetValue(Nothing).Bool)
    End Sub

    <Fact>
    Public Sub TestTemaryOperators()
        ' 三項演算子の解析をテスト
        Dim result1 = Analysis.ParserModule.Parse("true ? 1 : 0")
        Assert.Equal(1, result1.Expression.GetValue(Nothing).Number)
        Dim result2 = Analysis.ParserModule.Parse("false ? 1 : 0")
        Assert.Equal(0, result2.Expression.GetValue(Nothing).Number)
        Dim result3 = Analysis.ParserModule.Parse("1 > 0 ? ""Yes"" : ""No""")
        Assert.True(result3.Expression.GetValue(Nothing).Str.Equals("Yes"))
        Dim result4 = Analysis.ParserModule.Parse("0 > 1 ? ""Yes"" : ""No""")
        Assert.True(result4.Expression.GetValue(Nothing).Str.Equals("No"))
    End Sub

    <Fact>
    Public Sub TestInvalidSyntax()
        ' 無効な構文の解析をテスト
        Assert.Throws(Of Analysis.AnalysisException)(
            Sub()
                Analysis.ParserModule.Parse("123.456 +")
            End Sub
        )
        Assert.Throws(Of Analysis.AnalysisException)(
            Sub()
                Analysis.ParserModule.Parse("true and")
            End Sub
        )
        Assert.Throws(Of Analysis.AnalysisException)(
            Sub()
                Analysis.ParserModule.Parse("1 ? 2")
            End Sub
        )
    End Sub

    <Fact>
    Public Sub TestArrayFactor()
        Dim result1 = Analysis.ParserModule.Parse("[100, 110, 120][2]")
        Assert.Equal(120.0, result1.Expression.GetValue(Nothing).Number)
    End Sub

    <Fact>
    Public Sub TestNullFactor()
        Dim result1 = Analysis.ParserModule.Parse("null")
        Assert.Null(result1.Expression.GetValue(Nothing).Obj)
    End Sub

    <Fact>
    Public Sub TestTimeSpanFactor()
        Dim result1 = Analysis.ParserModule.Parse("23:25:17")
        Assert.Equal(New TimeSpan(23, 25, 17), result1.Expression.GetValue(Nothing).ToTimeSpan())
        Assert.Equal(U8String.NewString("23:25:17"), result1.Expression.GetValue(Nothing).Str())

        Dim result2 = Analysis.ParserModule.Parse("01:30:45")
        Assert.Equal(New TimeSpan(1, 30, 45), result2.Expression.GetValue(Nothing).ToTimeSpan())
        Assert.Equal(U8String.NewString("01:30:45"), result2.Expression.GetValue(Nothing).Str())

        Dim result3 = Analysis.ParserModule.Parse("2025-07-12")
        Assert.Equal(New DateTime(2025, 7, 12), result3.Expression.GetValue(Nothing).ToDate())
        Assert.Equal(U8String.NewString("2025-07-12T00:00:00.000"), result3.Expression.GetValue(Nothing).Str())

        Dim result4 = Analysis.ParserModule.Parse("2025-07-12T14:30:00")
        Assert.Equal(New DateTime(2025, 7, 12, 14, 30, 0), result4.Expression.GetValue(Nothing).ToDate())

        Dim result5 = Analysis.ParserModule.Parse("2025-07-12T13:28:05Z")
        Assert.Equal(New DateTime(2025, 7, 12, 13, 28, 5, DateTimeKind.Utc), result5.Expression.GetValue(Nothing).ToDate())

        Dim result6 = Analysis.ParserModule.Parse("2025-07-12T13:28:05+09:00")
        Assert.Equal(New DateTime(2025, 7, 12, 13, 28, 5).AddHours(9), result6.Expression.GetValue(Nothing).ToDate())

        Dim result7 = Analysis.ParserModule.Parse("2025-07-12T13:28:05-05:00")
        Assert.Equal(New DateTime(2025, 7, 12, 13, 28, 5).AddHours(-5), result7.Expression.GetValue(Nothing).ToDate())

        Dim result8 = Analysis.ParserModule.Parse("2025-07-12T13:28:05.001")
        Assert.Equal(New DateTime(2025, 7, 12, 13, 28, 5, 1), result8.Expression.GetValue(Nothing).ToDate())
        Dim result9 = Analysis.ParserModule.Parse("2025-07-12T13:28:05.001")
    End Sub

    <Fact>
    Public Sub TestCalcError()
        Dim result1 = Analysis.ParserModule.Parse("'AAA' xor 'BBB'")
        Assert.Throws(Of InvalidOperationException)(
            Function() result1.Expression.GetValue(Nothing)
        )

        Dim result2 = Analysis.ParserModule.Parse("true or false")
        Assert.True(result2.Expression.GetValue(Nothing).Bool)


        Dim result3 = Analysis.ParserModule.Parse("'A' or true")
        Assert.Throws(Of InvalidOperationException)(
            Function() result3.Expression.GetValue(Nothing)
        )

        Dim result4 = Analysis.ParserModule.Parse("true and 'B'")
        Assert.Throws(Of InvalidOperationException)(
            Function() result4.Expression.GetValue(Nothing)
        )

        Dim result5 = Analysis.ParserModule.Parse("true and false")
        Assert.False(result5.Expression.GetValue(Nothing).Bool)

        Dim result6 = Analysis.ParserModule.Parse("'A' and false")
        Assert.Throws(Of InvalidOperationException)(
            Function() result6.Expression.GetValue(Nothing)
        )
    End Sub

End Class
