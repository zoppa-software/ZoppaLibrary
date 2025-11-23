Imports System
Imports Xunit
Imports ZoppaLibrary
Imports ZoppaLibrary.Strings

Public Class StringCharTest

    ''' <summary>
    ''' U8Char の生成テスト。
    ''' </summary>
    <Fact>
    Public Sub CreateU8CharTest()
        ' 1バイト文字
        Dim c1 = U8Char.NewChar("A"c)
        Assert.Equal("A", c1.ToString())
        Assert.Equal(1, c1.Size)

        ' マルチバイト文字
        Dim c2 = U8Char.NewChar("あ"c)
        Assert.Equal("あ", c2.ToString())
        Assert.Equal(3, c2.Size)
    End Sub

    ''' <summary>
    ''' U8Char の空白文字判定テスト。
    ''' </summary>
    <Fact>
    Public Sub U8CharIsWhiteSpaceTest()
        ' 半角スペース
        Dim c1 = U8Char.NewChar(" "c)
        Assert.True(c1.IsWhiteSpace())

        ' タブ
        Dim c2 = U8Char.NewChar(ChrW(&H9))
        Assert.True(c2.IsWhiteSpace())

        ' 全角スペース
        Dim c3 = U8Char.NewChar("　"c)
        Assert.True(c3.IsWhiteSpace())

        ' 改行
        Dim c4 = U8Char.NewChar(ChrW(&HA))
        Assert.True(c4.IsWhiteSpace())

        ' キャリッジリターン
        Dim c5 = U8Char.NewChar(ChrW(&HD))
        Assert.True(c5.IsWhiteSpace())

        ' 空白でない文字
        Dim c6 = U8Char.NewChar("A"c)
        Assert.False(c6.IsWhiteSpace())
    End Sub

    <Fact>
    Public Sub NewString_ValidString_ReturnsCorrectU8String()
        Dim s As String = "あいう"
        Dim u8 = U8String.NewString(s)
        Assert.Equal(s, u8.ToString())
    End Sub

    <Fact>
    Public Sub NewSlice_ValidRange_ReturnsCorrectSlice()
        Dim s As String = "テスト文字列"
        Dim u8 = U8String.NewString(s)
        ' 先頭2文字（UTF-8: 3バイト×2 = 6バイト）
        Dim slice = U8String.NewSlice(u8, 0, 6)
        Assert.Equal(s.Substring(0, 2), slice.ToString())
    End Sub

    <Fact>
    Public Sub NewSlice_NullSource_ThrowsArgumentNullException()
        Dim dummy As U8String
        Assert.Throws(Of ArgumentNullException)(Function() U8String.NewSlice(dummy, 0, 1))
    End Sub

    <Fact>
    Public Sub ToString_Empty_ReturnsEmptyString()
        Dim u8 = U8String.NewString("")
        Assert.Equal("", u8.ToString())
    End Sub

    <Fact>
    Public Sub Iterator_WalksAllCharacters()
        Dim s As String = "abcあいう"
        Dim u8 = U8String.NewString(s)
        Dim it = u8.GetIterator()
        Dim result As String = ""
        While it.HasNext
            Dim c = it.MoveNext()
            If c.HasValue Then
                result &= c.Value.ToString()
            End If
        End While
        Assert.Equal(s, result)

        Assert.Null(it.Current)
    End Sub

    <Fact>
    Public Sub Iterator_MiddleString_ReturnsNoCharacters()
        Dim u8 = U8String.NewString("ゲゲゲの鬼太郎とネズミ男")
        Dim slice = U8String.NewSlice(u8, 12, 9)
        Dim it = slice.GetIterator()
        Assert.Equal(U8Char.NewChar("鬼"c), it.Peek(0))
        Assert.Equal(U8Char.NewChar("太"c), it.Peek(1))
        Assert.Equal(U8Char.NewChar("郎"c), it.Peek(2))
    End Sub

    <Fact>
    Public Sub Iterator_Peek_SkipsCorrectly()
        Dim s As String = "abc"
        Dim u8 = U8String.NewString(s)
        Dim it = u8.GetIterator()
        Dim peeked = it.Peek(2)
        Assert.Equal("c"c, peeked.Value.ToString())
    End Sub

    <Fact>
    Public Sub Empty_ReturnsEmptyU8String()
        Dim empty = U8String.Empty
        Assert.Equal("", empty.ToString())
    End Sub

    <Fact>
    Public Sub Concat_AsciiStrings_ReturnsConcatenated()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("123")
        Dim result = s1.Concat(s2)
        Assert.Equal("abc123", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_UnicodeStrings_ReturnsConcatenated()
        Dim s1 = U8String.NewString("あいう")
        Dim s2 = U8String.NewString("えお")
        Dim result = s1.Concat(s2)
        Assert.Equal("あいうえお", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_WithEmptyLeft_ReturnsRight()
        Dim s1 = U8String.Empty
        Dim s2 = U8String.NewString("abc")
        Dim result = s1.Concat(s2)
        Assert.Equal("abc", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_WithEmptyRight_ReturnsLeft()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.Empty
        Dim result = s1.Concat(s2)
        Assert.Equal("abc", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_BothEmpty_ReturnsEmpty()
        Dim s1 = U8String.Empty
        Dim s2 = U8String.Empty
        Dim result = s1.Concat(s2)
        Assert.Equal("", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_LeftIsNothing_ReturnsRight()
        Dim s1 As U8String = Nothing
        Dim s2 = U8String.NewString("abc")
        Dim result = s1.Concat(s2)
        Assert.Equal("abc", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_RightIsNothing_ReturnsLeft()
        Dim s1 = U8String.NewString("abc")
        Dim s2 As U8String = Nothing
        Dim result = s1.Concat(s2)
        Assert.Equal("abc", result.ToString())
    End Sub

    <Fact>
    Public Sub Equals_SameU8String_ReturnsTrue()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("abc")
        Assert.True(s1.Equals(s2))
    End Sub

    <Fact>
    Public Sub Equals_DifferentU8String_ReturnsFalse()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("def")
        Assert.False(s1.Equals(s2))
    End Sub

    <Fact>
    Public Sub Equals_SameString_ReturnsTrue()
        Dim s1 = U8String.NewString("あいう")
        Assert.True(s1.Equals("あいう"))
    End Sub

    <Fact>
    Public Sub Equals_DifferentString_ReturnsFalse()
        Dim s1 = U8String.NewString("あいう")
        Assert.False(s1.Equals("えお"))
    End Sub

    <Fact>
    Public Sub Equals_DifferentLength_ReturnsFalse()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("abcd")
        Assert.False(s1.Equals(s2))
        Assert.False(s1.Equals("abcd"))
        Assert.False(s1.Equals("acd"))
    End Sub

    <Fact>
    Public Sub Equals_EmptyU8String_ReturnsTrue()
        Dim s1 = U8String.Empty
        Dim s2 = U8String.NewString("")
        Assert.True(s1.Equals(s2))
        Assert.True(s1.Equals(""))
    End Sub

    <Fact>
    Public Sub Equals_EmptyAndNonEmpty_ReturnsFalse()
        Dim s1 = U8String.Empty
        Dim s2 = U8String.NewString("abc")
        Assert.False(s1.Equals(s2))
        Assert.False(s1.Equals("abc"))
    End Sub

    <Fact>
    Public Sub Equals_NullObject_ReturnsFalse()
        Dim s1 = U8String.NewString("abc")
        Assert.False(s1.Equals(Nothing))
    End Sub

    <Fact>
    Public Sub Equals_DifferentType_ReturnsFalse()
        Dim s1 = U8String.NewString("abc")
        Assert.False(s1.Equals(123))
    End Sub

    <Fact>
    Public Sub Mid_MultibyteString_ReturnsCorrectSubstring()
        Dim u8 = U8String.NewString("あいうえおabcd")
        Dim mid = u8.Mid(3, 4)
        Assert.Equal("えおab", mid.ToString())
    End Sub

    <Fact>
    Public Sub At_AsciiString_ReturnsCorrectChar()
        Dim u8 = U8String.NewString("abcde")
        Assert.Equal("a", u8.At(0).Value.ToString())
        Assert.Equal("c", u8.At(2).Value.ToString())
        Assert.Equal("e", u8.At(4).Value.ToString())
    End Sub

    <Fact>
    Public Sub At_MultibyteString_ReturnsCorrectChar()
        Dim u8 = U8String.NewString("あいうえお")
        Assert.Equal("あ", u8.At(0).Value.ToString())
        Assert.Equal("う", u8.At(2).Value.ToString())
        Assert.Equal("お", u8.At(4).Value.ToString())
    End Sub

    <Fact>
    Public Sub At_MixedString_ReturnsCorrectChar()
        Dim u8 = U8String.NewString("aいcえお")
        Assert.Equal("a", u8.At(0).Value.ToString())
        Assert.Equal("い", u8.At(1).Value.ToString())
        Assert.Equal("c", u8.At(2).Value.ToString())
        Assert.Equal("え", u8.At(3).Value.ToString())
        Assert.Equal("お", u8.At(4).Value.ToString())

        Assert.Null(u8.At(-1))
        Assert.Null(u8.At(5))
    End Sub

    <Fact>
    Public Sub CompareTo_EqualStrings_ReturnsZero()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("abc")
        Assert.Equal(0, s1.CompareTo(s2))
    End Sub

    <Fact>
    Public Sub CompareTo_LessThan_ReturnsNegative()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("abd")
        Assert.True(s1.CompareTo(s2) < 0)
    End Sub

    <Fact>
    Public Sub CompareTo_GreaterThan_ReturnsPositive()
        Dim s1 = U8String.NewString("abd")
        Dim s2 = U8String.NewString("abc")
        Assert.True(s1.CompareTo(s2) > 0)
    End Sub

    <Fact>
    Public Sub CompareTo_PrefixIsLessThanLongerString()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("abcd")
        Assert.True(s1.CompareTo(s2) < 0)
        Assert.True(s2.CompareTo(s1) > 0)
    End Sub

    <Fact>
    Public Sub CompareTo_EmptyVsNonEmpty()
        Dim s1 = U8String.NewString("")
        Dim s2 = U8String.NewString("a")
        Assert.True(s1.CompareTo(s2) < 0)
        Assert.True(s2.CompareTo(s1) > 0)
        Assert.Equal(0, s1.CompareTo(U8String.NewString("")))
    End Sub

    <Fact>
    Public Sub CompareTo_MultibyteStrings()
        Dim s1 = U8String.NewString("あいう")
        Dim s2 = U8String.NewString("あいえ")
        Assert.True(s1.CompareTo(s2) < 0)
        Assert.True(s2.CompareTo(s1) > 0)
        Assert.Equal(0, s1.CompareTo(U8String.NewString("あいう")))
    End Sub

    <Fact>
    Public Sub CompareTo_MixedAsciiAndMultibyte()
        Dim s1 = U8String.NewString("abcあ")
        Dim s2 = U8String.NewString("abcい")
        Assert.True(s1.CompareTo(s2) < 0)
        Assert.True(s2.CompareTo(s1) > 0)
    End Sub

    <Fact>
    Public Sub CompareTo_BothEmpty_ReturnsZero()
        Dim s1 = U8String.NewString("")
        Dim s2 = U8String.NewString("")
        Assert.Equal(0, s1.CompareTo(s2))
    End Sub

    <Fact>
    Public Sub StartWith_U8String_AsciiAndUnicode()
        Dim s = U8String.NewString("abcdef")
        Assert.True(s.StartWith(U8String.NewString("abc")))
        Assert.False(s.StartWith(U8String.NewString("abd")))
        Assert.True(s.StartWith(U8String.NewString("abcdef")))
        Assert.False(s.StartWith(U8String.NewString("abcdefg")))
        Assert.False(s.StartWith(U8String.NewString(""))) ' 空文字列はFalse

        Dim s2 = U8String.NewString("あいうえお")
        Assert.True(s2.StartWith(U8String.NewString("あい")))
        Assert.False(s2.StartWith(U8String.NewString("いう")))
        Assert.True(s2.StartWith(U8String.NewString("あいうえお")))
        Assert.False(s2.StartWith(U8String.NewString("あいうえおか")))
    End Sub

    <Fact>
    Public Sub StartWith_String_AsciiAndUnicode()
        Dim s = U8String.NewString("abcdef")
        Assert.True(s.StartWith("abc"))
        Assert.False(s.StartWith("abd"))
        Assert.True(s.StartWith("abcdef"))
        Assert.False(s.StartWith("abcdefg"))
        Assert.False(s.StartWith(""))

        Dim s2 = U8String.NewString("あいうえお")
        Assert.True(s2.StartWith("あい"))
        Assert.False(s2.StartWith("いう"))
        Assert.True(s2.StartWith("あいうえお"))
        Assert.False(s2.StartWith("あいうえおか"))
    End Sub

    <Fact>
    Public Sub StartWith_NullOrEmpty()
        Dim s = U8String.NewString("abc")
        Dim empty = U8String.Empty
        Assert.False(s.StartWith(U8String.Empty))
        Assert.False(empty.StartWith(U8String.NewString("a")))
        Assert.False(empty.StartWith(U8String.Empty))
        Assert.False(s.StartWith(""))
        Assert.False(empty.StartWith("a"))
        Assert.False(empty.StartWith(""))
    End Sub

    <Fact>
    Public Sub Operator_Equals_ReturnsTrueForSameContent()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("abc")
        Assert.True(s1 = s2)
        Assert.False(s1 <> s2)
    End Sub

    <Fact>
    Public Sub Operator_Equals_ReturnsFalseForDifferentContent()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("abd")
        Assert.False(s1 = s2)
        Assert.True(s1 <> s2)
    End Sub

    <Fact>
    Public Sub Operator_Equals_ReturnsFalseForDifferentLength()
        Dim s1 = U8String.NewString("abc")
        Dim s2 = U8String.NewString("ab")
        Assert.False(s1 = s2)
        Assert.True(s1 <> s2)
    End Sub

    <Fact>
    Public Sub Operator_Equals_EmptyStrings()
        Dim s1 = U8String.Empty
        Dim s2 = U8String.NewString("")
        Assert.True(s1 = s2)
        Assert.False(s1 <> s2)
    End Sub

    <Fact>
    Public Sub Operator_Equals_Unicode()
        Dim s1 = U8String.NewString("あいう")
        Dim s2 = U8String.NewString("あいう")
        Dim s3 = U8String.NewString("あいえ")
        Assert.True(s1 = s2)
        Assert.False(s1 <> s2)
        Assert.False(s1 = s3)
        Assert.True(s1 <> s3)
    End Sub

    <Fact>
    Public Sub Operator_Equals_NullLike()
        Dim s1 As U8String = Nothing
        Dim s2 As U8String = Nothing
        Assert.True(s1 = s2)
        Assert.False(s1 <> s2)
        Dim s3 = U8String.NewString("abc")
        Assert.False(s1 = s3)
        Assert.True(s1 <> s3)
    End Sub

    <Fact>
    Public Sub Mid_OutOfRange_ThrowsArgumentOutOfRangeException()
        Dim u8 = U8String.NewString("abcde")

        ' first < 0
        Assert.Throws(Of ArgumentOutOfRangeException)(Function() u8.Mid(-1, 2))

        ' length < 0
        Assert.Throws(Of ArgumentOutOfRangeException)(Function() u8.Mid(1, -1))

        ' first + length > Length
        Assert.Throws(Of ArgumentOutOfRangeException)(Function() u8.Mid(3, 3))

        ' first > Length
        Assert.Throws(Of ArgumentOutOfRangeException)(Function() u8.Mid(6, 1))

        ' length > Length
        Assert.Throws(Of ArgumentOutOfRangeException)(Function() u8.Mid(0, 6))

        ' 空文字列で範囲外
        Dim empty = U8String.NewString("")
        Assert.Throws(Of ArgumentOutOfRangeException)(Function() empty.Mid(0, 1))
        Assert.Throws(Of ArgumentOutOfRangeException)(Function() empty.Mid(1, 0))
        Assert.Throws(Of ArgumentOutOfRangeException)(Function() empty.Mid(-1, 1))
    End Sub

    <Fact>
    Public Sub Concat_String_Ascii()
        Dim s1 = U8String.NewString("abc")
        Dim result = s1.Concat("def")
        Assert.Equal("abcdef", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_String_Unicode()
        Dim s1 = U8String.NewString("あいう")
        Dim result = s1.Concat("えお")
        Assert.Equal("あいうえお", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_String_LeftEmpty()
        Dim s1 = U8String.Empty
        Dim result = s1.Concat("xyz")
        Assert.Equal("xyz", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_String_RightEmpty()
        Dim s1 = U8String.NewString("xyz")
        Dim result = s1.Concat("")
        Assert.Equal("xyz", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_String_BothEmpty()
        Dim s1 = U8String.Empty
        Dim result = s1.Concat("")
        Assert.Equal("", result.ToString())
    End Sub

    <Fact>
    Public Sub Concat_String_Null_Throws()
        Dim s1 = U8String.NewString("abc")
        Assert.Throws(Of ArgumentNullException)(Function() s1.Concat(CType(Nothing, String)))
    End Sub

End Class
