Option Explicit On
Option Strict On

Imports Xunit
Imports ZoppaLibrary.EBNF

Namespace Analysis

    Public Class TerminalAnalysisTest

        <Fact>
        Public Sub UnescapedString_WithNull_ReturnsEmpty()
            ' Arrange
            Dim input As String = Nothing

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal(String.Empty, result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithEmptyString_ReturnsEmpty()
            ' Arrange
            Dim input = String.Empty

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal(String.Empty, result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithNoEscapeSequences_ReturnsSameString()
            ' Arrange
            Dim input = "Hello World"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Hello World", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithNewlineEscape_ReturnsLineFeed()
            ' Arrange
            Dim input = "Hello\nWorld"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Hello" & vbLf & "World", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithCarriageReturnEscape_ReturnsCarriageReturn()
            ' Arrange
            Dim input = "Hello\rWorld"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Hello" & vbCr & "World", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithTabEscape_ReturnsTab()
            ' Arrange
            Dim input = "Hello\tWorld"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Hello" & vbTab & "World", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithBackslashEscape_ReturnsBackslash()
            ' Arrange
            Dim input = "Hello\\World"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Hello\World", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithMultipleEscapes_ReturnsCorrectString()
            ' Arrange
            Dim input = "Line1\nLine2\tTabbed\rReturn"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Line1" & vbLf & "Line2" & vbTab & "Tabbed" & vbCr & "Return", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithUnknownEscape_ReturnsCharacterAsIs()
            ' Arrange
            Dim input = "Hello\xWorld"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("HelloxWorld", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithBackslashAtEnd_ReturnsStringWithoutChange()
            ' Arrange
            Dim input = "Hello\"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Hello\", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithMultipleBackslashes_HandlesCorrectly()
            ' Arrange
            Dim input = "Path\\To\\File"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Path\To\File", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithAllSupportedEscapes_ReturnsCorrectString()
            ' Arrange
            Dim input = "\n\r\t\\"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal(vbLf & vbCr & vbTab & "\", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithMixedContent_ReturnsCorrectString()
            ' Arrange
            Dim input = "Start\tMiddle\nEnd\\Done"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Start" & vbTab & "Middle" & vbLf & "End\Done", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithConsecutiveEscapes_HandlesCorrectly()
            ' Arrange
            Dim input = "Test\n\n\t\tValue"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Test" & vbLf & vbLf & vbTab & vbTab & "Value", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithEscapeAtStart_ReturnsCorrectString()
            ' Arrange
            Dim input = "\nStartWithNewline"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal(vbLf & "StartWithNewline", result)
        End Sub

        <Fact>
        Public Sub UnescapedString_WithSpecialCharactersAndEscapes_ReturnsCorrectString()
            ' Arrange
            Dim input = "Quote""Test\nWith\\Backslash"

            ' Act
            Dim result = InvokeUnescapedString(input)

            ' Assert
            Assert.Equal("Quote""Test" & vbLf & "With\Backslash", result)
        End Sub

        ''' <summary>
        ''' リフレクションを使用してプライベートメソッド UnescapedString を呼び出します。
        ''' </summary>
        Private Function InvokeUnescapedString(input As String) As String
            Dim methodInfo = GetType(TerminalNode).GetMethod("UnescapedString",
                System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Static)
            Return CStr(methodInfo.Invoke(Nothing, New Object() {input}))
        End Function

    End Class

End Namespace