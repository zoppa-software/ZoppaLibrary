Option Explicit On
Option Strict On

Imports ZoppaLibrary.Parser
Imports Xunit

Public Class ExpressionRangeTests

    <Fact>
    Public Sub ToString_Invalid_ReturnsEmpty()
        Assert.Equal(String.Empty, ExpressionRange.Invalid.ToString())
    End Sub

End Class