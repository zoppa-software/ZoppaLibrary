Imports System
Imports Xunit
Imports ZoppaLibrary.Strings
Imports ZoppaLibrary.Toml

Public Class InvalidTest

    <Fact>
    Public Sub InvalidTomlTest()
        Dim src1 = U8String.NewString("[t1]
t2.t3.v = 0
[t1.t2]")
        Assert.Throws(Of TomlTableDuplicationException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

End Class
