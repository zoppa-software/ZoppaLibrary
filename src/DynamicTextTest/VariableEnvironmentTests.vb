Imports Xunit
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.Strings

Public Class VariableEnvironmentTests

    <Fact>
    Public Sub RegistNumber_And_Get_ReturnsCorrectValue()
        Dim env As New AnalysisEnvironment()
        env.RegisterNumber("num", 42.0)
        Dim v = env.Get("num")
        Assert.Equal(VariableType.Number, v.Type)
        Assert.Equal(42.0, v.Number)
    End Sub

    <Fact>
    Public Sub RegistStr_And_Get_ReturnsCorrectValue()
        Dim env As New AnalysisEnvironment()
        env.RegisterStr("str", "hello")
        Dim v = env.Get("str")
        Assert.Equal(VariableType.Str, v.Type)
        Assert.Equal(U8String.NewString("hello"), v.Str(env))
    End Sub

    <Fact>
    Public Sub RegistBool_And_Get_ReturnsCorrectValue()
        Dim env As New AnalysisEnvironment()
        env.RegisterBool("flag", True)
        Dim v = env.Get("flag")
        Assert.Equal(VariableType.Bool, v.Type)
        Assert.True(v.Bool)
    End Sub

    <Fact>
    Public Sub Unregist_RemovesVariable()
        Dim env As New AnalysisEnvironment()
        env.RegisterNumber("x", 1)
        env.Unregister("x")
        Assert.Throws(Of KeyNotFoundException)(Sub() env.Get("x"))
    End Sub

    <Fact>
    Public Sub Hierarchy_ScopeTest()
        Dim env As New AnalysisEnvironment()
        env.RegisterNumber("x", 1)
        Using env.GetScope()
            env.RegisterNumber("x", 2)
            Assert.Equal(2, env.Get("x").Number)
        End Using
        Assert.Equal(1, env.Get("x").Number)
    End Sub

    <Fact>
    Public Sub Get_ThrowsIfNotFound()
        Dim env As New AnalysisEnvironment()
        Assert.Throws(Of KeyNotFoundException)(Sub() env.Get("notfound"))
    End Sub

    <Fact>
    Public Sub TimeSpanVariableTest()
        Dim env As New AnalysisEnvironment()

        env.RegisterTimeSpan("duration", New TimeSpan(1, 2, 3))
        Dim v = env.Get("duration")
        Assert.Equal(VariableType.Time, v.Type)
        Assert.Equal(New TimeSpan(1, 2, 3), v.ToTime())
        env.Unregister("duration")

        env.RegisterTimeSpan("duration", TimeSpan.FromMinutes(90))
        v = env.Get("duration")
        Assert.Equal(VariableType.Time, v.Type)
        Assert.Equal(TimeSpan.FromMinutes(90), v.ToTime())
        env.Unregister("duration")
    End Sub

    <Fact>
    Public Sub DateTimeVariableTest()
        Dim env As New AnalysisEnvironment()
        env.RegisterDateTime("date", New DateTime(2023, 10, 1, 12, 0, 0))
        Dim v = env.Get("date")
        Assert.Equal(VariableType.Date, v.Type)
        Assert.Equal(New DateTime(2023, 10, 1, 12, 0, 0), v.ToDate())
        env.Unregister("date")

        env.RegisterDateTime("date", DateTime.Now)
        v = env.Get("date")
        Assert.Equal(VariableType.Date, v.Type)
        Assert.True(v.ToDate() <= DateTime.Now)
    End Sub

End Class