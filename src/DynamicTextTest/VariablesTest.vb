Imports Xunit
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.Strings

' テスト用のダミーIVariable実装
Friend Class DummyVariable
    Implements IVariable

    Public Property Data As U8String

    Public ReadOnly Property Type As VariableType Implements IVariable.Type
        Get
            Return VariableType.Str
        End Get
    End Property

    Public Sub New(data As U8String)
        Me.Data = data
    End Sub
End Class

Public Class VariablesTest

    <Fact>
    Public Sub Regist_NewVariable_AddsEntry()
        Dim env As New AnalysisEnvironment()

        Dim vars As New Variables()
        Dim v = U8String.NewString("foo")
        vars.Register("x", New DummyVariable(v))
        Assert.Equal(v, CType(vars.Get("x"), DummyVariable).Data)

        ' 新しい変数を登録
        Dim v1 = U8String.NewString("bar")
        vars.Register("x", New DummyVariable(v1))
        Assert.Equal(v1, CType(vars.Get("x"), DummyVariable).Data)

        ' もう一度同じ値で登録してもエラーにならない
        vars.Register("x", New DummyVariable(v1))
        Assert.Equal(v1, CType(vars.Get("x"), DummyVariable).Data)

        ' 変数を登録解除
        vars.Unregister("x")
        Assert.Throws(Of KeyNotFoundException)(Function() vars.Get("x"))
    End Sub

    <Fact>
    Public Sub RegistExprTest()
        Dim env As New AnalysisEnvironment()
        env.RegisterExpr("define", "${a=10}")
        Dim result = ParserModule.Translate("#{define}変数a = '#{a}'")
        Assert.True(result.Expression.GetValue(env).Str.Equals("変数a = '10'"))
    End Sub


End Class