Imports System
Imports Xunit
Imports ZoppaLibrary
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.Strings
Imports ZoppaLibrary.Switches

Public Class AnalysisSwitchTest

    <Fact>
    Public Sub AnalysisSwitchTest_1()
        Dim d1 = New SubCommandDefine("run", "アプリケーションを実行します")
        Dim s1 = d1.ToString()
        Dim d2 = New SwitchDefine("verbose", False, SwitchType.DoubleHyphen, ParameterType.Str, "Enable verbose output")
        Dim s2 = d2.ToString()

        Dim analysisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            "Sample",
            "1.0.0.0",
            "このアプリケーションはサンプルです",
            "© 2025",
            "Sample Inc.",
            "MIT License",
            False,
            New SubCommandDefine() {
                New SubCommandDefine("run", "アプリケーションを実行します"),
                New SubCommandDefine("list", "アプリケーションのリストを表示します")
            },
            New SwitchDefine() {
                New SwitchDefine("verbose", False, SwitchType.DoubleHyphen, ParameterType.Str, "Enable verbose output"),
                New SwitchDefine("config", False, SwitchType.DoubleHyphen, ParameterType.Int, "Path to configuration file")
            },
            ParameterType.Str
        )

        Dim sw = analysisSwitch.Parse(New String() {"list", "--verbose", "true", "--config", "10", "config.json"})

        sw.CheckRequired()

        Assert.NotNull(sw)
        Assert.Equal("list", sw.SubCommand.Name)
        Assert.Equal(2, sw.SwitchOptions.Length)
        Assert.Equal(sw.SwitchOptions(0).sw.Name, "verbose")
        Assert.Equal(sw.SwitchOptions(1).sw.Name, "config")
        Assert.Equal("true", sw.SwitchOptions(0).prm(0))
        Assert.Equal("10", sw.SwitchOptions(1).prm(0))
        Assert.Equal("config.json", sw.Parameters(0))
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_2()
        Assert.Throws(Of ArgumentException)(
            Sub()
                Dim analysisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
                    "Sample",
                    "1.0.0.0",
                    "このアプリケーションはサンプルです",
                    "© 2025",
                    "Sample Inc.",
                    "MIT License",
                    False,
                    New SubCommandDefine() {
                        New SubCommandDefine("run", "アプリケーションを実行します"),
                        New SubCommandDefine("list", "アプリケーションのリストを表示します")
                    },
                    New SwitchDefine() {
                        New SwitchDefine("ve", False, SwitchType.SingleHyphen, ParameterType.Str, "Enable verbose output")
                    },
                    ParameterType.Str
                )
            End Sub
        )
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_3()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appName:="ZoppaDynamicText",
            appVersion:="1.0.0.0",
            appDescription:="埋込式を解析します。",
            appCopyright:="© 2025",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("template", True, SwitchType.DoubleHyphen, ParameterType.URI, "埋込式を記述したテンプレートファイルパス。"),
                New SwitchDefine("param", True, SwitchType.DoubleHyphen, ParameterType.URI, "パラメータファイルパス。")
            },
            ParameterType.Str
        )

        Dim sw = anaisSwitch.Parse(New String() {"--help"})
        If sw.ContainsOption("help") Then
            Dim msg = anaisSwitch.GetHelp()
            Assert.Equal("名前
    ZoppaDynamicText
    version 1.0.0.0 © 2025 Sample Inc. MIT License
説明
    埋込式を解析します。
文法
    ZoppaDynamicText [--help] [--template <URI>] [--param <URI>]  <文字列>
オプション
    help : ヘルプを表示します。
    template : 埋込式を記述したテンプレートファイルパス。
    param : パラメータファイルパス。
", msg)
        Else
            Assert.Fail()
        End If
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_4()
        Dim analysisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            "Sample",
            "1.0.0.0",
            "このアプリケーションはサンプルです",
            "© 2025",
            "Sample Inc.",
            "MIT License",
            True,
            New SubCommandDefine() {
                New SubCommandDefine("run", "アプリケーションを実行します"),
                New SubCommandDefine("list", "アプリケーションのリストを表示します")
            },
            New SwitchDefine() {
                New SwitchDefine("verbose", False, SwitchType.DoubleHyphen, ParameterType.Str, "Enable verbose output"),
                New SwitchDefine("config", False, SwitchType.DoubleHyphen, ParameterType.Int, "Path to configuration file")
            },
            ParameterType.Str
        )

        Dim sw = analysisSwitch.Parse(New String() {"--verbose", "true", "--config", "10", "config.json"})

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw.CheckRequired()
            End Sub
        )
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_5()
        Dim analysisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            "Sample",
            "1.0.0.0",
            "このアプリケーションはサンプルです",
            "© 2025",
            "Sample Inc.",
            "MIT License",
            True,
            New SubCommandDefine() {
                New SubCommandDefine("run", "アプリケーションを実行します"),
                New SubCommandDefine("list", "アプリケーションのリストを表示します")
            },
            New SwitchDefine() {
                New SwitchDefine("verbose", True, SwitchType.DoubleHyphen, ParameterType.Str, "Enable verbose output"),
                New SwitchDefine("config", False, SwitchType.DoubleHyphen, ParameterType.Int, "Path to configuration file")
            },
            ParameterType.Str
        )

        Dim sw = analysisSwitch.Parse(New String() {"run", "--config", "10", "config.json"})

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw.CheckRequired()
            End Sub
        )
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_7()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("template", True, SwitchType.DoubleHyphen, ParameterType.URI, "埋込式を記述したテンプレートファイルパス。"),
                New SwitchDefine("param", True, SwitchType.DoubleHyphen, ParameterType.URI, "パラメータファイルパス。")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"--template", "resources\template.txt", "--param", "resources\param.txt"})
        Dim tempPath = sw.GetOption("template").GetURI()
        Dim paramPath = sw.GetOption("param").GetURI()
        Assert.Equal("resources\template.txt", tempPath.OriginalString)
        Assert.Equal("resources\param.txt", paramPath.OriginalString)
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_8()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("p", True, SwitchType.SingleHyphen, ParameterType.Int, "整数"),
                New SwitchDefine("n", True, SwitchType.SingleHyphen, ParameterType.None, "なし")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"-np", "10"})
        sw.CheckRequired()
        Dim sw_n = sw.GetOption("p").GetInteger()
        Dim sw_p = sw.ContainsOption("n")
        Assert.Equal(10, sw_n)
        Assert.True(sw_p)
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_9()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("p", True, SwitchType.SingleHyphen, ParameterType.Int, "整数"),
                New SwitchDefine("n", True, SwitchType.SingleHyphen, ParameterType.None, "なし"),
                New SwitchDefine("i", False, SwitchType.SingleHyphen, ParameterType.None, "なし")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"-np", "A"})

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw.CheckRequired()
            End Sub
        )

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw.GetOption("p").GetInteger()
            End Sub
        )

        Dim sw_p = sw.ContainsOption("n")
        Assert.True(sw_p)

        Dim sw_i = sw.ContainsOption("i")
        Assert.False(sw_i)
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_10()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("p", True, SwitchType.SingleHyphen, ParameterType.Dbl, "実数"),
                New SwitchDefine("n", True, SwitchType.SingleHyphen, ParameterType.None, "なし")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"-np", "10.1"})
        sw.CheckRequired()
        Dim sw_n = sw.GetOption("p").GetDouble()
        Dim sw_p = sw.ContainsOption("n")
        Assert.Equal(10.1, sw_n)
        Assert.True(sw_p)
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_11()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("p", True, SwitchType.SingleHyphen, ParameterType.Dbl, "実数"),
                New SwitchDefine("n", True, SwitchType.SingleHyphen, ParameterType.None, "なし"),
                New SwitchDefine("i", False, SwitchType.SingleHyphen, ParameterType.None, "なし")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"-np", "A"})

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw.CheckRequired()
            End Sub
        )

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw.GetOption("p").GetDouble()
            End Sub
        )

        Dim sw_p = sw.ContainsOption("n")
        Assert.True(sw_p)

        Dim sw_i = sw.ContainsOption("i")
        Assert.False(sw_i)
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_12()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("p", True, SwitchType.SingleHyphen, ParameterType.URI, "Uri"),
                New SwitchDefine("n", True, SwitchType.SingleHyphen, ParameterType.None, "なし"),
                New SwitchDefine("i", False, SwitchType.SingleHyphen, ParameterType.None, "なし")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"-np", "\\\\\\"})

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw.CheckRequired()
            End Sub
        )

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw.GetOption("p").GetDouble()
            End Sub
        )

        Dim sw_p = sw.ContainsOption("n")
        Assert.True(sw_p)

        Dim sw_i = sw.ContainsOption("i")
        Assert.False(sw_i)
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_13()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("p", True, SwitchType.SingleHyphen, ParameterType.Str, "文字列"),
                New SwitchDefine("n", True, SwitchType.SingleHyphen, ParameterType.None, "なし")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"-np", "ABC"})
        sw.CheckRequired()
        Dim sw_n = sw.GetOption("p").GetStr()
        Dim sw_p = sw.ContainsOption("n")
        Assert.Equal("ABC", sw_n)
        Assert.True(sw_p)

        Dim sw2 = anaisSwitch.Parse(New String() {"-np"})

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw2.CheckRequired()
            End Sub
        )

        Assert.Throws(Of ArgumentException)(
            Sub()
                sw2.GetOption("p").GetStr()
            End Sub
        )
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_20()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("p", True, SwitchType.Slash, ParameterType.Str Or ParameterType.Array, "文字列"),
                New SwitchDefine("n", True, SwitchType.Slash, ParameterType.None, "なし")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"/p", "ABC", "EFG", "/n"})
        sw.CheckRequired()
        Dim sw_n = sw.GetOption("p").GetStrArray()
        Dim sw_p = sw.ContainsOption("n")
        Assert.Equal("ABC", sw_n(0))
        Assert.Equal("EFG", sw_n(1))
        Assert.True(sw_p)
    End Sub

    <Fact>
    Public Sub AnalysisSwitchTest_21()
        Dim anaisSwitch As AnalysisSwitch = AnalysisSwitch.Create(
            appDescription:="埋込式を解析します。",
            appAuthor:="Sample Inc.",
            appLicense:="MIT License",
            subCommandRequired:=False,
            subCommand:=Array.Empty(Of SubCommandDefine)(),
            options:=New SwitchDefine() {
                New SwitchDefine("help", False, SwitchType.DoubleHyphen, ParameterType.None, "ヘルプを表示します。"),
                New SwitchDefine("p", True, SwitchType.Slash, ParameterType.Int Or ParameterType.Array, "整数"),
                New SwitchDefine("n", True, SwitchType.Slash, ParameterType.None, "なし")
            }
        )

        Dim sw = anaisSwitch.Parse(New String() {"/n", "/p", "1", "2"})
        sw.CheckRequired()
        Dim sw_n = sw.GetOption("p").GetIntegerArray()
        Dim sw_p = sw.ContainsOption("n")
        Assert.Equal(1, sw_n(0))
        Assert.Equal(2, sw_n(1))
        Assert.True(sw_p)
    End Sub

End Class
