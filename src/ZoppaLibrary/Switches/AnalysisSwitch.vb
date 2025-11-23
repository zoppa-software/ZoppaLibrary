Option Strict On
Option Explicit On

Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.Strings

Namespace Switches

    ''' <summary>コマンドライン解析のためのスイッチを定義します。</summary>
    ''' <remarks>
    ''' このクラスは、アプリケーションの基本情報とサブコマンド、オプションの定義を提供します。
    ''' </remarks>
    Public NotInheritable Class AnalysisSwitch

#Region "properties"

        ''' <summary>アプリケーション名。</summary>
        ''' <remarks>このプロパティは、アプリケーションの名前を表します。</remarks>
        Public ReadOnly Property AppName As String

        ''' <summary>アプリケーションのバージョン。</summary>
        ''' <remarks>このプロパティは、アプリケーションのバージョンを表します。</remarks>
        Public ReadOnly Property AppVersion As String

        ''' <summary>アプリケーションの説明。</summary>
        ''' <remarks>このプロパティは、アプリケーションの簡単な説明を表します。</remarks>
        Public ReadOnly Property AppDescription As String

        ''' <summary>アプリケーションの著作権情報。</summary>
        ''' <remarks>このプロパティは、アプリケーションの著作権情報を表します。</remarks>
        Public ReadOnly Property AppCopyright As String

        ''' <summary>アプリケーションの制作者情報。</summary>
        ''' <remarks>このプロパティは、アプリケーションの制作者情報を表します。</remarks>
        Public ReadOnly Property AppAuthor As String

        ''' <summary>アプリケーションのライセンス情報。</summary>
        ''' <remarks>このプロパティは、アプリケーションのライセンス情報を表します。</remarks>
        Public ReadOnly Property AppLicense As String

        ''' <summary>サブコマンドの定義を表します。</summary>
        ''' <remarks>
        ''' このプロパティは、サブコマンドの必要性とその定義を含むSubCommandSwitchオブジェクトを返します。
        ''' </remarks>
        Public ReadOnly Property SubCommands As SubCommandSwitch

        ''' <summary>サブコマンドを使用するかどうかを示します。</summary>
        Public ReadOnly Property UseSubCommand As Boolean
            Get
                Return Me.SubCommands IsNot Nothing AndAlso Me.SubCommands.Commands.Length > 0
            End Get
        End Property

        ''' <summary>オプションの定義を表します。</summary>
        ''' <remarks>
        ''' このプロパティは、コマンドラインオプションの定義を含むSwitchDefineのリストを返します。
        ''' </remarks>
        Public ReadOnly Property SwitchOptions As SwitchDefine()

        ''' <summary>パラメータの型を表します。</summary>
        ''' <remarks>
        ''' このプロパティは、パラメータの型を示すParameterType列挙体の値を返します。
        ''' </remarks>
        Public ReadOnly Property ParameterType As ParameterType

#End Region

        ''' <summary>
        ''' コンストラクタ。
        ''' <para>アプリケーションの基本情報とサブコマンド、オプションの定義を指定して、AnalysisSwitchオブジェクトを初期化します。</para>
        ''' </summary>
        ''' <param name="appName">アプリケーション名。</param>
        ''' <param name="appVersion">アプリケーションのバージョン。</param>
        ''' <param name="appDescription">アプリケーションの説明。</param>
        ''' <param name="appCopyright">アプリケーションの著作権情報。</param>
        ''' <param name="appAuthor">アプリケーションの制作者情報。</param>
        ''' <param name="appLicense">アプリケーションのライセンス情報。</param>
        ''' <param name="subCommandRequired">サブコマンドが必要かどうか。</param>
        ''' <param name="subCommand">サブコマンドの定義の配列。</param>
        ''' <param name="options">コマンドラインオプションの定義の配列。</param>
        ''' <param name="paramType">パラメータタイプ。</param>
        Private Sub New(appName As String,
                        appVersion As String,
                        appDescription As String,
                        appCopyright As String,
                        appAuthor As String,
                        appLicense As String,
                        subCommandRequired As Boolean,
                        subCommand() As SubCommandDefine,
                        options() As SwitchDefine,
                        paramType As ParameterType)
            Me.AppName = appName
            Me.AppVersion = appVersion
            Me.AppDescription = appDescription
            Me.AppCopyright = appCopyright
            Me.AppAuthor = appAuthor
            Me.AppLicense = appLicense
            Me.SubCommands = New SubCommandSwitch(subCommandRequired, subCommand)
            If options IsNot Nothing Then
                Me.SwitchOptions = options
            End If
            Me.ParameterType = paramType
        End Sub

        ''' <summary>
        ''' コマンドライン解析オブジェクトを作成します。
        ''' </summary>
        ''' <param name="appName">アプリケーション名。</param>
        ''' <param name="appVersion">アプリケーションのバージョン。</param>
        ''' <param name="appDescription">アプリケーションの説明。</param>
        ''' <param name="appCopyright">アプリケーションの著作権情報。</param>
        ''' <param name="appAuthor">アプリケーションの制作者情報。</param>
        ''' <param name="appLicense">アプリケーションのライセンス情報。</param>
        ''' <param name="subCommandRequired">サブコマンドが必要かどうか。</param>
        ''' <param name="subCommand">サブコマンドの定義の配列。</param>
        ''' <param name="options">コマンドラインオプションの定義の配列。</param>
        ''' <param name="paramType">パラメータタイプ。</param>
        ''' <returns>AnalysisSwitchオブジェクト。</returns>
        Public Shared Function Create(appName As String,
                                      appVersion As String,
                                      appDescription As String,
                                      appCopyright As String,
                                      appAuthor As String,
                                      appLicense As String,
                                      subCommandRequired As Boolean,
                                      subCommand() As SubCommandDefine,
                                      options() As SwitchDefine,
                                      paramType As ParameterType) As AnalysisSwitch
            Return New AnalysisSwitch(appName, appVersion, appDescription, appCopyright, appAuthor,
                                      appLicense, subCommandRequired, subCommand, options, paramType)
        End Function

        ''' <summary>
        ''' コマンドライン解析オブジェクトを作成します。
        ''' </summary>
        ''' <param name="appDescription">アプリケーションの説明。</param>
        ''' <param name="appAuthor">アプリケーションの制作者情報。</param>
        ''' <param name="appLicense">アプリケーションのライセンス情報。</param>
        ''' <param name="subCommandRequired">サブコマンドが必要かどうか。</param>
        ''' <param name="subCommand">サブコマンドの定義の配列。</param>
        ''' <param name="options">コマンドラインオプションの定義の配列。</param>
        ''' <returns>AnalysisSwitchオブジェクト。</returns>
        Public Shared Function Create(appDescription As String,
                                      appAuthor As String,
                                      appLicense As String,
                                      subCommandRequired As Boolean,
                                      subCommand() As SubCommandDefine,
                                      options() As SwitchDefine,
                                      Optional paramType As ParameterType = ParameterType.None) As AnalysisSwitch
            Dim appAsm = Assembly.GetEntryAssembly()
            Dim fileInfo = FileVersionInfo.GetVersionInfo(appAsm.Location)
            Return New AnalysisSwitch(appAsm.GetName().Name, appAsm.GetName().Version.ToString(), appDescription,
                                      fileInfo.LegalCopyright, appAuthor, appLicense, subCommandRequired, subCommand, options, paramType)
        End Function

        ''' <summary>
        ''' ヘルプメッセージを取得します。
        ''' </summary>
        ''' <returns>ヘルプメッセージの文字列。</returns>
        ''' <remarks>
        ''' このメソッドは、アプリケーションの基本情報、サブコマンド、オプションの定義を含むヘルプメッセージを生成します。
        ''' </remarks>
        Public Function GetHelp() As String
            ' 解析環境を初期化し、オブジェクトを登録します。
            Dim env As New AnalysisEnvironment()
            env.RegisterReflectObject(Me)

            ' サブコマンドとオプションの展開式を登録します。
            env.RegisterExpr("subCommand", "{trim}{trim '|'}{for sw in SubCommands.Commands}#{sw.Name} | {/for}{/trim}{/trim}")
            env.RegisterExpr("options", "{select sw.SwType}{case 1}-{case 2}--{case 3}/{/select}#{sw.Name}#{convParamType(sw.ParamType)}")
            env.AddFunction(U8String.NewString("convParamType"), AddressOf GetParamType)

            ' ヘルプメッセージを定義します。
            Dim helpMessage As String = "名前
    #{AppName}
    version #{AppVersion} #{AppCopyright} #{AppAuthor} #{AppLicense}
説明
    #{AppDescription}
文法
    #{AppName} {if UseSubCommand}{if SubCommands.Required}\{#{subCommand}\}{else}[#{subCommand}]{/if} {/if}{if SwitchOptions.Length > 0}{for sw in SwitchOptions}{if SubCommands.Required}#{options}{else}[#{options}]{/if} {/for}{/if}#{convParamType(ParameterType)}
{if UseSubCommand}{trim}サブコマンド
{for sw in SubCommands.Commands}    #{sw.Name}:#{sw.Description}
{/for}{/trim}
{/if}{if SwitchOptions.Length > 0}{trim}オプション
{for sw in SwitchOptions}    #{sw.Name} : #{sw.Description}
{/for}{/trim}
{/if}"

            ' ヘルプメッセージを解析して結果を取得します。
            Dim result = ParserModule.Translate(helpMessage)
            Return result.Expression.GetValue(env).Str.ToString()
        End Function

        ''' <summary>
        ''' パラメータの型を取得します。
        ''' </summary>
        ''' <param name="prm">パラメータ値。</param>
        ''' <returns>パラメータの型を表す文字列。</returns>
        ''' <remarks>
        ''' このメソッドは、パラメータの型に応じて適切な文字列を返します。
        ''' </remarks>
        Private Function GetParamType(prm As IValue) As IValue
            Dim typeValue = CInt(prm.Number)
            Dim pType As String
            Select Case typeValue And Not ParameterType.Array
                Case ParameterType.None
                    pType = ""
                Case ParameterType.Str
                    pType = "文字列"
                Case ParameterType.Int
                    pType = "整数"
                Case ParameterType.Dbl
                    pType = "実数"
                Case ParameterType.URI
                    pType = "URI"
                Case Else
                    pType = "不明な型"
            End Select

            ' 引数を表示
            If pType <> "" Then
                If (typeValue And ParameterType.Array) <> 0 Then
                    Return $" <{pType}[]]>".ToStringValue()
                Else
                    Return $" <{pType}>".ToStringValue()
                End If
            Else
                Return StringValue.Empty
            End If
        End Function

        ''' <summary>
        ''' コマンドライン引数を解析します。
        ''' </summary>
        ''' <param name="args">コマンドライン引数の配列。</param>
        ''' <returns>解析結果を含むResultオブジェクト。</returns>
        ''' <remarks>
        ''' このメソッドは、サブコマンドとオプションを解析し、結果をResultオブジェクトとして返します。
        ''' </remarks>
        Public Function Parse(args As String()) As Result
            ' サブコマンドの解析
            Dim subcommand As SubCommandDefine = Nothing
            Dim index = 0
            If Me.UseSubCommand Then
                subcommand = Me.SubCommands.Commands.Where(Function(sc) sc.Name = args(index)).FirstOrDefault()
                If subcommand IsNot Nothing Then
                    index += 1
                End If
            End If

            ' オプションの解析
            Dim optList As New List(Of (sw As SwitchDefine, prm As String()))()
            Dim i As Integer = 0
            For i = index To args.Length - 1
                Dim arg = args(i)
                If arg.StartsWith("--") Then
                    ' オプションの開始
                    Dim opt = GetOption(arg.Substring(2), SwitchType.DoubleHyphen)
                    Dim prm = GetParameter(opt, args, i + 1)
                    optList.Add((opt, prm.ToArray()))
                    i += prm.Count ' パラメータの数だけインデックスを進める

                ElseIf arg.StartsWith("-") Then
                    ' ショートオプションの開始(-)
                    Dim opts = GetShortOption(arg.Substring(1))
                    Dim maxi As Integer
                    For Each opt In opts
                        Dim prm = GetParameter(opt, args, i + 1)
                        optList.Add((opt, prm.ToArray()))
                        maxi = Math.Max(maxi, i + prm.Count)
                    Next
                    i = maxi ' 最後のオプションのパラメータ数だけインデックスを進める

                ElseIf arg.StartsWith("/") Then
                    ' ショートオプションの開始(/)
                    Dim opt = GetOption(arg.Substring(1), SwitchType.Slash)
                    Dim prm = GetParameter(opt, args, i + 1)
                    optList.Add((opt, prm.ToArray()))
                    i += prm.Count ' パラメータの数だけインデックスを進める
                Else
                    ' パラメータとして扱う
                    Exit For
                End If
            Next

            ' パラメータの解析
            Dim parameters As New List(Of String)()
            If Me.ParameterType <> ParameterType.None AndAlso index < args.Length Then
                ' パラメータが必要な場合、残りの引数をパラメータとして追加
                For j = i To args.Length - 1
                    parameters.Add(args(j))
                Next
            Else
                ' パラメータが必要ない場合は空のリストを返す
                parameters = New List(Of String)()
            End If

            ' 結果オブジェクトを作成
            Return New Result(Me, subcommand, optList, parameters, Me.ParameterType)
        End Function

        ''' <summary>
        ''' 指定された名前とオプションタイプに基づいてオプションを取得します。
        ''' </summary>
        ''' <param name="name">オプションの名前。</param>
        ''' <param name="swType">オプションの種類。</param>
        ''' <returns>指定されたオプションのSwitchDefineオブジェクト。</returns>
        ''' <remarks>
        ''' このメソッドは、指定された名前とオプションタイプに一致するSwitchDefineオブジェクトを返します。
        ''' </remarks>
        Private Function GetOption(name As String, swType As SwitchType) As SwitchDefine
            Dim res = SwitchOptions.Where(Function(sw) sw.SwType = swType).Where(Function(op) op.Name = name).FirstOrDefault()
            If res Is Nothing Then
                ' オプションが見つからない場合は例外をスロー
                Throw New ArgumentException($"オプション '{name}' が見つかりません。")
            End If
            Return res
        End Function

        ''' <summary>
        ''' 指定された名前に一致するショートオプションを取得します。
        ''' </summary>
        ''' <param name="name">オプションの名前。</param>
        ''' <returns>ショートオプションのSwitchDefineの配列。</returns>
        ''' <remarks>
        ''' このメソッドは、指定された名前に一致するショートオプションを検索し、配列として返します。
        ''' </remarks>
        Private Function GetShortOption(name As String) As SwitchDefine()
            ' オプションの名前を取得
            Dim opts As New List(Of SwitchDefine)()
            For Each opt In SwitchOptions.Where(Function(sw) sw.SwType = SwitchType.SingleHyphen)
                ' オプションの名前が含まれている場合、リストに追加
                If name.ToCharArray().Any(Function(c) c = opt.Name) Then
                    opts.Add(opt)
                End If
            Next
            Return opts.ToArray()
        End Function

        ''' <summary>
        ''' 指定されたオプションとパラメータのインデックスに基づいてパラメータを取得します。
        ''' </summary>
        ''' <param name="opt">オプションの定義。</param>
        ''' <param name="param">コマンドライン引数の配列。</param>
        ''' <param name="index">パラメータの開始インデックス。</param>
        ''' <returns>指定されたオプションに関連するパラメータのリスト。</returns>
        ''' <remarks>
        ''' このメソッドは、オプションのパラメータを収集し、リストとして返します。
        ''' </remarks>
        Private Function GetParameter(opt As SwitchDefine, param() As String, index As Integer) As List(Of String)
            Dim res As New List(Of String)()

            ' パラメータなし
            If opt Is Nothing OrElse opt.ParamType = ParameterType.None Then
                Return res
            End If

            ' パラメータを収集
            If (opt.ParamType And ParameterType.Array) <> 0 Then
                For i As Integer = index To param.Length - 1
                    If param(i).StartsWith("--") OrElse param(i).StartsWith("-") OrElse param(i).StartsWith("/") Then
                        ' 次のオプションが来たら終了
                        Exit For
                    End If
                    res.Add(param(i))
                Next
            Else
                ' 単一のパラメータ
                If index < param.Length Then
                    res.Add(param(index))
                End If
            End If
            Return res
        End Function

        ''' <summary>
        ''' 整数値をチェックし、例外をスローします。
        ''' </summary>
        ''' <param name="prm">パラメータ値。</param>
        ''' <param name="msg">エラーメッセージ。</param>
        ''' <returns>整数値。</returns>
        ''' <remarks>
        ''' このメソッドは、指定されたパラメータが整数であることを確認し、そうでない場合は例外をスローします。
        ''' </remarks>
        Private Shared Function CheckIntegerThrowException(prm As String, msg As String) As Integer
            Dim res As Integer
            If Not Integer.TryParse(prm, res) Then
                Throw New ArgumentException(msg)
            End If
            Return res
        End Function

        ''' <summary>
        ''' 実数値をチェックし、例外をスローします。
        ''' </summary>
        ''' <param name="prm">パラメータ値。</param>
        ''' <param name="msg">エラーメッセージ。</param>
        ''' <returns>実数値。</returns>
        ''' <remarks>
        ''' このメソッドは、指定されたパラメータが実数であることを確認し、そうでない場合は例外をスローします。
        ''' </remarks>
        Private Shared Function CheckDoubleThrowException(prm As String, msg As String) As Double
            Dim res As Double
            If Not Double.TryParse(prm, res) Then
                Throw New ArgumentException(msg)
            End If
            Return res
        End Function

        ''' <summary>
        ''' URIをチェックし、例外をスローします。
        ''' </summary>
        ''' <param name="prm">パラメータ値。</param>
        ''' <param name="msg">エラーメッセージ。</param>
        ''' <returns>URIオブジェクト。</returns>
        ''' <remarks>
        ''' このメソッドは、指定されたパラメータがURI形式であることを確認し、そうでない場合は例外をスローします。
        ''' </remarks>
        Private Shared Function CheckUriThrowException(prm As String, msg As String) As Uri
            Try
                Return New Uri(prm, UriKind.RelativeOrAbsolute)
            Catch ex As UriFormatException
                Throw New ArgumentException(msg, ex)
            End Try
        End Function

        ''' <summary>
        ''' コマンドライン解析の結果を表すクラスです。
        ''' </summary>
        ''' <remarks>
        ''' このクラスは、サブコマンド、オプションリスト、およびパラメータを含む解析結果を提供します。
        ''' </remarks>
        Public NotInheritable Class Result

            ' コマンドライン解析
            Private ReadOnly _parent As AnalysisSwitch

            ''' <summary>
            ''' サブコマンドを取得します。
            ''' <para>サブコマンドが指定されていない場合は、Nothingを返します。</para>
            ''' </summary>
            Public ReadOnly Property SubCommand As SubCommandDefine

            ''' <summary>
            ''' オプションリストを取得します。
            ''' <para>各要素は、オプション定義とそのパラメータの配列を含むタプルです。</para>
            ''' </summary>
            Public ReadOnly Property SwitchOptions As (sw As SwitchDefine, prm As String())()

            ''' <summary>
            ''' パラメータのリストを取得します。
            ''' <para>コマンドライン引数の残りの部分をパラメータとして扱います。</para>
            ''' </summary>
            Public ReadOnly Property Parameters As String()

            ''' <summary>
            ''' パラメータの型を取得します。
            ''' </summary>
            ''' <remarks>
            ''' このプロパティは、パラメータの型を示すParameterType列挙体の値を返します。
            ''' </remarks>
            Public ReadOnly Property ParameterType As ParameterType

            ''' <summary>
            ''' コンストラクタ。
            ''' <para>サブコマンド、オプションリスト、およびパラメータを指定して、Resultオブジェクトを初期化します。</para>
            ''' </summary>
            ''' <param name="parent">コマンドライン解析。</param>
            ''' <param name="command">サブコマンドの定義。</param>
            ''' <param name="optList">オプションリストのタプルのリスト。</param>
            ''' <param name="parameters">パラメータのリスト。</param>
            ''' <remarks>
            ''' このコンストラクタは、解析結果を表すResultオブジェクトを初期化します。
            ''' </remarks>
            Public Sub New(parent As AnalysisSwitch, command As SubCommandDefine, optList As List(Of (sw As SwitchDefine, prm As String())), parameters As List(Of String), prmType As ParameterType)
                Me._parent = parent
                Me.SubCommand = command
                Me.SwitchOptions = optList.ToArray()
                Me.Parameters = parameters.ToArray()
                Me.ParameterType = prmType
            End Sub

            ''' <summary>
            ''' 指定されたオプションが存在するかどうかを確認します。
            ''' <para>オプションが存在する場合はTrueを返し、存在しない場合はFalseを返します。</para>
            ''' </summary>
            ''' <param name="swName">オプション名。</param>
            ''' <returns>オプションが存在していたら真。</returns>
            Public Function ContainsOption(swName As String) As Boolean
                Dim opt = Me.SwitchOptions.Where(Function(o) o.sw.Name = swName).Select(Function(t) t.sw).FirstOrDefault()
                Return opt IsNot Nothing
            End Function

            ''' <summary>
            ''' 必須なサブコマンドとオプションが指定されているかをチェックします。
            ''' <para>サブコマンドが必要で、指定されていない場合は例外をスローします。</para>
            ''' <para>必須のオプションが指定されていない場合も例外をスローします。</para>
            ''' </summary>
            Public Sub CheckRequired()
                ' サブコマンドが必要で、指定されていない場合は例外をスロー
                If Me._parent.SubCommands.Required AndAlso SubCommand Is Nothing Then
                    Throw New ArgumentException("サブコマンドが必要ですが、指定されていません。")
                End If

                ' 必須なオプションが指定されていない場合は例外をスロー
                Dim requiredOptions = Me._parent.SwitchOptions.Where(Function(sw) sw.Required)
                For Each opt In requiredOptions
                    If Not Me.SwitchOptions.Any(Function(o) o.sw.Name = opt.Name) Then
                        Throw New ArgumentException($"必須のオプション '{opt.Name}' が指定されていません。")
                    End If
                Next

                ' オプションのパラメータが必要な場合、パラメータが指定されているかをチェック
                For Each opt In Me.SwitchOptions
                    If opt.sw.ParamType <> ParameterType.None AndAlso opt.prm.Length = 0 Then
                        Throw New ArgumentException($"オプション '{opt.sw.Name}' に必要なパラメータが指定されていません。")
                    End If
                Next

                ' オプションの型が正しいかをチェック
                For Each opt In Me.SwitchOptions
                    Dim ptype = (opt.sw.ParamType And Not ParameterType.Array)
                    For Each prm In opt.prm
                        Select Case ptype
                            Case ParameterType.Str
                                ' 文字列型は特にチェックなし
                            Case ParameterType.Int
                                ' 整数型のチェック
                                CheckIntegerThrowException(prm, $"オプション '{opt.sw.Name}' のパラメータ '{prm}' は整数ではありません。")
                            Case ParameterType.Dbl
                                ' 実数型のチェック
                                CheckDoubleThrowException(prm, $"オプション '{opt.sw.Name}' のパラメータ '{prm}' は実数ではありません。")
                            Case ParameterType.URI
                                ' URI型のチェック
                                CheckUriThrowException(prm, $"オプション '{opt.sw.Name}' のパラメータ '{prm}' はURI形式ではありません。")
                        End Select
                    Next
                Next
            End Sub

            ''' <summary>
            ''' 指定されたオプション名に対応するSwitchValueを取得します。
            ''' </summary>
            ''' <param name="swName">オプション名。</param>
            ''' <returns>指定されたオプションのSwitchValue。</returns>
            ''' <remarks>
            ''' このメソッドは、指定されたオプション名に対応するSwitchValueを返します。
            ''' </remarks>
            Public Function GetOption(swName As String) As SwitchValue
                Dim opt = Me.SwitchOptions.Where(Function(o) o.sw.Name = swName).FirstOrDefault()
                Return New SwitchValue(opt.sw, opt.prm)
            End Function

            ''' <summary>パラメータの値をチェックし、指定された型であることを確認します。</summary>
            ''' <param name="prmType">期待されるパラメータの型。</param>
            ''' <param name="prmTypeMsg">エラーメッセージに使用するパラメータの型の説明。</param>
            ''' <remarks>
            ''' このメソッドは、パラメータの値が指定された型であることを確認し、そうでない場合は例外をスローします。
            ''' </remarks>
            Private Sub CheckParameterType(prmType As ParameterType, prmTypeMsg As String)
                If Me.ParameterType = prmType Then
                    If Me.Parameters.Length <= 0 Then
                        Throw New ArgumentException("パラメータが設定されていません")
                    End If
                Else
                    Throw New ArgumentException($"パラメータは{prmTypeMsg}ではありません。")
                End If
            End Sub

            ''' <summary>パラメータの値を文字列として取得します。</summary>
            ''' <returns>パラメータの値を表す文字列。</returns>
            Public Function GetStr() As String
                CheckParameterType(ParameterType.Str, "文字列型")
                Return Me.Parameters(0)
            End Function

            ''' <summary>パラメータの値を整数値として取得します。</summary>
            ''' <returns>パラメータの値を表す整数値。</returns>
            Public Function GetInteger() As Integer
                CheckParameterType(ParameterType.Int, "整数")
                Return CheckIntegerThrowException(Me.Parameters(0), $"パラメータの値は整数ではありません。")
            End Function

            ''' <summary>パラメータの値を実数値として取得します。</summary>
            ''' <returns>パラメータの値を表す実数値。</returns>
            Public Function GetDouble() As Double
                CheckParameterType(ParameterType.Dbl, "実数値")
                Return CheckDoubleThrowException(Me.Parameters(0), $"パラメータの値は実数ではありません。")
            End Function

            ''' <summary>パラメータの値をURIとして取得します。</summary>
            ''' <returns>パラメータの値を表すURI。</returns>
            Public Function GetURI() As Uri
                CheckParameterType(ParameterType.URI, "URI型")
                Return CheckUriThrowException(Me.Parameters(0), $"パラメータの値は URI型ではありません。")
            End Function

            ''' <summary>パラメータの値のURIを絶対パスで取得します。</summary>
            ''' <returns>絶対パス</returns>
            Public Function GetAbsolutePathOfURI() As String
                Dim pathUri = Me.GetURI()
                If pathUri.IsAbsoluteUri Then
                    Return pathUri.LocalPath
                Else
                    Return IO.Path.Combine(Environment.CurrentDirectory, pathUri.OriginalString)
                End If
            End Function

            ''' <summary>パラメータの値を文字列配列として取得します。</summary>
            ''' <returns>パラメータの値を表す文字列配列。</returns>
            Public Function GetStrArray() As String()
                CheckParameterType(ParameterType.Str Or ParameterType.Array, "文字列型")
                Return CType(Me.Parameters.Clone(), String())
            End Function

            ''' <summary>パラメータの値を整数値配列として取得します。</summary>
            ''' <returns>パラメータの値を表す整数配列。</returns>
            Public Function GetIntegerArray() As Integer()
                CheckParameterType(ParameterType.Int Or ParameterType.Array, "整数")
                Dim res As Integer() = New Integer(Me.Parameters.Length - 1) {}
                For i As Integer = 0 To Me.Parameters.Length - 1
                    res(i) = CheckIntegerThrowException(Me.Parameters(i), $"パラメータの値は整数ではありません。")
                Next
                Return res
            End Function

            ''' <summary>パラメータの値を実数値として取得します。</summary>
            ''' <returns>パラメータの値を表す実数値。</returns>
            Public Function GetDoubleArray() As Double()
                CheckParameterType(ParameterType.Dbl Or ParameterType.Array, "実数値")
                Dim res As Double() = New Double(Me.Parameters.Length - 1) {}
                For i As Integer = 0 To Me.Parameters.Length - 1
                    res(i) = CheckDoubleThrowException(Me.Parameters(i), $"パラメータの値は実数ではありません。")
                Next
                Return res
            End Function

            ''' <summary>パラメータの値をURIとして取得します。</summary>
            ''' <returns>パラメータの値を表すURI。</returns>
            Public Function GetURIArray() As Uri()
                CheckParameterType(ParameterType.URI Or ParameterType.Array, "URI型")
                Dim res As Uri() = New Uri(Me.Parameters.Length - 1) {}
                For i As Integer = 0 To Me.Parameters.Length - 1
                    res(i) = CheckUriThrowException(Me.Parameters(i), $"パラメータの値は URI型ではありません。")
                Next
                Return res
            End Function

            ''' <summary>パラメータの値のURIを絶対パスで取得します。</summary>
            ''' <returns>絶対パス</returns>
            Public Function GetAbsoluteArrayPathOfURI() As String()
                CheckParameterType(ParameterType.URI Or ParameterType.Array, "URI型")
                Dim res As String() = New String(Me.Parameters.Length - 1) {}
                For i As Integer = 0 To Me.Parameters.Length - 1
                    Dim pathUri = CheckUriThrowException(Me.Parameters(i), $"パラメータの値は URI型ではありません。")
                    If pathUri.IsAbsoluteUri Then
                        res(i) = pathUri.LocalPath
                    Else
                        res(i) = IO.Path.Combine(Environment.CurrentDirectory, pathUri.OriginalString)
                    End If
                Next
                Return res
            End Function

        End Class

        ''' <summary>
        ''' オプションの値を表す構造体です。
        ''' </summary>
        ''' <remarks>
        ''' この構造体は、オプションの定義とその値を保持します。
        ''' </remarks>
        Public Structure SwitchValue

            ''' <summary>オプションの定義を取得します。</summary>
            Public ReadOnly swOption As SwitchDefine

            ''' <summary>オプションの値を取得します。</summary>
            Public ReadOnly swValue As String()

            ''' <summary>
            ''' コンストラクタ。
            ''' <para>オプションの定義と値を指定して、SwitchValueオブジェクトを初期化します。</para>
            ''' </summary>
            ''' <param name="sw">オプションの定義。</param>
            ''' <param name="prm">オプションの値の配列。</param>
            Public Sub New(sw As SwitchDefine, prm As String())
                Me.swOption = sw
                Me.swValue = prm
            End Sub

            ''' <summary>オプションの値をチェックし、指定された型であることを確認します。</summary>
            ''' <param name="prmType">期待されるパラメータの型。</param>
            ''' <param name="prmTypeMsg">エラーメッセージに使用するパラメータの型の説明。</param>
            ''' <remarks>
            ''' このメソッドは、オプションの値が指定された型であることを確認し、そうでない場合は例外をスローします。
            ''' </remarks>
            Private Sub CheckOptionType(prmType As ParameterType, prmTypeMsg As String)
                If Me.swOption.ParamType = prmType Then
                    If Me.swValue.Length <= 0 Then
                        Throw New ArgumentException($"オプション '{Me.swOption.Name}' が設定されていません")
                    End If
                Else
                    Throw New ArgumentException($"オプション '{Me.swOption.Name}' は{prmTypeMsg}ではありません。")
                End If
            End Sub

            ''' <summary>オプションの値を文字列として取得します。</summary>
            ''' <returns>オプションの値を表す文字列。</returns>
            Public Function GetStr() As String
                CheckOptionType(ParameterType.Str, "文字列型")
                Return Me.swValue(0)
            End Function

            ''' <summary>オプションの値を整数値として取得します。</summary>
            ''' <returns>オプションの値を表す整数値。</returns>
            Public Function GetInteger() As Integer
                CheckOptionType(ParameterType.Int, "整数")
                Return CheckIntegerThrowException(Me.swValue(0), $"オプション '{Me.swOption.Name}' の値は整数ではありません。")
            End Function

            ''' <summary>オプションの値を実数値として取得します。</summary>
            ''' <returns>オプションの値を表す実数値。</returns>
            Public Function GetDouble() As Double
                CheckOptionType(ParameterType.Dbl, "実数値")
                Return CheckDoubleThrowException(Me.swValue(0), $"オプション '{Me.swOption.Name}' の値は実数ではありません。")
            End Function

            ''' <summary>オプションの値をURIとして取得します。</summary>
            ''' <returns>オプションの値を表すURI。</returns>
            Public Function GetURI() As Uri
                CheckOptionType(ParameterType.URI, "URI型")
                Return CheckUriThrowException(Me.swValue(0), $"オプション '{Me.swOption.Name}' の値は URI型ではありません。")
            End Function

            ''' <summary>オプションの値のURIを絶対パスで取得します。</summary>
            ''' <returns>絶対パス</returns>
            Public Function GetAbsolutePathOfURI() As String
                Dim pathUri = Me.GetURI()
                If pathUri.IsAbsoluteUri Then
                    Return pathUri.LocalPath
                Else
                    Return IO.Path.Combine(Environment.CurrentDirectory, pathUri.OriginalString)
                End If
            End Function

            ''' <summary>オプションの値を文字列配列として取得します。</summary>
            ''' <returns>オプションの値を表す文字列配列。</returns>
            Public Function GetStrArray() As String()
                CheckOptionType(ParameterType.Str Or ParameterType.Array, "文字列型")
                Return CType(Me.swValue.Clone(), String())
            End Function

            ''' <summary>オプションの値を整数値配列として取得します。</summary>
            ''' <returns>オプションの値を表す整数配列。</returns>
            Public Function GetIntegerArray() As Integer()
                CheckOptionType(ParameterType.Int Or ParameterType.Array, "整数")
                Dim res As Integer() = New Integer(Me.swValue.Length - 1) {}
                For i As Integer = 0 To Me.swValue.Length - 1
                    res(i) = CheckIntegerThrowException(Me.swValue(i), $"オプション '{Me.swOption.Name}' の値は整数ではありません。")
                Next
                Return res
            End Function

            ''' <summary>オプションの値を実数値として取得します。</summary>
            ''' <returns>オプションの値を表す実数値。</returns>
            Public Function GetDoubleArray() As Double()
                CheckOptionType(ParameterType.Dbl Or ParameterType.Array, "実数値")
                Dim res As Double() = New Double(Me.swValue.Length - 1) {}
                For i As Integer = 0 To Me.swValue.Length - 1
                    res(i) = CheckDoubleThrowException(Me.swValue(i), $"オプション '{Me.swOption.Name}' の値は実数ではありません。")
                Next
                Return res
            End Function

            ''' <summary>オプションの値をURIとして取得します。</summary>
            ''' <returns>オプションの値を表すURI。</returns>
            Public Function GetURIArray() As Uri()
                CheckOptionType(ParameterType.URI Or ParameterType.Array, "URI型")
                Dim res As Uri() = New Uri(Me.swValue.Length - 1) {}
                For i As Integer = 0 To Me.swValue.Length - 1
                    res(i) = CheckUriThrowException(Me.swValue(i), $"オプション '{Me.swOption.Name}' の値は URI型ではありません。")
                Next
                Return res
            End Function

            ''' <summary>オプションの値のURIを絶対パスで取得します。</summary>
            ''' <returns>絶対パス</returns>
            Public Function GetAbsoluteArrayPathOfURI() As String()
                CheckOptionType(ParameterType.URI Or ParameterType.Array, "URI型")
                Dim res As String() = New String(Me.swValue.Length - 1) {}
                For i As Integer = 0 To Me.swValue.Length - 1
                    Dim pathUri = CheckUriThrowException(Me.swValue(i), $"オプション '{Me.swOption.Name}' の値は URI型ではありません。")
                    If pathUri.IsAbsoluteUri Then
                        res(i) = pathUri.LocalPath
                    Else
                        res(i) = IO.Path.Combine(Environment.CurrentDirectory, pathUri.OriginalString)
                    End If
                Next
                Return res
            End Function

        End Structure

    End Class

End Namespace
