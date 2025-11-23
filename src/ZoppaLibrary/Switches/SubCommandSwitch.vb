Option Strict On
Option Explicit On

Namespace Switches

    ''' <summary>サブコマンドの定義を表します。</summary>
    ''' <remarks>
    ''' このクラスは、サブコマンドの必要性とその定義を含むプロパティを提供します。
    ''' サブコマンドは、特定の機能を実行するために使用されます。
    ''' </remarks>
    Public NotInheritable Class SubCommandSwitch

        ''' <summary>サブコマンドが必要かどうかを示すプロパティ。</summary>
        ''' <remarks>
        ''' Trueの場合、サブコマンドは必須であり、Falseの場合はオプションです。
        ''' </remarks>
        Public ReadOnly Property Required As Boolean

        ''' <summary>サブコマンドの定義を含む配列。</summary>
        ''' <remarks>
        ''' 各サブコマンドは、名前と説明を持つSubCommandDefineオブジェクトとして定義されます。
        ''' </remarks>
        Public ReadOnly Property Commands As SubCommandDefine() = New SubCommandDefine() {}

        ''' <summary>
        ''' コンストラクタ。
        ''' <para>サブコマンドの必要性とその定義を指定して、SubCommandSwitchオブジェクトを初期化します。</para>
        ''' </summary>
        ''' <param name="required">サブコマンドが必要かどうか。</param>
        ''' <param name="commands">サブコマンドの定義の配列。</param>
        Public Sub New(required As Boolean, commands As SubCommandDefine())
            Me.Required = required
            If commands IsNot Nothing Then
                Me.Commands = commands
            End If
        End Sub

    End Class

End Namespace
