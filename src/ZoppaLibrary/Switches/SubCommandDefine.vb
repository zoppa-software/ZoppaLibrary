Option Strict On
Option Explicit On

Namespace Switches

    ''' <summary>サブコマンドの定義を表します。</summary>
    ''' <remarks>
    ''' このクラスは、サブコマンドの名前と説明を定義します。
    ''' サブコマンドは、コマンドラインで特定の機能を実行するために使用されます。
    ''' </remarks>
    Public NotInheritable Class SubCommandDefine

        ''' <summary>オプションの名前。</summary>
        ''' <remarks>オプションの名前は、コマンドラインで使用される識別子です。</remarks>
        Public ReadOnly Property Name As String

        ''' <summary>オプションの説明。</summary>
        ''' <remarks>オプションの説明は、コマンドラインでの使用方法や目的を説明します。</remarks>
        ''' <example>このオプションは、特定の機能を有効にします。</example>
        ''' <example>このオプションは、データの出力形式を指定します。</example>
        Public ReadOnly Property Description As String

        ''' <summary>
        ''' コンストラクタ。
        ''' <para>オプションの名前と説明を指定して、SubCommandDefineオブジェクトを初期化します。</para>
        ''' </summary>
        ''' <param name="name">名前。</param>
        ''' <param name="description">説明。</param>
        Public Sub New(name As String, description As String)
            Me.Name = name
            Me.Description = description
        End Sub

        ''' <summary>オブジェクトの文字列表現を返します。</summary>
        Public Overrides Function ToString() As String
            Return $"{Name}: {Description}"
        End Function

    End Class

End Namespace
