Option Strict On
Option Explicit On

Namespace Switches

    ''' <summary>コマンドラインオプションの定義を表します。</summary>
    ''' <remarks>
    ''' このクラスは、コマンドラインオプションの名前、説明、およびパラメータの型を定義します。
    ''' </remarks>
    Public NotInheritable Class SwitchDefine

        ''' <summary>オプションの名前。</summary>
        ''' <remarks>オプションの名前は、コマンドラインで使用される識別子です。</remarks>
        ''' <example>--option1</example>
        ''' <example>-o1</example>
        ''' <example>/option1</example>
        Public ReadOnly Property Name As String

        ''' <summary>オプションが必要かどうかを示すプロパティ。</summary>
        ''' <remarks>
        ''' Trueの場合、オプションは必須であり、Falseの場合はオプションです。
        ''' </remarks>
        Public ReadOnly Property Required As Boolean

        ''' <summary>オプションのスイッチの種類。</summary>
        ''' <remarks>
        ''' スイッチの種類は、オプションがどのようにコマンドラインで指定されるかを示します。
        ''' 例: 単一ハイフン（-）、二重ハイフン（--）、スラッシュ（/）など。
        ''' </remarks>
        Public ReadOnly Property SwType As SwitchType

        ''' <summary>オプションのパラメータの型。</summary>
        ''' <remarks>パラメータの型は、オプションが受け取る値の型を示します。</remarks>
        Public ReadOnly Property ParamType As ParameterType

        ''' <summary>オプションの説明。</summary>
        ''' <remarks>オプションの説明は、コマンドラインでの使用方法や目的を説明します。</remarks>
        Public ReadOnly Property Description As String

        ''' <summary>
        ''' コンストラクタ。
        ''' <para>オプションの名前、説明、およびパラメータの型を指定して、SwitchDefineオブジェクトを初期化します。</para>
        ''' </summary>
        ''' <param name="name">名前。</param>
        ''' <param name="swType"スイッチの種類。</param>
        ''' <param name="paramType">パラメータの型。</param>
        ''' <param name="description">説明。</param>
        ''' <remarks>
        ''' このコンストラクタは、オプションの名前、説明、スイッチの種類、およびパラメータの型を指定して、
        ''' SwitchDefineオブジェクトを初期化します。
        ''' </remarks>
        Public Sub New(name As String, required As Boolean, swType As SwitchType, paramType As ParameterType, description As String)
            Me.Name = name
            Me.Required = required
            Me.SwType = swType
            If swType <> SwitchType.DoubleHyphen Then
                If name.Length <> 1 Then
                    Throw New ArgumentException("ショートスイッチの名前は1文字でなければなりません。", NameOf(name))
                End If
            End If
            Me.ParamType = paramType
            Me.Description = description
        End Sub

        ''' <summary>オブジェクトの文字列表現を返します。</summary>
        Public Overrides Function ToString() As String
            Return $"{Name}: {SwType},{ParamType}:{Description}"
        End Function

    End Class

End Namespace