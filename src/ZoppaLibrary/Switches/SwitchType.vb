Option Strict On
Option Explicit On

Namespace Switches

    ''' <summary>
    ''' コマンドラインスイッチの種類を表す列挙体。
    ''' </summary>
    ''' <remarks>
    ''' SingleHyphen: 単一ハイフン（-）で始まるスイッチ。
    ''' DoubleHyphen: 二重ハイフン（--）で始まるスイッチ。
    ''' Slash: スラッシュ（/）で始まるスイッチ。
    ''' </remarks>
    Public Enum SwitchType

        ''' <summary>単一ハイフン（-）で始まるスイッチ。</summary>
        ''' <remarks>
        ''' 単一ハイフンは、短いオプション名を指定するために使用されます。
        ''' 例: -o
        ''' </remarks>
        SingleHyphen = 1

        ''' <summary>二重ハイフン（--）で始まるスイッチ。</summary>
        ''' <remarks>
        ''' 二重ハイフンは、長いオプション名を指定するために使用されます。
        ''' 例: --option-name
        ''' </remarks>
        DoubleHyphen = 2

        ''' <summary>スラッシュ（/）で始まるスイッチ。</summary>
        ''' <remarks>
        ''' スラッシュは、特にWindowsのコマンドラインで使用されることが多い形式です。
        ''' 例: /option
        ''' </remarks>  
        Slash = 3

    End Enum

End Namespace
