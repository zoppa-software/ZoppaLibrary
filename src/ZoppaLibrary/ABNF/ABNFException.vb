Option Explicit On
Option Strict On

Namespace ABNF

    ''' <summary>
    ''' ABNF 解析例外を表します。
    ''' </summary>
    Public Class ABNFException
        Inherits Exception

        ''' <summary>
        ''' 新しい EBNFException のインスタンスを初期化する。
        ''' </summary>
        ''' <param name="message">例外メッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

    End Class

End Namespace
