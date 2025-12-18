Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' EBNF に関する例外を表します。
    ''' </summary>
    Public Class EBNFException
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
