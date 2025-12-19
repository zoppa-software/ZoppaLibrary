Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 位置調整可能なテキストリーダーインターフェースを表します。
    ''' </summary>
    Public Interface IPositionAdjustReader
        Inherits IDisposable

        ''' <summary>
        ''' 現在位置を取得します。
        ''' </summary>
        ReadOnly Property Position As Integer

        ''' <summary>
        ''' 次に読み取る文字を確認します。
        ''' </summary>
        ''' <returns>次に読み取る文字。</returns>
        Function Peek() As Integer

        ''' <summary>
        ''' 指定位置にシークします。
        ''' </summary>
        ''' <param name="searchStart">シーク位置。</param>
        Sub Seek(searchStart As Integer)

        ''' <summary>
        ''' 1文字読み取ります。
        ''' </summary>
        ''' <returns>読み取った文字。</returns>
        Function Read() As Integer

        ''' <summary>
        ''' 複数文字を読み取ります。
        ''' </summary>
        ''' <param name="buffer">読み取り先バッファ。</param>
        ''' <param name="index">書き込み開始位置。</param>
        ''' <param name="count">読み取り文字数。</param>
        ''' <returns>実際に読み取った文字数。</returns>
        Function Read(buffer() As Char, index As Integer, count As Integer) As Integer

        ''' <summary>
        ''' 1文字読み取り、終了の場合はNothingを返します。
        ''' </summary>
        ''' <returns>読み取った文字。</returns>
        Function ReadChar() As Char?

        ''' <summary>
        ''' 指定位置から読み込み済みの部分文字列を取得します。
        ''' </summary>
        ''' <param name="startIndex">開始位置。</param>
        ''' <returns>取得した部分文字列。</returns>
        Function Substring(startIndex As Integer) As String

        ''' <summary>
        ''' 指定位置から指定長さの部分文字列を取得します。
        ''' </summary>
        ''' <param name="startIndex">開始位置。</param>
        ''' <param name="length">長さ。</param>
        ''' <returns>取得した部分文字列。</returns>
        Function Substring(startIndex As Integer, length As Integer) As String

        ''' <summary>
        ''' 指定位置の文字を取得します。
        ''' </summary>
        ''' <param name="pos">位置。</param>
        ''' <returns>取得した文字。</returns>
        Function SubChar(pos As Integer) As Char

        ''' <summary>
        ''' 現在位置のスナップショットを取得します。
        ''' </summary>
        ''' <returns>スナップショット。</returns>
        Function MemoryPosition() As IPosition

        ''' <summary>
        ''' スナップショットインターフェース。
        ''' </summary>
        Public Interface IPosition

            ''' <summary>
            ''' 位置を復元します。
            ''' </summary>
            Sub Restore()

        End Interface

    End Interface

End Namespace