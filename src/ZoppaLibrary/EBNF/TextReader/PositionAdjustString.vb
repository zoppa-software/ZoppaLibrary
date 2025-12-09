Option Explicit On
Option Strict On

Namespace EBNF

    ''' <summary>
    ''' 位置調整可能な文字列リーダーを表します。
    ''' </summary>
    Public NotInheritable Class PositionAdjustString
        Implements IPositionAdjustReader

        ''' <summary>
        ''' 元ソース
        ''' </summary>
        Private _source As String

        ''' <summary>
        ''' 現在位置
        ''' </summary>
        Private _position As Integer

        ''' <summary>
        ''' 現在位置を取得します。
        ''' </summary>
        Public ReadOnly Property Position As Integer Implements IPositionAdjustReader.Position
            Get
                Return Me._position
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="source">元ソース。</param>
        Public Sub New(source As String)
            Me._source = source
        End Sub

        ''' <summary>
        ''' リソースを解放します。
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            ' 何もしない
        End Sub

        ''' <summary>
        ''' 次に読み取る文字を確認します。
        ''' </summary>
        ''' <returns>次に読み取る文字。</returns>
        Public Function Peek() As Integer Implements IPositionAdjustReader.Peek
            If Me._position < Me._source.Length Then
                Return AscW(Me._source.Chars(Me._position))
            Else
                Return -1
            End If
        End Function

        ''' <summary>
        ''' 1文字読み取ります。
        ''' </summary>
        ''' <returns>読み取った文字。</returns>
        Public Function Read() As Integer Implements IPositionAdjustReader.Read
            If Me._position < Me._source.Length Then
                Dim c = Me._source.Chars(Me._position)
                Me._position += 1
                Return AscW(c)
            Else
                Return -1
            End If
        End Function

        ''' <summary>
        ''' 複数文字を読み取ります。
        ''' </summary>
        ''' <param name="buffer">読み取り先バッファ。</param>
        ''' <param name="index">書き込み開始位置。</param>
        ''' <param name="count">読み取り文字数。</param>
        ''' <returns>実際に読み取った文字数。</returns>
        Public Function Read(buffer() As Char, index As Integer, count As Integer) As Integer Implements IPositionAdjustReader.Read
            Dim readCount As Integer = 0
            Do While readCount < count AndAlso Me._position < Me._source.Length
                buffer(index + readCount) = Me._source.Chars(Me._position)
                Me._position += 1
                readCount += 1
            Loop
            Return readCount
        End Function

        ''' <summary>
        ''' 1文字読み取り、終了の場合はNothingを返します。
        ''' </summary>
        ''' <returns>読み取った文字。</returns>
        Public Function ReadChar() As Char? Implements IPositionAdjustReader.ReadChar
            If Me._position < Me._source.Length Then
                Dim c = Me._source.Chars(Me._position)
                Me._position += 1
                Return c
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' 指定位置から読み込み済みの部分文字列を取得します。
        ''' </summary>
        ''' <param name="startIndex">開始位置。</param>
        ''' <returns>取得した部分文字列。</returns>
        Public Function Substring(startIndex As Integer) As String Implements IPositionAdjustReader.Substring
            Return Me._source.Substring(startIndex, Me._position - startIndex)
        End Function

        ''' <summary>
        ''' 指定位置から指定長さの部分文字列を取得します。
        ''' </summary>
        ''' <param name="startIndex">開始位置。</param>
        ''' <param name="length">長さ。</param>
        ''' <returns>取得した部分文字列。</returns>
        Public Function Substring(startIndex As Integer, length As Integer) As String Implements IPositionAdjustReader.Substring
            Return Me._source.Substring(startIndex, length)
        End Function

        ''' <summary>
        ''' 指定位置の文字を取得します。
        ''' </summary>
        ''' <param name="pos">位置。</param>
        ''' <returns>取得した文字。</returns>
        Public Function SubChar(pos As Integer) As Char Implements IPositionAdjustReader.SubChar
            Return Me._source.Chars(pos)
        End Function

        ''' <summary>
        ''' 現在位置のスナップショットを取得します。
        ''' </summary>
        ''' <returns>スナップショット。</returns>
        Public Function MemoryPosition() As IPositionAdjustReader.IPosition Implements IPositionAdjustReader.MemoryPosition
            Return New SnapshotPosition(Me)
        End Function

        ''' <summary>
        ''' 指定された文字数分、末尾からの部分文字列を取得します。
        ''' </summary>
        ''' <param name="count">文字数。</param>
        ''' <returns>取得した部分文字列。</returns>
        Public Function ToLastString(count As Integer) As String Implements IPositionAdjustReader.ToLastString
            Return Me._source.Substring(Math.Max(0, Me._position - count), Math.Min(count, Me._position))
        End Function

        ''' <summary>
        ''' スナップショット位置。
        ''' </summary>
        Public NotInheritable Class SnapshotPosition
            Implements IPositionAdjustReader.IPosition

            ''' <summary>
            ''' 元の文字列リーダー。
            ''' </summary>
            Private ReadOnly _tr As PositionAdjustString

            ''' <summary>
            ''' スナップショット位置。
            ''' </summary>
            Private ReadOnly _position As Integer

            ''' <summary>
            ''' コンストラクタ。
            ''' </summary>
            ''' <param name="tr">元の文字列リーダー。</param>
            Public Sub New(tr As PositionAdjustString)
                Me._tr = tr
                Me._position = tr._position
            End Sub

            ''' <summary>
            ''' 位置を復元します。
            ''' </summary>
            Public Sub Restore() Implements IPositionAdjustReader.IPosition.Restore
                Me._tr._position = Me._position
            End Sub

        End Class

    End Class

End Namespace
