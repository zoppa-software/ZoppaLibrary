Option Explicit On
Option Strict On

Namespace BNF

    ''' <summary>
    ''' 位置調整可能なバイトリーダーを表します。
    ''' </summary>
    Public NotInheritable Class PositionAdjustBytes

        ''' <summary>
        ''' 元ソース
        ''' </summary>
        Private _source As Byte()

        ''' <summary>
        ''' 現在位置
        ''' </summary>
        Private _position As Integer

        ''' <summary>
        ''' 現在位置を取得します。
        ''' </summary>
        Public ReadOnly Property Position As Integer
            Get
                Return Me._position
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="source">元ソース。</param>
        Public Sub New(source As Byte())
            Me._source = source
        End Sub

        ''' <summary>
        ''' リソースを解放します。
        ''' </summary>
        Public Sub Dispose()
            ' 何もしない
        End Sub

        ''' <summary>
        ''' 次に読み取る文字を確認します。
        ''' </summary>
        ''' <returns>次に読み取る文字。</returns>
        Public Function Peek() As Integer
            If Me._position < Me._source.Length Then
                Return Me._source(Me._position)
            Else
                Return -1
            End If
        End Function

        ''' <summary>
        ''' 指定位置にシークします。
        ''' </summary>
        ''' <param name="searchStart">シーク位置。</param>
        Public Sub Seek(searchStart As Integer)
            If searchStart < Me._source.Length Then
                Me._position = searchStart
            Else
                Me._position = Me._source.Length
            End If
        End Sub

        ''' <summary>
        ''' 1バイト読み取ります。
        ''' </summary>
        ''' <returns>読み取った文字。</returns>
        Public Function Read() As Integer
            If Me._position < Me._source.Length Then
                Dim c = Me._source(Me._position)
                Me._position += 1
                Return c
            Else
                Return -1
            End If
        End Function

        ''' <summary>
        ''' 複数バイトを読み取ります。
        ''' </summary>
        ''' <param name="buffer">読み取り先バッファ。</param>
        ''' <param name="index">書き込み開始位置。</param>
        ''' <param name="count">読み取り文字数。</param>
        ''' <returns>実際に読み取った文字数。</returns>
        Public Function Read(buffer() As Byte, index As Integer, count As Integer) As Integer
            Dim readCount As Integer = 0
            Do While readCount < count AndAlso Me._position < Me._source.Length
                buffer(index + readCount) = Me._source(Me._position)
                Me._position += 1
                readCount += 1
            Loop
            Return readCount
        End Function

        ''' <summary>
        ''' 複数バイトを指定位置から読み取ります。
        ''' </summary>
        ''' <param name="buffer">読み取り先バッファ。</param>
        ''' <param name="srcIndex">読み取り開始位置。</param>
        ''' <param name="destIndex">書き込み開始位置。</param>
        ''' <param name="count">読み取り文字数。</param>
        ''' <returns>実際に読み取った文字数。</returns>
        Public Function Read(buffer() As Byte, srcIndex As Integer, destIndex As Integer, count As Integer) As Integer
            Dim readCount As Integer = 0
            Dim position As Integer = srcIndex
            Do While readCount < count AndAlso position < Me._source.Length
                buffer(destIndex + readCount) = Me._source(position)
                position += 1
                readCount += 1
            Loop
            Return readCount
        End Function

        ''' <summary>
        ''' 指定位置の1バイト読み取ります。
        ''' </summary>
        ''' <param name="position">読み取り位置。</param>
        ''' <returns>読み取った文字。</returns>
        Public Function ReadAt(position As Integer) As Byte
            If position >= 0 AndAlso position < Me._source.Length Then
                Dim c = Me._source(position)
                Return c
            Else
                Throw New IndexOutOfRangeException("読み取り位置が範囲を超えています")
            End If
        End Function

        ''' <summary>
        ''' 現在位置のスナップショットを取得します。
        ''' </summary>
        ''' <returns>スナップショット。</returns>
        Public Function MemoryPosition() As IPositionAdjustReader.IPosition
            Return New SnapshotPosition(Me)
        End Function

        ''' <summary>
        ''' 指定位置から部分文字列を取得します。
        ''' </summary>
        ''' <param name="position">開始位置。</param>
        ''' <param name="length">長さ。</param>
        ''' <returns>部分文字列。</returns>
        Friend Function Substring(position As Integer, Optional length As Integer = Integer.MaxValue) As String
            Dim strBuilder As New System.Text.StringBuilder()
            For i As Integer = position To position + length - 1
                If i >= Me._source.Length Then
                    Exit For
                End If
                Dim b As Byte = Me._source(i)
                strBuilder.Append(String.Format("{0:X2}({1}) ", b, If(b >= &H20 AndAlso b <= &H7E, ChrW(b), " "c)))
            Next

            Return strBuilder.ToString().TrimEnd()
        End Function

        ''' <summary>
        ''' スナップショット位置。
        ''' </summary>
        Public NotInheritable Class SnapshotPosition
            Implements IPositionAdjustReader.IPosition

            ''' <summary>
            ''' 元の文字列リーダー。
            ''' </summary>
            Private ReadOnly _tr As PositionAdjustBytes

            ''' <summary>
            ''' スナップショット位置。
            ''' </summary>
            Private ReadOnly _position As Integer

            ''' <summary>
            ''' コンストラクタ。
            ''' </summary>
            ''' <param name="tr">元の文字列リーダー。</param>
            Public Sub New(tr As PositionAdjustBytes)
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
