Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text

Namespace Parser

    ''' <summary>
    ''' 位置調整可能な文字列リーダー
    ''' </summary>
    Public NotInheritable Class PositionAdjustStringReader
        Implements IPositionAdjustReader

        ''' <summary>
        ''' 元ソース
        ''' </summary>
        Private _source As TextReader

        ''' <summary>
        ''' 読込済みバッファ
        ''' </summary>
        Private ReadOnly _readed As StringBuilder

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
        ''' リソースを解放します。
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            If _source IsNot Nothing Then
                _source.Dispose()
                _source = Nothing
            End If
        End Sub

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="source">元データテキストリーダー。</param>
        Public Sub New(source As TextReader)
            _source = source
            _readed = New StringBuilder()
            _position = 0
        End Sub

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="source">元データ文字列。</param>
        Public Sub New(source As String)
            Me.New(New StringReader(source))
        End Sub

        ''' <summary>
        ''' 1文字読み取り、終了の場合はNothingを返します。
        ''' </summary>
        ''' <returns>読み取った文字。</returns>
        Public Function ReadChar() As Char? Implements IPositionAdjustReader.ReadChar
            Dim res = Me.Read()
            Return If(res = -1, CType(Nothing, Char?), ChrW(res))
        End Function

        ''' <summary>
        ''' 指定位置から指定長さの部分文字列を取得します。
        ''' </summary>
        ''' <param name="startIndex">開始位置。</param>
        ''' <param name="length">長さ。</param>
        ''' <returns>取得した部分文字列。</returns>
        Public Function Substring(startIndex As Integer, length As Integer) As String Implements IPositionAdjustReader.Substring
            ' 引数チェック
            If startIndex < 0 OrElse length < 0 Then
                Throw New ArgumentOutOfRangeException("指定された範囲が不正です")
            End If

            ' 文字列取得
            Dim res As New StringBuilder()
            Dim readcount = 0
            While readcount < length
                Dim c = Me.SubChar(startIndex + readcount)
                If c <> ChrW(0) Then
                    res.Append(c)
                    readcount += 1
                Else
                    Exit While
                End If
            End While

            Return res.ToString()
        End Function

        ''' <summary>
        ''' 指定位置の文字を取得します。
        ''' </summary>
        ''' <param name="pos">位置。</param>
        ''' <returns>取得した文字。</returns>
        Public Function SubChar(pos As Integer) As Char Implements IPositionAdjustReader.SubChar
            If pos < Me._readed.Length Then
                ' 既に読み取り済み
                Dim res = Me._readed(pos)
                Return res
            ElseIf pos = Me._readed.Length Then
                ' まだ読み取っていない
                Dim nextChar As Integer = Me._source.Read()
                If nextChar = -1 Then
                    Return ChrW(0)
                End If
                Dim res = ChrW(nextChar)
                _readed.Append(res)
                Return res
            Else
                Throw New IndexOutOfRangeException("位置が読み取り範囲を超えています")
            End If
        End Function

        ''' <summary>
        ''' 現在位置のスナップショットを取得します。
        ''' </summary>
        ''' <returns>スナップショット。</returns>
        Public Function MemoryPosition() As IPositionAdjustReader.IPosition Implements IPositionAdjustReader.MemoryPosition
            Return New SnapshotPosition(Me)
        End Function

        ''' <summary>
        ''' 次に読み取る文字を確認します。
        ''' </summary>
        ''' <returns>次に読み取る文字。</returns>
        Public Function Peek() As Integer Implements IPositionAdjustReader.Peek
            If Me._position < Me._readed.Length Then
                Return AscW(Me._readed(Me._position))
            ElseIf Me._position = Me._readed.Length Then
                Return Me._source.Peek()
            Else
                Throw New IndexOutOfRangeException("位置が読み取り範囲を超えています")
            End If
        End Function

        ''' <summary>
        ''' 1文字読み取ります。
        ''' </summary>
        ''' <returns>読み取った文字。</returns>
        Public Function Read() As Integer Implements IPositionAdjustReader.Read
            If Me._position < Me._readed.Length Then
                ' 既に読み取り済み
                Dim res = AscW(Me._readed(Me._position))
                Me._position += 1
                Return res
            ElseIf Me._position = Me._readed.Length Then
                ' まだ読み取っていない
                Dim nextChar As Integer = Me._source.Read()
                If nextChar = -1 Then
                    Return -1
                End If
                Me._position += 1
                _readed.Append(ChrW(nextChar))
                Return nextChar
            Else
                Throw New IndexOutOfRangeException("位置が読み取り範囲を超えています")
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
            Dim readCount = 0
            While readCount < count
                Dim nextChar = Me.ReadChar()
                If Not nextChar.HasValue Then
                    Exit While
                End If
                buffer(index + readCount) = nextChar.Value
                readCount += 1
            End While
            Return readCount
        End Function

        ''' <summary>
        ''' 指定された文字数分、末尾からの部分文字列を取得します。
        ''' </summary>
        ''' <param name="count">文字数。</param>
        ''' <returns>取得した部分文字列。</returns>
        Public Function ToLastString(count As Integer) As String Implements IPositionAdjustReader.ToLastString
            Return Me._readed.ToString(Math.Max(0, Me._readed.Length - count), Math.Min(count, Me._readed.Length))
        End Function

        ''' <summary>
        ''' スナップショット位置。
        ''' </summary>
        Public NotInheritable Class SnapshotPosition
            Implements IPositionAdjustReader.IPosition

            ''' <summary>
            ''' 元の文字列リーダー。
            ''' </summary>
            Private ReadOnly _tr As PositionAdjustStringReader

            ''' <summary>
            ''' スナップショット位置。
            ''' </summary>
            Private ReadOnly _position As Integer

            ''' <summary>
            ''' コンストラクタ。
            ''' </summary>
            ''' <param name="tr">元の文字列リーダー。</param>
            Public Sub New(tr As PositionAdjustStringReader)
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