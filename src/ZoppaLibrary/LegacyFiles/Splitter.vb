Option Strict On
Option Explicit On

Imports System.IO
Imports ZoppaLibrary.Analysis
Imports ZoppaLibrary.Strings

Namespace LegacyFiles

    ''' <summary>文字列分割機能（共通）</summary>
    Public MustInherit Class Splitter

        ' 内部ストリーム
        Private ReadOnly _innerStream As TextReader

        ' 列名の配列
        Private ReadOnly _columns As String()

        ''' <summary>EOFかどうかを取得します。</summary>
        ''' <returns>EOFならTrue、そうでなければFalse。</returns>
        ''' <remarks>
        ''' EOFは、ストリームの終端に達したことを示します。
        ''' </remarks>
        Public ReadOnly Property IsEOF As Boolean
            Get
                Return Me._innerStream Is Nothing OrElse Me._innerStream.Peek() = -1
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputStream">入力ストリーム。</param>
        Protected Sub New(inputStream As StreamReader)
            Me._innerStream = inputStream
            Me._columns = ReadHeader()
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputText">入力文字列。</param>
        Protected Sub New(inputText As String)
            Me._innerStream = New StringReader(inputText)
            Me._columns = ReadHeader()
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputStream">テキストストリーム。</param>
        Protected Sub New(reader As TextReader)
            Me._innerStream = reader
            Me._columns = ReadHeader()
        End Sub

        ''' <summary>列名を取得します。</summary>
        ''' <returns>列名の配列。</returns>
        Public Function ReadHeader() As String()
            Return ReadLine().Select(Function(s) s.ToString().Trim()).ToArray()
        End Function

        ''' <summary>内部より一行を読み込み、分割して返します。</summary>
        ''' <returns>分割した項目の配列。</returns>
        Public Function Split() As DynamicObject
            Dim ans = Me.ReadLine()
            Dim res As New DynamicObject()
            For i As Integer = 0 To Math.Min(ans.Count, Me._columns.Length) - 1
                res(Me._columns(i)) = UnEscape(ans(i))
            Next
            Return res
        End Function

        ''' <summary>一行読み込み、読み込み結果を取得します。</summary>
        ''' <returns>読み込み結果。</returns>
        Protected Function ReadLine() As List(Of U8String)
            If Me._innerStream IsNot Nothing Then
                Return Me.ReadLine(Me._innerStream)
            Else
                Throw New NullReferenceException("テキストストリームが設定されていません")
            End If
        End Function

        ''' <summary>一行読み込み、分割を行います。</summary>
        ''' <param name="readStream">読み込みストリーム。</param>
        ''' <returns>読み込んだ文字列と分割位置リスト。</returns>
        Protected MustOverride Function ReadLine(readStream As TextReader) As List(Of U8String)

        ''' <summary>文字列をエスケープ解除します。</summary>
        ''' <param name="target">エスケープされた文字列。</param>
        ''' <returns>エスケープ解除された文字列。</returns>
        ''' <remarks>
        ''' 例えば、"abc,def" のような文字列を abc,def に変換します。
        ''' </remarks>
        Protected MustOverride Function UnEscape(target As U8String) As U8String

    End Class

End Namespace
