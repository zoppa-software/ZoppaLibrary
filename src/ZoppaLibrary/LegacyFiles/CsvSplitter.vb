Option Strict On
Option Explicit On

Imports System.IO
Imports ZoppaLibrary.Strings

Namespace LegacyFiles

    ''' <summary>カンマ区切りで文字列を分割する機能です（EXCEL）</summary>
    Public NotInheritable Class CsvSplitter
        Inherits Splitter

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputStream">入力ストリーム。</param>
        Private Sub New(inputStream As StreamReader)
            MyBase.New(inputStream)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputText">入力文字列。</param>
        Private Sub New(inputText As String)
            MyBase.New(inputText)
        End Sub

        ''' <summary>カンマ区切り分割機能を生成します。</summary>
        ''' <param name="inputStream">入力ストリーム。</param>
        ''' <returns>カンマ区切り分割機能。</returns>
        Public Shared Function CreateSplitter(inputStream As StreamReader) As CsvSplitter
            Return New CsvSplitter(inputStream)
        End Function

        ''' <summary>カンマ区切り分割機能を生成します。</summary>
        ''' <param name="inputText">分割する文字列。</param>
        ''' <returns>カンマ区切り分割機能。</returns>
        Public Shared Function CreateSplitter(inputText As String) As CsvSplitter
            Return New CsvSplitter(inputText)
        End Function

        ''' <summary>行を読み取り、分割された文字列のリストを返します。</summary>
        ''' <param name="readStream">読み取り用のテキストリーダー。</param>
        ''' <returns>分割された文字列のリスト。</returns>
        ''' <remarks>
        ''' このメソッドは、CSV形式の行を読み取り、カンマで区切られた値を分割してリストとして返します。
        ''' エスケープ文字（"）や改行コード（CR, LF）も考慮されます。
        ''' </remarks>
        Protected Overrides Function ReadLine(readStream As TextReader) As List(Of U8String)
            Dim rchars As New List(Of Byte)(4096)
            Dim spoint As New List(Of Integer)(256)
            Dim esc As Boolean = False
            Dim index As Integer = 0

            spoint.Add(0)
            Do While readStream.Peek() <> -1
                Dim c = Text.Encoding.UTF8.GetBytes(Convert.ToChar(readStream.Read()))

                If c.Length = 1 Then
                    Select Case c(0)
                        Case &HD ' CR
                            rchars.Add(&HD)
                            index += 1
                            If Not esc AndAlso readStream.Peek() = &HA Then
                                readStream.Read()
                                rchars.Add(&HA)
                                index -= 1
                                Exit Do
                            End If

                        Case &HA ' LF
                            rchars.Add(&HA)
                            index += 1
                            If Not esc Then
                                index -= 1
                                Exit Do
                            End If

                        Case &H22 ' "
                            rchars.Add(&H22)
                            index += 1

                            If esc Then
                                If readStream.Peek() = &H22 Then
                                    rchars.Add(&H22)
                                    index += 1
                                    readStream.Read()
                                Else
                                    esc = False
                                End If
                            ElseIf rchars.Count < 2 OrElse
                                   rchars(rchars.Count - 2) = &H2C Then
                                esc = True
                            End If

                        Case &H2C ' ,
                            rchars.Add(&H2C)
                            index += 1
                            If Not esc Then
                                spoint.Add(index)
                            End If

                        Case Else
                            rchars.Add(c(0))
                            index += 1
                    End Select

                ElseIf c.Length > 1 Then
                    rchars.AddRange(c)
                    index += c.Length
                End If

            Loop
            spoint.Add(index + 1)

            ' 分割された文字列を生成
            Dim src = U8String.NewStringChangeOwner(rchars.ToArray())
            Dim split As New List(Of U8String)(spoint.Count - 1)
            If src.Length > 0 Then
                For i As Integer = 0 To spoint.Count - 2
                    split.Add(U8String.NewSlice(src, spoint(i), spoint(i + 1) - 1 - spoint(i)))
                Next
            End If
            Return split
        End Function

        ''' <summary>文字列をエスケープ解除します。</summary>
        ''' <param name="target">エスケープされた文字列。</param>
        ''' <returns>エスケープ解除された文字列。</returns>
        ''' <remarks>
        ''' 例えば、"abc,def" のような文字列を abc,def に変換します。
        ''' </remarks>
        Protected Overrides Function UnEscape(target As U8String) As U8String
            Dim str = target.Trim()

            Dim buf As New List(Of Byte)(str.Length)
            Dim esc As Boolean = False
            Dim iter = str.GetIterator()

            While iter.HasNext()
                Dim c As U8Char = iter.Current.Value
                If c.Raw0 = &H22 Then ' "
                    If esc Then
                        If iter.HasNext() AndAlso iter.Peek(1)?.Raw0 = &H22 Then
                            buf.Add(&H2C)
                            iter.MoveNext()
                        Else
                            esc = False
                        End If
                    ElseIf buf.Count < 1 OrElse buf(buf.Count - 1) = &H2C Then
                        esc = True
                    Else
                        buf.Add(c.Raw0)
                    End If
                Else
                    Select Case c.Size
                        Case 1
                            buf.Add(c.Raw0)
                        Case 2
                            buf.Add(c.Raw0)
                            buf.Add(c.Raw1)
                        Case 3
                            buf.Add(c.Raw0)
                            buf.Add(c.Raw1)
                            buf.Add(c.Raw2)
                        Case Else
                            buf.Add(c.Raw0)
                            buf.Add(c.Raw1)
                            buf.Add(c.Raw2)
                            buf.Add(c.Raw3)
                    End Select
                End If
                iter.MoveNext()
            End While

            Return U8String.NewStringChangeOwner(buf.ToArray())
        End Function


    End Class

End Namespace
