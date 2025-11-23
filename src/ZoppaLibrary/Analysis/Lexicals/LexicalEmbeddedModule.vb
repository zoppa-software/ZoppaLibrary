Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 埋込式の字句解析に関連する機能を提供します。。
    ''' このモジュールは、埋め込み式の解析や評価に関連する機能を提供します。
    ''' </summary>
    ''' <remarks>
    ''' 埋込式は、他の式と組み合わせて使用されることがあります。
    ''' </remarks>
    Public Module LexicalEmbeddedModule

        ''' <summary>Ifブロック。</summary>
        Private ReadOnly IfBlockString As U8String = U8String.NewString("{if")

        ''' <summary>ElseIfブロック。</summary>
        Private ReadOnly ElseIfBlockString As U8String = U8String.NewString("{else if")

        ''' <summary>Elseブロック。</summary>
        Private ReadOnly ElseBlockString As U8String = U8String.NewString("{else}")

        ''' <summary>EndIfブロック。</summary>
        Private ReadOnly EndIfBlockString As U8String = U8String.NewString("{/if}")

        ''' <summary>Forブロック。</summary>
        Private ReadOnly ForBlockString As U8String = U8String.NewString("{for")

        ''' <summary>EndForブロック。</summary>
        Private ReadOnly EndForBlockString As U8String = U8String.NewString("{/for}")

        ''' <summary>Selectブロック。</summary>
        Private ReadOnly SelectBlockString As U8String = U8String.NewString("{select")

        ''' <summary>SelectCaseブロック。</summary>
        Private ReadOnly SelectCaseBlockString As U8String = U8String.NewString("{case")

        ''' <summary>SelectDefaultブロック。</summary>
        Private ReadOnly SelectDefaultBlockString As U8String = U8String.NewString("{default}")

        ''' <summary>EndSelectブロック。</summary>
        Private ReadOnly EndSelectBlockString As U8String = U8String.NewString("{/select}")

        ''' <summary>Setブロック。</summary>
        Private ReadOnly SetBlockString As U8String = U8String.NewString("{set")

        ''' <summary>Brブロック。</summary>
        Private ReadOnly BrBlockString As U8String = U8String.NewString("{br}")

        ''' <summary>仮想Brブロック。</summary>
        Private ReadOnly VlBrBlockString As U8String = U8String.NewString("{vr}")

        ''' <summary>Trimブロック。</summary>
        Private ReadOnly TrimBlockString As U8String = U8String.NewString("{trim")

        ''' <summary>EndTrimブロック。</summary>
        Private ReadOnly EndTrimBlockString As U8String = U8String.NewString("{/trim}")

        ''' <summary>Remブロック。</summary>
        Private ReadOnly RemBlockString As U8String = U8String.NewString("{remove")

        ''' <summary>EndRemブロック。</summary>
        Private ReadOnly EndRemBlockString As U8String = U8String.NewString("{/remove}")

        ''' <summary>Emptyブロック。</summary>
        Private ReadOnly EmptyBlockString As U8String = U8String.NewString("{}")

        ''' <summary>
        ''' 埋込ブロックを文字列から分割します。
        ''' このメソッドは、入力文字列を解析して埋込ブロックのリストを生成します。
        ''' 埋込ブロックは、特定の文字（例: {, #, !, $）で始まる部分を表します。
        ''' </summary>
        ''' <param name="input">入力文字列。</param>
        ''' <returns>埋込ブロックリスト。</returns>
        <Extension()>
        Public Function SplitEmbeddedText(input As U8String) As EmbeddedBlock()
            Dim embedded As New List(Of EmbeddedBlock)()

            Dim iter = input.GetIterator()
            While iter.HasNext()
                If iter.Current IsNot Nothing Then
                    Dim c = iter.Current.Value
                    Dim embed As EmbeddedBlock
                    If c.Size = 1 Then
                        ' 1文字の場合はトークン解析します
                        Select Case c.Raw0
                            Case &H7B ' {
                                ' 埋込式ブロックを取得します
                                embed = GetStatementBlock(input, iter)
                            Case &H21 ' !
                                ' 非エスケープ埋込ブロック
                                embed = GetSpecialBlock(input, iter, EmbeddedType.NoEscapeUnfold)
                            Case &H23 ' #
                                ' 展開埋込ブロック
                                embed = GetSpecialBlock(input, iter, EmbeddedType.Unfold)
                            Case &H24 ' $
                                ' 変数宣言ブロック
                                embed = GetSpecialBlock(input, iter, EmbeddedType.VariableDefine)
                            Case Else
                                ' 非埋込ブロック
                                embed = New EmbeddedBlock(EmbeddedType.None, GetNoneEmbeddedBlock(input, iter))
                        End Select
                    Else
                        ' 非埋込ブロックを取得します
                        embed = New EmbeddedBlock(EmbeddedType.None, GetNoneEmbeddedBlock(input, iter))
                    End If
                    embedded.Add(embed)
                End If
            End While

            ' 埋込ブロックリストを返します
            Return embedded.ToArray()
        End Function

        ''' <summary>埋込ブロックを取得します。</summary>
        ''' <param name="input">入力文字列。</param>
        ''' <param name="iter">文字列のイテレーター。</param>
        ''' <param name="isEmbeddedText">先頭1文字をスキップする。</param>
        ''' <returns>埋込ブロック。</returns>
        ''' <remarks>
        ''' このメソッドは、埋込ブロックを取得します。
        ''' </remarks>
        Private Function GetEmbeddedBlock(input As U8String,
                                          iter As U8String.U8StringIterator,
                                          isEmbeddedText As Boolean) As U8String
            Dim startIndex = iter.CurrentIndex
            Dim closed = False

            ' #, !, $ の場合は一文字飛ばす
            If isEmbeddedText Then
                iter.MoveNext()
            End If
            iter.MoveNext()

            ' 埋め込み式の終わりを探す
            While iter.HasNext()
                Dim pc = iter.MoveNext
                If pc?.Size = 1 Then
                    Select Case pc?.Raw0
                        Case &H7D ' }
                            ' } が見つかった場合、ループを終了
                            closed = True
                            Exit While

                        Case &H5C ' \
                            ' エスケープ文字が見つかった場合の { または } を無視
                            Dim nc = iter.Current
                            If nc?.Size = 1 AndAlso (nc?.Raw0 = &H7B OrElse nc?.Raw0 = &H7D) Then ' { or }
                                iter.MoveNext()
                            End If

                        Case Else
                            ' 他の文字は無視
                    End Select
                End If
            End While

            If Not closed Then
                ' 埋め込み式が閉じられていない場合はエラー
                Throw New AnalysisException("埋め込み式が閉じられていません。")
            End If

            ' 文字列を切り出して返す
            Return U8String.NewSlice(input, startIndex, iter.CurrentIndex - startIndex)
        End Function

        ''' <summary>非埋込ブロックを取得します。</summary>
        ''' <param name="input">入力文字列。</param>
        ''' <param name="iter">文字列のイテレーター。</param>
        ''' <returns>埋込ブロック。</returns>
        ''' <remarks>
        ''' このメソッドは、非埋込ブロックを取得します。
        ''' </remarks>
        Private Function GetNoneEmbeddedBlock(input As U8String, iter As U8String.U8StringIterator) As U8String
            Dim startIndex = iter.CurrentIndex

            While iter.HasNext()
                Dim pc = iter.Current
                If pc?.Size = 1 Then
                    Select Case pc?.Raw0
                        Case &H7B ' {
                            ' { が見つかったらテキスト部の終了
                            Exit While

                        Case &H23, &H21, &H24 ' #, !, $
                            ' #{, !{, ${ が見つかったらテキスト部の終了
                            Dim nc = iter.Peek(1)
                            If nc?.Size = 1 AndAlso nc?.Raw0 = &H7B Then ' {
                                Exit While
                            End If

                        Case &H5C ' \
                            ' エスケープ文字が見つかった場合は次の文字を無視
                            Dim nc1 = iter.Peek(1)
                            If nc1?.Size = 1 AndAlso (nc1?.Raw0 = &H7B OrElse nc1?.Raw0 = &H7D) Then ' { or }
                                iter.MoveNext()
                            Else
                                Dim nc2 = iter.Peek(2)
                                If nc1?.Size = 1 AndAlso (nc1?.Raw0 = &H23 OrElse nc1?.Raw0 = &H21 OrElse nc1?.Raw0 = &H24) AndAlso
                                   nc2?.Size = 1 AndAlso nc2?.Raw0 = &H7B Then ' 1文字目(#, !, $), 2文字目({)
                                    iter.MoveNext()
                                    iter.MoveNext()
                                End If
                            End If
                    End Select
                End If
                iter.MoveNext()
            End While

            ' 文字列を切り出して返す
            Return U8String.NewSlice(input, startIndex, iter.CurrentIndex - startIndex)
        End Function

        ''' <summary>埋込式ブロックを取得します。</summary>
        ''' <param name="input">入力文字列。</param>
        ''' <param name="iter">文字列のイテレーター。</param>
        ''' <returns>埋込式ブロック。</returns>
        ''' <remarks>
        ''' このメソッドは、埋込式ブロックを取得します。
        ''' </remarks>
        Private Function GetStatementBlock(input As U8String, iter As U8String.U8StringIterator) As EmbeddedBlock
            Dim cmd = GetEmbeddedBlock(input, iter, False)
            If cmd.StartWith(IfBlockString) AndAlso If(cmd.At(3)?.IsWhiteSpace, True) Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.IfBlock, cmd.Mid(4, cmd.Length - 5))
            ElseIf cmd.StartWith(ElseIfBlockString) AndAlso If(cmd.At(8)?.IsWhiteSpace, True) Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.ElseIfBlock, cmd.Mid(9, cmd.Length - 10))
            ElseIf cmd = ElseBlockString Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.ElseBlock, cmd)
            ElseIf cmd = EndIfBlockString Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.EndIfBlock, cmd)
            ElseIf cmd.StartWith(ForBlockString) AndAlso If(cmd.At(4)?.IsWhiteSpace, True) Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.ForBlock, cmd.Mid(5, cmd.Length - 6))
            ElseIf cmd = EndForBlockString Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.EndForBlock, cmd)
            ElseIf cmd.StartWith(SelectBlockString) AndAlso If(cmd.At(7)?.IsWhiteSpace, True) Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.SelectBlock, cmd.Mid(8, cmd.Length - 9))
            ElseIf cmd.StartWith(SelectCaseBlockString) AndAlso If(cmd.At(5)?.IsWhiteSpace, True) Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.SelectCaseBlock, cmd.Mid(6, cmd.Length - 7))
            ElseIf cmd = SelectDefaultBlockString Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.SelectDefaultBlock, cmd)
            ElseIf cmd = EndSelectBlockString Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.EndSelectBlock, cmd)
            ElseIf cmd = EmptyBlockString Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.EmptyBlock, cmd)
            ElseIf cmd.StartWith(SetBlockString) Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.SetBlock, cmd.Mid(5, cmd.Length - 6))
            ElseIf cmd.StartWith(TrimBlockString) Then
                Return New EmbeddedBlock(EmbeddedType.TrimBlock, If(cmd.Length > 7, cmd.Mid(6, cmd.Length - 7), U8String.Empty))
            ElseIf cmd = EndTrimBlockString Then
                Return New EmbeddedBlock(EmbeddedType.EndTrimBlock, cmd)
            ElseIf cmd = BrBlockString Then
                Return New EmbeddedBlock(EmbeddedType.BrBlock, cmd)
            ElseIf cmd = VlBrBlockString Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.VlBrBlock, cmd)
            ElseIf cmd.StartWith(RemBlockString) Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.RemoveBlock, If(cmd.Length > 9, cmd.Mid(8, cmd.Length - 9), U8String.Empty))
            ElseIf cmd = EndRemBlockString Then
                SkipBrWord(iter)
                Return New EmbeddedBlock(EmbeddedType.EndRemoveBlock, cmd)
            Else
                ' 埋込ブロックが認識できない場合はエラーを返す
                Throw New AnalysisException("無効な埋込ブロック: " & cmd.ToString())
            End If
        End Function

        ''' <summary>仮想ブロックの空白をスキップします。</summary>
        ''' <param name="iter">文字列のイテレーター。</param>
        ''' <remarks>
        ''' 仮想ブロックは改行を含む可能性があるため、改行をスキップします。
        ''' </remarks>
        Private Sub SkipBrWord(iter As U8String.U8StringIterator)
            Dim bkidx = iter.CurrentIndex

            ' 空白スペースをスキップします
            Do
                Dim pc = iter.Current
                If Not pc?.IsWhiteSpace OrElse pc?.Raw0 = &HA OrElse pc?.Raw0 = &HD Then
                    Exit Do
                End If
                iter.MoveNext()
            Loop While iter.HasNext()

            ' 仮想ブロックは改行を含む可能性があるため、改行をスキップ
            Dim remf = False
            While iter.HasNext() AndAlso (iter.Current?.Raw0 = &HA OrElse iter.Current?.Raw0 = &HD)
                iter.MoveNext()
                remf = True
            End While

            If Not remf Then
                iter.SetCurrentIndex(bkidx)
            End If
        End Sub

        ''' <summary>特殊ブロックを取得します。</summary>
        ''' <param name="input">入力文字列。</param>
        ''' <param name="iter">文字列のイテレーター。</param>
        ''' <param name="kind">ブロックの種類。</param>
        ''' <returns>埋込ブロック。</returns>
        ''' <remarks>
        ''' このメソッドは、特殊な埋込ブロックを取得します。
        ''' </remarks>
        Private Function GetSpecialBlock(input As U8String, iter As U8String.U8StringIterator, kind As EmbeddedType) As EmbeddedBlock
            Dim lc = iter.Peek(1)
            If lc?.Size = 1 AndAlso lc?.Raw0 = &H7B Then
                Return New EmbeddedBlock(kind, GetEmbeddedBlock(input, iter, True))
            Else
                Return New EmbeddedBlock(EmbeddedType.None, GetNoneEmbeddedBlock(input, iter))
            End If
        End Function

        ''' <summary>埋込ブロックを表す構造体です。</summary>
        ''' <remarks>
        ''' 埋込ブロックは、文字列とその種類を持つブロックとして表現されます。
        ''' </remarks>
        Public Structure EmbeddedBlock

            ''' <summary>埋込ブロックの種類。</summary>
            Public ReadOnly Property Kind As EmbeddedType

            ''' <summary>埋込式の文字列。</summary>
            Public ReadOnly Property Str As U8String

            ''' <summary>埋込ブロックの種類を設定します。</summary>
            ''' <param name="kind">埋込ブロックの種類。</param>
            ''' <param name="str">埋込式の文字列。</param>
            Public Sub New(kind As EmbeddedType, str As U8String)
                Me.Kind = kind
                Me.Str = str
            End Sub

        End Structure

    End Module

End Namespace
