Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 非埋込テキストを表す式。
    ''' この式は、埋め込まれていないテキストを保持します。
    ''' </summary>
    ''' <remarks>
    ''' この式は、埋め込まれていないテキストを表現し、式の型を提供します。
    ''' </remarks>
    Structure PlainTextExpression
        Implements IExpression

        ' 非埋込テキスト
        Private ReadOnly _text As U8String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="text">非埋込テキスト。</param>
        Public Sub New(text As U8String)
            _text = text
        End Sub

        ''' <summary>式の型を取得します。</summary>
        ''' <returns>式の型。</returns>
        Public ReadOnly Property Type As ExpressionType Implements IExpression.Type
            Get
                Return ExpressionType.PlainTextExpression
            End Get
        End Property

        ''' <summary>式の値を取得します。</summary>
        ''' <param name="venv">変数環境。</param>
        ''' <returns>PlainTextExpress の値。</returns>
        Public Function GetValue(venv As AnalysisEnvironment) As IValue Implements IExpression.GetValue
            If IsEscape(Me._text) Then
                ' エスケープ文字が含まれている場合は、エスケープを解除して返す
                Return New StringValue(Unescape(_text))
            Else
                ' エスケープ文字が含まれていない場合はそのまま返す
                Return New StringValue(_text)
            End If
        End Function

        ''' <summary>エスケープ文字が含まれているかどうかをチェックします。</summary>
        ''' <param name="target">チェックする文字列。</param>
        ''' <returns>エスケープ文字が含まれている場合はTrue、それ以外はFalse。</returns>
        Private Shared Function IsEscape(target As U8String) As Boolean
            Dim iter = target.GetIterator()
            While iter.HasNext
                With iter.Current
                    If .HasValue AndAlso .Value.Size = 1 AndAlso .Value.Raw0 = &H5C Then ' バックスラッシュ(\)
                        Return True
                    End If
                End With
                iter.MoveNext()
            End While
            Return False
        End Function

        ''' <summary>エスケープ文字を解除します。</summary>
        ''' <param name="target">エスケープ解除する文字列。</param>
        ''' <returns>エスケープ解除された文字列。</returns>
        Private Shared Function Unescape(target As U8String) As U8String
            Dim result As New List(Of Byte)(target.ByteLength)
            Dim iter = target.GetIterator()
            While iter.HasNext
                Dim esc = False

                Dim pc = iter.MoveNext
                If pc.HasValue Then
                    Dim c = pc.Value
                    If c.Size = 1 AndAlso c.Raw0 = &H5C Then
                        Dim nc1 = iter.Peek(0)
                        If nc1?.Size = 1 AndAlso (nc1?.Raw0 = &H7B OrElse nc1?.Raw0 = &H7D) Then ' { or }
                            iter.MoveNext()
                            esc = True
                            result.Add(nc1.Value.Raw0)
                        Else
                            Dim nc2 = iter.Peek(1)
                            If nc1?.Size = 1 AndAlso (nc1?.Raw0 = &H23 OrElse nc1?.Raw0 = &H21 OrElse nc1?.Raw0 = &H24) AndAlso
                                   nc2?.Size = 1 AndAlso nc2?.Raw0 = &H7B Then ' 1文字目(#, !, $), 2文字目({)
                                iter.MoveNext()
                                iter.MoveNext()
                                esc = True
                                result.Add(nc1.Value.Raw0)
                                result.Add(nc2.Value.Raw0)
                            End If
                        End If
                    End If

                    If Not esc Then
                        ' 通常の文字を追加
                        Select Case c.Size
                            Case 1
                                result.Add(c.Raw0)
                            Case 2
                                result.Add(c.Raw0)
                                result.Add(c.Raw1)
                            Case 3
                                result.Add(c.Raw0)
                                result.Add(c.Raw1)
                                result.Add(c.Raw2)
                            Case Else
                                result.Add(c.Raw0)
                                result.Add(c.Raw1)
                                result.Add(c.Raw2)
                                result.Add(c.Raw3)
                        End Select
                    End If
                End If
            End While
            Return U8String.NewStringChangeOwner(result.ToArray())
        End Function

    End Structure

End Namespace
