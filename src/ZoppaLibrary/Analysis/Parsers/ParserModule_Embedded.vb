Option Strict On
Option Explicit On

Imports ZoppaLibrary.Strings

Namespace Analysis

    Partial Module ParserModule

        ''' <summary>
        ''' 埋め込みテキストを解析します。
        ''' </summary>
        ''' <param name="iter">パーサーイテレーター。</param>
        ''' <returns>解析された式。</returns>
        ''' <remarks>
        ''' このメソッドは、埋め込みテキストを解析し、式を生成します。
        ''' </remarks>
        Private Function ParseEmbeddedText(iter As ParserIterator(Of EmbeddedBlock)) As IExpression
            ' 埋め込み式のリストを作成します
            Dim exprs As New List(Of IExpression)()

            ' 埋め込み式の解析を行います
            While iter.HasNext()
                Dim embedded = iter.Current
                Select Case embedded.Kind
                    Case EmbeddedType.None
                        ' 埋込式以外
                        exprs.Add(New PlainTextExpression(embedded.Str))
                        iter.Next()

                    Case EmbeddedType.Unfold
                        ' 展開埋込式
                        Dim inExpr = ParserModule.DirectParse(embedded.Str.Mid(2, embedded.Str.Length - 3))
                        exprs.Add(New UnfoldExpression(inExpr))
                        iter.Next()

                    Case EmbeddedType.NoEscapeUnfold
                        ' 非エスケープ展開埋込式
                        Dim inExpr = ParserModule.DirectParse(embedded.Str.Mid(2, embedded.Str.Length - 3))
                        exprs.Add(New NoEscapeUnfoldExpression(inExpr))
                        iter.Next()

                    Case EmbeddedType.VariableDefine
                        ' 変数定義埋込式
                        exprs.Add(ParseVariableDefineBlock(embedded.Str.Mid(2, embedded.Str.Length - 3)))
                        iter.Next()

                    Case EmbeddedType.IfBlock
                        ' Ifブロック
                        Dim expr = ParseIfStatement(iter)
                        If iter.HasNext() AndAlso iter.Current.Kind = EmbeddedType.EndIfBlock Then
                            exprs.Add(expr)
                            iter.Next()
                        Else
                            Throw New AnalysisException("Ifブロックが閉じられていません。")
                        End If

                    Case EmbeddedType.ElseIfBlock, EmbeddedType.ElseBlock, EmbeddedType.EndIfBlock
                        ' ElseIfブロック、Elseブロック、EndIfブロックはエラー
                        Throw New AnalysisException("Ifブロックが開始されていません。ElseIf、Else、EndIfはIfブロック内でのみ使用できます。")

                    Case EmbeddedType.ForBlock
                        ' Forブロック
                        Dim expr = ParseForStatement(iter)
                        If iter.HasNext() AndAlso iter.Current.Kind = EmbeddedType.EndForBlock Then
                            exprs.Add(expr)
                            iter.Next()
                        Else
                            Throw New AnalysisException("Forブロックが閉じられていません。")
                        End If

                    Case EmbeddedType.SelectBlock
                        ' Selectブロック
                        Dim expr = ParseSelectStatement(iter)
                        If iter.HasNext() AndAlso iter.Current.Kind = EmbeddedType.EndSelectBlock Then
                            exprs.Add(expr)
                            iter.Next()
                        Else
                            Throw New AnalysisException("Selectブロックが閉じられていません。")
                        End If

                    Case EmbeddedType.SelectCaseBlock, EmbeddedType.SelectDefaultBlock, EmbeddedType.EndSelectBlock
                        ' SelectCaseブロック、SelectDefaultブロック、EndSelectブロックはエラー
                        Throw New AnalysisException("Selectブロックが開始されていません。SelectCase、SelectDefault、EndSelectはSelectブロック内でのみ使用できます。")

                    Case EmbeddedType.SetBlock
                        ' Setブロック
                        exprs.Add(ParseSetStatement(embedded.Str))
                        iter.Next()

                    Case EmbeddedType.BrBlock
                        ' Brのブロック
                        exprs.Add(BrExpression.Instance)
                        iter.Next()

                    Case EmbeddedType.VlBrBlock
                        ' 仮想Brのブロック
                        exprs.Add(VlBrExpression.Instance)
                        iter.Next()

                    Case EmbeddedType.TrimBlock
                        ' Trimブロック
                        Dim expr = ParseTrimStatement(iter, embedded.Str)
                        If iter.HasNext() AndAlso iter.Current.Kind = EmbeddedType.EndTrimBlock Then
                            exprs.Add(expr)
                            iter.Next()
                        Else
                            Throw New AnalysisException("Trimブロックが閉じられていません。")
                        End If

                    Case EmbeddedType.EndTrimBlock
                        ' EndTrimのブロック
                        Throw New AnalysisException("Trimブロックが開始されていません。EndTrimはTrimブロック内でのみ使用できます。")

                    Case EmbeddedType.RemoveBlock
                        ' Remブロック
                        Dim expr = ParseRemoveStatement(iter, embedded.Str)
                        If iter.HasNext() AndAlso iter.Current.Kind = EmbeddedType.EndRemoveBlock Then
                            exprs.Add(expr)
                            iter.Next()
                        Else
                            Throw New AnalysisException("Trimブロックが閉じられていません。")
                        End If

                    Case EmbeddedType.EndRemoveBlock
                        ' EndTrimのブロック
                        Throw New AnalysisException("Remブロックが開始されていません。EndRemはRemブロック内でのみ使用できます。")

                    Case EmbeddedType.EmptyBlock
                        ' 空のブロック
                        exprs.Add(EmptyExpression.Instance)
                        iter.Next()
                End Select
            End While

            Return New ListExpression(exprs.ToArray())
        End Function

        ''' <summary>
        ''' 変数定義ブロックを解析します。
        ''' 変数定義ブロックは、変数名とその値を定義するために使用されます。
        ''' 変数名はU8String型で指定され、値はIExpression型で表されます。
        ''' </summary>
        ''' <param name="embeddedText">変数定義ブロック文字列。</param>
        ''' <returns>変数定義式。</returns>
        Friend Function ParseVariableDefineBlock(embeddedText As U8String) As IExpression
            ' 式バッファを生成
            Dim exprs As New List(Of VariableDefineExpression)()

            ' 入力文字列を単語に分割
            Dim words = LexicalModule.SplitWords(embeddedText)

            ' 変数式を解析します
            Dim iter As New ParserIterator(Of LexicalModule.Word)(words)
            While iter.HasNext()
                exprs.Add(ParseVvariable(iter))
                If iter.HasNext() Then
                    If iter.Current.Kind <> WordType.Semicolon Then
                        Throw New AnalysisException("変数定義はセミコロンで区切られている必要があります")
                    End If
                    iter.Next() ' セミコロンをスキップ
                End If
            End While

            ' 解析していない文が残っている場合はエラーを返します
            If iter.HasNext() Then
                Throw New AnalysisException("変数定義ブロックが正しく宣言されていません")
            End If

            ' 変数式をリストとして返します
            Return New VariableDefineListExpression(exprs.ToArray())
        End Function

        ''' <summary>
        ''' Ifステートメントを解析します。
        ''' Ifステートメントは、条件に基づいて異なる処理を実行するために使用されます。
        ''' このメソッドは、Ifブロック、ElseIfブロック、Elseブロック、およびEndIfブロックを解析します。
        ''' </summary>
        ''' <param name="iter">イテレータ。</param>
        ''' <returns>Ifステートメント。</returns>
        Private Function ParseIfStatement(iter As ParserIterator(Of EmbeddedBlock)) As IExpression
            ' 式バッファを生成します
            Dim exprs As New List(Of IExpression)()

            ' 最初の条件を取得
            Dim prevBlock = iter.Next()
            Dim prevType = EmbeddedType.IfBlock

            Dim st = iter.CurrentIndex
            Dim ed = iter.CurrentIndex
            Dim lv = 0
            Dim update = False
            While iter.HasNext()
                Dim stat = iter.Current
                Select Case stat.Kind
                    Case EmbeddedType.IfBlock
                        ' Ifブロックの開始
                        lv += 1

                    Case EmbeddedType.ElseIfBlock
                        ' Else Ifブロックの開始
                        If lv = 0 Then
                            exprs.Add(ParseIfExpression(iter, prevBlock, prevType, st, ed))
                            prevBlock = stat
                            prevType = EmbeddedType.ElseIfBlock
                            update = True
                        End If

                    Case EmbeddedType.ElseBlock
                        ' Elseブロックの処理
                        If lv = 0 Then
                            exprs.Add(ParseIfExpression(iter, prevBlock, prevType, st, ed))
                            prevBlock = stat
                            prevType = EmbeddedType.ElseBlock
                            update = True
                        End If

                    Case EmbeddedType.EndIfBlock
                        ' Ifブロックの終了
                        If lv > 0 Then
                            lv -= 1
                        Else
                            exprs.Add(ParseIfExpression(iter, prevBlock, prevType, st, ed))
                            Exit While ' ネストが終了でループも終了
                        End If

                    Case Else
                        ' 他の埋め込みブロックは無視するか、エラーを投げることも可能ですが、ここでは無視します。
                End Select
                iter.Next()

                If update Then
                    st = iter.CurrentIndex
                    update = False
                End If
                ed = iter.CurrentIndex
            End While

            ' Ifステートメントを解析します
            Return New IfStatementExpression(exprs.ToArray())
        End Function

        ''' <summary>
        ''' Ifブロックの式を解析します。
        ''' このメソッドは、Ifブロック、ElseIfブロック、Elseブロックの条件式を解析し、対応する式を返します。
        ''' </summary>
        ''' <param name="iter">イテレーター。</param>
        ''' <param name="prevBlock">前のIf式。</param>
        ''' <param name="prevType">前のIf式の型。</param>
        ''' <param name="st">開始位置。</param>
        ''' <param name="ed">終了位置。</param>
        ''' <returns>解析した式。</returns>
        Private Function ParseIfExpression(
            iter As ParserIterator(Of EmbeddedBlock),
            prevBlock As EmbeddedBlock,
            prevType As EmbeddedType,
            st As Integer,
            ed As Integer
        ) As IExpression
            ' ブロック内の要素のイテレータを作成します
            Dim inIter = iter.GetRangeIterator(st, ed)

            ' 条件式を解析します
            Select Case prevType
                Case EmbeddedType.IfBlock, EmbeddedType.ElseIfBlock
                    ' IfブロックまたはElseIfブロックの条件式を解析
                    Dim condition = ParserModule.DirectParse(prevBlock.Str)
                    ' Ifブロックの実行部を解析
                    Dim innerExpr = ParseEmbeddedText(inIter)
                    ' If条件式を作成
                    Return New IfExpression(condition, innerExpr)

                Case EmbeddedType.ElseBlock
                    ' Elseブロックの実行部を作成
                    Dim innerExpr = ParseEmbeddedText(inIter)
                    Return New ElseExpression(innerExpr)

                Case Else
                    Throw New AnalysisException("条件式の解析に失敗しました。")
            End Select
        End Function

        ''' <summary>
        ''' Forステートメントを解析します。
        ''' Forステートメントは、繰り返し処理を行うために使用されます。
        ''' </summary>
        ''' <param name="iter">イテレータ。</param>
        ''' <returns>Forステートメント。</returns>
        ''' <remarks>
        ''' このメソッドは、Forブロックの開始から終了までの範囲を解析し、対応する式を返します。
        ''' </remarks>
        Private Function ParseForStatement(iter As ParserIterator(Of EmbeddedBlock)) As IExpression
            ' forブロックの開始を取得
            Dim forCondition = iter.Next()

            ' forの繰り返し範囲を取得
            Dim st = iter.CurrentIndex
            Dim ed = iter.CurrentIndex
            Dim lv = 0
            Dim endFor = False
            While iter.HasNext()
                Select Case iter.Current.Kind
                    Case EmbeddedType.ForBlock
                        ' Forが開始された場合、ネストレベルを増やす
                        lv += 1

                    Case EmbeddedType.EndForBlock
                        ' Forが終了された場合、ネストレベルを減らす
                        If lv > 0 Then
                            lv -= 1
                        Else
                            endFor = True
                            Exit While ' ネストが終了でループも終了
                        End If

                    Case Else
                        ' 他の埋め込みブロックは無視するか、エラーを投げることも可能ですが、ここでは無視します。
                End Select

                iter.Next()
                ed = iter.CurrentIndex
            End While

            If endFor Then
                ' 入力文字列を単語に分割します
                Dim words = LexicalModule.SplitWords(forCondition.Str)

                ' forの繰り返し条件を解析します
                Dim iterWords = New ParserIterator(Of LexicalModule.Word)(words)
                Dim forExpr = ParserModule.ParseForStatement(iterWords)

                ' forの繰り返す範囲を式として解析します
                Dim bodyIter = iter.GetRangeIterator(st, ed)
                Dim bodyExpr = ParseEmbeddedText(bodyIter)

                Return New ForExpression(forExpr.varName, forExpr.collectionExpr, bodyExpr)
            Else
                ' Forブロックが閉じられていない場合はエラーを返す
                Throw New AnalysisException("Forブロックが閉じられていません。")
            End If
        End Function

        ''' <summary>
        ''' Selectステートメントを解析します。
        ''' Selectステートメントは、条件に基づいて異なる処理を実行するために使用されます。
        ''' </summary>
        ''' <param name="iter">イテレータ。</param>
        ''' <returns>Selectステートメント。</returns>
        ''' <remarks>
        ''' このメソッドは、Selectブロックの開始から終了までの範囲を解析し、対応する式を返します。
        ''' </remarks>
        Private Function ParseSelectStatement(iter As ParserIterator(Of EmbeddedBlock)) As IExpression
            ' 式バッファを生成します
            Dim exprs As New List(Of IExpression)()

            ' selectブロックの開始を取得
            Dim prevBlock = iter.Next()
            Dim prevType = EmbeddedType.SelectBlock

            Dim st = iter.CurrentIndex
            Dim ed = iter.CurrentIndex
            Dim lv = 0
            Dim update = False
            While iter.HasNext()
                Dim stat = iter.Current
                Select Case stat.Kind
                    Case EmbeddedType.SelectBlock
                        ' selectブロックの開始
                        lv += 1

                    Case EmbeddedType.SelectCaseBlock
                        ' caseブロックの開始
                        If lv = 0 Then
                            exprs.Add(ParseSelectExpression(iter, prevBlock, prevType, st, ed))
                            prevBlock = stat
                            prevType = EmbeddedType.SelectCaseBlock
                            update = True
                        End If

                    Case EmbeddedType.SelectDefaultBlock
                        ' defaultブロックの開始
                        If lv = 0 Then
                            exprs.Add(ParseSelectExpression(iter, prevBlock, prevType, st, ed))
                            prevBlock = stat
                            prevType = EmbeddedType.SelectDefaultBlock
                            update = True
                        End If

                    Case EmbeddedType.EndSelectBlock
                        ' selectブロックの終了
                        If lv > 0 Then
                            lv -= 1
                        Else
                            exprs.Add(ParseSelectExpression(iter, prevBlock, prevType, st, ed))
                            Exit While ' ネストが終了でループも終了
                        End If

                    Case Else
                        ' 他の埋め込みブロックは無視するか、エラーを投げることも可能ですが、ここでは無視します。
                End Select
                iter.Next()

                If update Then
                    st = iter.CurrentIndex
                    update = False
                End If
                ed = iter.CurrentIndex
            End While

            ' selectステートメントを解析します
            Dim selectExpr = exprs(0)
            exprs.RemoveAt(0)
            Return New SelectStatementExpression(selectExpr, exprs.ToArray())
        End Function

        ''' <summary>
        ''' Selectブロックの式を解析します。
        ''' このメソッドは、SelectブロックまたはCaseブロックの条件式を解析し、対応する式を返します。
        ''' </summary>
        ''' <param name="iter">イテレーター。</param>
        ''' <param name="prevBlock">前のSelectブロック。</param>
        ''' <param name="prevType">前のSelectブロックの型。</param>
        ''' <param name="st">開始位置。</param>
        ''' <param name="ed">終了位置。</param>
        ''' <returns>解析した式。</returns>
        Private Function ParseSelectExpression(iter As ParserIterator(Of EmbeddedBlock), prevBlock As EmbeddedBlock, prevType As EmbeddedType, st As Integer, ed As Integer) As IExpression
            ' ブロック内の要素のイテレータを作成します
            Dim inIter = iter.GetRangeIterator(st, ed)

            Select Case prevType
                Case EmbeddedType.SelectBlock
                    ' SelectブロックまたはCaseブロックの式を解析
                    Dim matchExpr = ParserModule.DirectParse(prevBlock.Str)
                    Return New SelectExpression(matchExpr, ParseEmbeddedText(inIter))

                Case EmbeddedType.SelectCaseBlock
                    ' SelectブロックまたはCaseブロックの式を解析
                    Dim matchExpr = ParserModule.DirectParse(prevBlock.Str)
                    Return New SelectCaseExpression(matchExpr, ParseEmbeddedText(inIter))

                Case EmbeddedType.SelectDefaultBlock
                    ' SelectDefaultブロックの本体部を解析
                    Dim innerExpr = ParseEmbeddedText(inIter)
                    Return New SelectDefaultExpression(innerExpr)

                Case Else
                    Throw New AnalysisException("Selectブロックの解析に失敗しました。")
            End Select
        End Function

        ''' <summary>
        ''' Setステートメントを解析します。
        ''' Setステートメントは、変数の定義を行うために使用されます。
        ''' </summary>
        ''' <param name="embeddedText">Setブロックの文字列。</param>
        ''' <returns>変数定義式。</returns>
        ''' <remarks>
        ''' このメソッドは、Setブロック内の変数定義を解析し、対応する式を返します。
        ''' </remarks>
        Private Function ParseSetStatement(embeddedText As U8String) As IExpression
            ' 式バッファを生成
            Dim exprs As New List(Of VariableDefineExpression)()

            ' 入力文字列を単語に分割します
            Dim words = LexicalModule.SplitWords(embeddedText)

            ' 変数代入式を解析します
            Dim iter As New ParserIterator(Of LexicalModule.Word)(words)
            While iter.HasNext()
                exprs.Add(ParseVvariable(iter))
                If iter.HasNext() Then
                    If iter.Current.Kind <> WordType.Semicolon Then
                        Throw New AnalysisException("変数定義はセミコロンで区切られている必要があります")
                    End If
                    iter.Next() ' セミコロンをスキップ
                End If
            End While

            ' 解析していない文が残っている場合はエラーを返します
            If iter.HasNext() Then
                Throw New AnalysisException("変数代入定義ブロックが正しく宣言されていません")
            End If

            ' 変数式をリストとして返します
            Return New VariableDefineListExpression(exprs.ToArray())
        End Function

        ''' <summary>
        ''' Trimステートメントを解析します。
        ''' Trimステートメントは、文字列の前後の空白を削除するために使用されます。
        ''' </summary>
        ''' <param name="iter">イテレータ。</param>
        ''' <param name="trimCmd">Trimコマンド文字列。</param>
        ''' <returns>Trimステートメント。</returns>
        ''' <remarks>
        ''' このメソッドは、Trimブロック内の文字列を解析し、対応する式を返します。
        ''' </remarks>
        Private Function ParseTrimStatement(iter As ParserIterator(Of EmbeddedBlock), trimCmd As U8String) As IExpression
            iter.Next()

            ' Trimする文字列を解析します
            Dim words = LexicalModule.SplitWords(trimCmd)
            Dim inIter As New ParserIterator(Of LexicalModule.Word)(words)

            Dim exper As New List(Of IExpression)()
            While (inIter.HasNext())
                ' 要素を取得
                exper.Add(ParseTernaryOperator(inIter))

                ' カンマを評価
                If inIter.HasNext() Then
                    Select Case inIter.Current.Kind
                        Case WordType.Comma
                            ' カンマをスキップ
                            inIter.Next()

                        Case Else
                            Throw New AnalysisException("無効な式です。")
                    End Select
                End If
            End While

            ' Trimブロックの開始と終了位置を取得
            Dim st = iter.CurrentIndex
            Dim ed = iter.CurrentIndex
            Dim lv = 0
            While iter.HasNext()
                Dim stat = iter.Current
                Select Case stat.Kind
                    Case EmbeddedType.TrimBlock
                        ' Trimブロックの開始
                        lv += 1

                    Case EmbeddedType.EndTrimBlock
                        ' Trimブロックの終了
                        If lv > 0 Then
                            lv -= 1
                        Else
                            Exit While ' ネストが終了でループも終了
                        End If

                    Case Else
                        ' 他の埋め込みブロックは無視するか、エラーを投げることも可能ですが、ここでは無視します。
                End Select
                iter.Next()

                ed = iter.CurrentIndex
            End While

            ' Trimステートメントを解析します
            ' ブロック内の要素のイテレータを作成します
            Dim contIter = iter.GetRangeIterator(st, ed)
            Return New TrimStatementExpression(exper.ToArray(), ParseEmbeddedText(contIter))
        End Function

        Private Function ParseRemoveStatement(iter As ParserIterator(Of EmbeddedBlock), remCmd As U8String) As IExpression
            iter.Next()

            ' Removeする文字列を解析します
            Dim words = LexicalModule.SplitWords(remCmd)
            Dim inIter As New ParserIterator(Of LexicalModule.Word)(words)

            Dim exper As New List(Of IExpression)()
            While (inIter.HasNext())
                ' 要素を取得
                exper.Add(ParseTernaryOperator(inIter))

                ' カンマを評価
                If inIter.HasNext() Then
                    Select Case inIter.Current.Kind
                        Case WordType.Comma
                            ' カンマをスキップ
                            inIter.Next()

                        Case Else
                            Throw New AnalysisException("無効な式です。")
                    End Select
                End If
            End While

            ' Remブロックの開始と終了位置を取得
            Dim st = iter.CurrentIndex
            Dim ed = iter.CurrentIndex
            Dim lv = 0
            While iter.HasNext()
                Dim stat = iter.Current
                Select Case stat.Kind
                    Case EmbeddedType.RemoveBlock
                        ' Remブロックの開始
                        lv += 1

                    Case EmbeddedType.EndRemoveBlock
                        ' Remブロックの終了
                        If lv > 0 Then
                            lv -= 1
                        Else
                            Exit While ' ネストが終了でループも終了
                        End If

                    Case Else
                        ' 他の埋め込みブロックは無視するか、エラーを投げることも可能ですが、ここでは無視します。
                End Select
                iter.Next()

                ed = iter.CurrentIndex
            End While

            ' Remステートメントを解析します
            ' ブロック内の要素のイテレータを作成します
            Dim contIter = iter.GetRangeIterator(st, ed)
            Return New RemStatementExpression(exper.ToArray(), ParseEmbeddedText(contIter))
        End Function

    End Module

End Namespace
