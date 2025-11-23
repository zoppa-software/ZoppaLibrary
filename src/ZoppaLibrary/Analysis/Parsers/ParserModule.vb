Option Strict On
Option Explicit On

Imports ZoppaLibrary.Analysis.LexicalModule
Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' パーサーモジュールを定義するモジュールです。
    ''' このモジュールは、式の解析や変数の管理など、解析に関連する機能を提供します。
    ''' </summary>
    Public Module ParserModule

        ''' <summary>文字列を解析し、結果を取得します。</summary>
        ''' <param name="input">解析する文字列。</param>
        ''' <returns>解析結果。</returns>
        ''' <remarks>
        ''' このメソッドは、指定された文字列を解析し、結果を返します。
        ''' </remarks>
        Public Function Parse(input As String) As AnalysisResults
            Return Parse(U8String.NewString(input))
        End Function

        ''' <summary>文字列を解析し、結果を取得します。</summary>
        ''' <param name="input">解析する文字列。</param>
        ''' <returns>解析結果。</returns>
        ''' <remarks>
        ''' このメソッドは、指定された文字列を解析し、結果を返します。
        ''' </remarks>
        Public Function Parse(input As U8String) As AnalysisResults
            ' 入力文字列をクローン
            Dim newInput = input.NewAllocate()

            ' 式を解析します
            Return New AnalysisResults(newInput, DirectParse(newInput))
        End Function

        ''' <summary>文字列を解析し、結果を取得します。</summary>
        ''' <param name="input">解析する文字列。</param>
        ''' <returns>解析結果の式。</returns>
        ''' <remarks>
        ''' このメソッドは、指定された文字列を解析し、結果の式を返します。
        ''' </remarks>
        Function DirectParse(input As U8String) As IExpression
            ' 入力文字列を単語に分割します
            Dim words = input.SplitWords()

            ' 単語のイテレーターを作成します
            Dim iter As New ParserIterator(Of LexicalModule.Word)(words)

            ' 式を解析します
            Dim exper = ParseTernaryOperator(iter)
            If iter.HasNext() Then
                Throw New AnalysisException("式の解析に失敗しました。")
            End If
            Return exper
        End Function

        ''' <summary>埋め込みテキストを解析します。</summary>
        ''' <param name="input">解析する埋め込みテキスト。</param>
        ''' <returns>解析結果。</returns>
        ''' <remarks>
        ''' このメソッドは、埋め込みテキストを解析し、結果を返します。
        ''' </remarks>
        Public Function Translate(input As U8String) As AnalysisResults
            Return DirectTranslate(input.NewAllocate())
        End Function

        ''' <summary>埋め込みテキストを解析します。</summary>
        ''' <param name="input">解析する埋め込みテキスト。</param>
        ''' <returns>解析結果。</returns>
        ''' <remarks>
        ''' このメソッドは、埋め込みテキストを解析し、結果を返します。
        ''' </remarks>
        Public Function Translate(input As String) As AnalysisResults
            Return DirectTranslate(U8String.NewString(input))
        End Function

        ''' <summary>埋め込みテキストを解析します。</summary>
        ''' <param name="input">解析する文字列。</param>
        ''' <returns>解析結果。</returns>
        ''' <remarks>
        ''' このメソッドは、埋め込みテキストを解析し、結果を返します。
        ''' </remarks>
        Private Function DirectTranslate(input As U8String) As AnalysisResults
            ' 入力文字列を埋込ブロックに分割します
            Dim blocks = input.SplitEmbeddedText()

            ' 単語のイテレーターを作成します
            Dim iter As New ParserIterator(Of LexicalEmbeddedModule.EmbeddedBlock)(blocks)

            ' 式を解析します
            Dim exper = ParseEmbeddedText(iter)
            If iter.HasNext() Then
                Throw New AnalysisException("埋込式の解析に失敗しました。")
            End If
            Return New AnalysisResults(input, exper)
        End Function

        ''' <summary>変数定義ブロックを解析します。</summary>
        ''' <param name="input">解析する文字列。</param>
        ''' <param name="venv">変数定義を登録する解析環境。</param>
        ''' <returns>解析結果の式。</returns>
        ''' <remarks>
        ''' このメソッドは、変数定義ブロックを解析し、解析環境に登録します。
        ''' </remarks>
        Public Sub TranslateVariablesToRegist(input As String, venv As AnalysisEnvironment)
            Dim ans = ParseVariableDefineBlock(U8String.NewString(input))
            ans.GetValue(venv)
        End Sub

    End Module

End Namespace
