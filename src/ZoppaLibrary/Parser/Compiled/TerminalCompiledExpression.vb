Option Explicit On
Option Strict On

Namespace Parser

    ''' <summary>
    ''' 終端記号にマッチするコンパイル済み式を表します。
    ''' </summary>
    Public NotInheritable Class TerminalCompiledExpression
        Implements ICompiledExpression

        ''' <summary>
        ''' マッチ対象の文字列。
        ''' </summary>
        Private ReadOnly _target As ExpressionRange

        ''' <summary>
        ''' 文字列の値。
        ''' </summary>
        Private ReadOnly _stringValue As String

        ''' <summary>
        ''' 読み取り用バッファ。
        ''' </summary>
        Private ReadOnly _readbuffer As Char()

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="target">マッチ対象の文字列を表す <see cref="ExpressionRange"/>。</param>
        Public Sub New(target As ExpressionRange)
            Me._target = target
            Me._stringValue = target.ToString()
            Me._readbuffer = New Char(Me._stringValue.Length - 1) {}
        End Sub

        ''' <summary>
        ''' 指定された <see cref="IPositionAdjustReader"/> の現在位置にある文字列が
        ''' この式にマッチするかどうかを判定します。
        ''' </summary>
        ''' <param name="tr">入力ソースを表す <see cref="IPositionAdjustReader"/>。</param>
        ''' <param name="ruleTable">ルールテーブル。</param>
        ''' <param name="specialMethods">特殊メソッドのテーブル。</param>
        ''' <param name="answers">解析結果を格納する範囲のリスト。</param>
        ''' <param name="debugMode">デバッグモード。</param>
        ''' <param name="messages">返却メッセージリスト。</param>
        ''' <returns>マッチした場合は true。それ以外は false。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              ruleTable As SortedDictionary(Of String, RuleCompiledExpression),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              answers As List(Of AnalysisRange),
                              debugMode As Boolean,
                              messages As DebugMessage) As Boolean Implements ICompiledExpression.Match
            Dim startPos = tr.Position
            Dim snap = tr.MemoryPosition()

            Dim rcnt = tr.Read(Me._readbuffer, 0, Me._stringValue.Length)
            If EqualString(Me._readbuffer, rcnt, Me._stringValue) Then
                If debugMode Then
                    messages.Add($"一致:{Me._stringValue}")
                End If
                answers.Add(New AnalysisRange("literal", New List(Of AnalysisRange)(), tr, startPos, tr.Position))
                Return True
            Else
                snap.Restore()
                Return False
            End If
        End Function

        ''' <summary>
        ''' 読み取った文字列と指定された文字列が等しいかどうかを判定します。
        ''' </summary>
        ''' <param name="readbuffer">読み取りバッファ。</param>
        ''' <param name="rcnt">読み取り文字数。</param>
        ''' <param name="stringValue">比較対象の文字列。</param>
        ''' <returns>等しい場合は true。それ以外は false。</returns>
        Private Shared Function EqualString(readbuffer() As Char, rcnt As Integer, stringValue As String) As Boolean
            If rcnt <> stringValue.Length Then
                Return False
            End If
            For i As Integer = 0 To stringValue.Length - 1
                If readbuffer(i) <> stringValue.Chars(i) Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' この式を文字列として表現します。
        ''' </summary>
        ''' <returns>この式の文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me._target.ToString()
        End Function

    End Class

End Namespace
