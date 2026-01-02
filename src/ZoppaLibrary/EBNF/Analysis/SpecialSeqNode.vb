Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 特殊シーケンスノード。
    ''' </summary>
    NotInheritable Class SpecialSeqNode
        Inherits AnalysisNode

        ''' <summary>識別子名。</summary>
        Private ReadOnly _name As String

        ''' <summary>評価範囲。</summary>
        Public Overrides ReadOnly Property Range As ExpressionRange

        ''' <summary>
        ''' 再試行可能かを取得する。
        ''' </summary>
        Public Overrides ReadOnly Property IsRetry As Boolean
            Get
                Return False
            End Get
        End Property

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ノードID。</param>
        ''' <param name="range">評価範囲。</param>
        Public Sub New(id As Integer, range As ExpressionRange)
            MyBase.New(id)
            Me._name = range.SubRanges(0).ToString().Trim()
            Me.range = range
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>マッチ結果。</returns>
        Public Overrides Function Match(tr As IPositionAdjustReader, env As EBNFEnvironment, ruleName As String) As (success As Boolean, answer As EBNFAnalysisItem)
            ' 特殊メソッドが存在しない場合はコメント扱いで成功とする
            If Not env.MethodTable.ContainsKey(Me._name) Then
                Return (True, Nothing)
            End If

            Dim snapPos = tr.MemoryPosition()
            Dim startPos = tr.Position

            ' 特殊メソッドを評価
            If env.MethodTable(Me._name)(tr) Then
                Return (True, New EBNFAnalysisItem(Me._name, New List(Of EBNFAnalysisItem), tr, startPos, tr.Position))
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos, Me.Range)

            ' 一致しない場合は偽を返す
            snapPos.Restore()
            Return (False, Nothing)
        End Function

        ''' <summary>
        ''' 次のパターンのマッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">EBNF環境。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public Overrides Function MoveNext(tr As IPositionAdjustReader, env As EBNFEnvironment) As (success As Boolean, answer As EBNFAnalysisItem)
            Return (False, Nothing)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"Special:{Me._name}"
        End Function

    End Class

End Namespace
