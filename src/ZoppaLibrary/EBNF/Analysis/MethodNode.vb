Option Explicit On
Option Strict On

Imports ZoppaLibrary.ABNF
Imports ZoppaLibrary.ABNF.ABNFSyntaxAnalysis
Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' メソッドノード。
    ''' </summary>
    NotInheritable Class MethodNode
        Inherits AnalysisNode

        ''' <summary>メソッドを取得します。</summary>
        Private ReadOnly _method As Func(Of IPositionAdjustReader, Boolean)

        ''' <summary>名前を取得します。</summary>
        Public ReadOnly Property Name As String

        ''' <summary>評価範囲を取得します。</summary>
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
        ''' <param name="id">ID。</param>
        ''' <param name="name">名前。</param>
        ''' <param name="method">マッチ対象を判定する関数。</param>
        Public Sub New(id As Integer, name As String, method As Func(Of IPositionAdjustReader, Boolean))
            MyBase.New(id)
            Me._method = method
            Me.Name = name
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>マッチ結果。</returns>
        Public Overrides Function Match(tr As IPositionAdjustReader, env As EBNFEnvironment, ruleName As String) As (success As Boolean, answer As EBNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()
            Dim startPos = tr.Position

            Dim hit = True

            If Me._method(tr) Then
                ' 成功した場合は真を返す
                Return (True, New EBNFAnalysisItem(Me.Name, New List(Of EBNFAnalysisItem), tr, startPos, tr.Position))
            Else
                ' 失敗情報を設定
                env.SetFailureInformation(ruleName, tr, startPos, Me.Range)

                ' 一致しない場合は偽を返す
                snapPos.Restore()
                Return (False, Nothing)
            End If
        End Function

        ''' <summary>
        ''' 次のパターンのマッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <returns>
        ''' success: マッチが成功した場合にTrue。
        ''' answer: 解析結果アイテム。
        ''' </returns>
        Public Overrides Function MoveNext(tr As IPositionAdjustReader,
                                           env As EBNFEnvironment) As (success As Boolean, answer As EBNFAnalysisItem)
            Return (False, Nothing)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"Method:{Me._Name}"
        End Function

    End Class

End Namespace
