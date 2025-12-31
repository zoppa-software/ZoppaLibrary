Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' ε遷移ノード（空遷移）。
    ''' </summary>
    ''' <remarks>
    ''' <para>このノードは、入力を消費せずに常にマッチ成功を返します。</para>
    ''' <para>グラフ構造における空のエッジを表現するために使用されます。</para>
    ''' </remarks>
    Public NotInheritable Class EpsilonNode
        Inherits AnalysisNode

        ''' <summary>評価範囲（常に無効）。</summary>
        Public Overrides ReadOnly Property Range As ExpressionRange
            Get
                Return ExpressionRange.Invalid
            End Get
        End Property

        ''' <summary>
        ''' 再試行可能かを取得する（常にFalse）。
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
        Public Sub New(id As Integer)
            MyBase.New(id)
        End Sub

        ''' <summary>
        ''' マッチを試みる（常に成功）。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <param name="ruleName">ルール名。</param>
        ''' <returns>
        ''' success: 常にTrue。
        ''' answer: 常にNothing（入力を消費しない）。
        ''' </returns>
        Public Overrides Function Match(tr As PositionAdjustBytes,
                                        env As ABNFEnvironment,
                                        ruleName As String) As (success As Boolean, answer As ABNFAnalysisItem)
            ' ε遷移は入力を消費せずに常に成功
            Return (True, Nothing)
        End Function

        ''' <summary>
        ''' 次のパターンのマッチを試みる（常に失敗）。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <returns>
        ''' success: 常にFalse（リトライ不可）。
        ''' answer: 常にNothing。
        ''' </returns>
        Public Overrides Function MoveNext(tr As PositionAdjustBytes,
                                           env As ABNFEnvironment) As (success As Boolean, answer As ABNFAnalysisItem)
            ' ε遷移にはリトライの概念がない
            Return (False, Nothing)
        End Function

    End Class

End Namespace