Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 端点解析（完了）を表します。
    ''' </summary>
    Public NotInheritable Class CompletedAnalysis
        Implements IAnalysis

        ''' <summary>シングルトンインスタンス。</summary>
        Private Shared _singleton As New Lazy(Of CompletedAnalysis)(
            Function() New CompletedAnalysis()
        )

        ''' <summary>
        ''' シングルトンインスタンスを取得する。
        ''' </summary>
        ''' <returns>シングルトンインスタンス。</returns>
        Public Shared Function Instance() As CompletedAnalysis
            Return _singleton.Value
        End Function

        ' 空のパターンリストを共有
        Private _empty As New List(Of IAnalysis)()

        ''' <summary>
        ''' 解析パターンを取得する。
        ''' </summary>
        ''' <returns>解析パターン。</returns>
        Public ReadOnly Property Pattern As List(Of IAnalysis) Implements IAnalysis.Pattern
            Get
                Return _empty
            End Get
        End Property

        ''' <summary>
        ''' 解析を実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <param name="ruleTable">ルール解析テーブル。</param>
        ''' <param name="specialMethods">特殊メソッドテーブル。</param>
        ''' <param name="ruleName">現在のルール名。</param>
        ''' <param name="answers">解析結果のリスト。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              env As EBNFEnvironment,
                              ruleTable As SortedDictionary(Of String, OldRuleAnalysis),
                              specialMethods As SortedDictionary(Of String, Func(Of IPositionAdjustReader, Boolean)),
                              ruleName As String,
                              answers As List(Of EBNFAnalysisItem)) As (sccess As Boolean, shift As Integer) Implements IAnalysis.Match
            Return (True, 0)
        End Function

        ''' <summary>
        ''' 文字列表現を取得する。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "<終了>"
        End Function

    End Class

End Namespace
