Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 分析マッチャーインターフェース。
    ''' </summary>
    Public Interface IAnalysisMatcher

        ''' <summary>
        ''' キャッシュをクリアします。
        ''' </summary>
        Sub ClearCache()

        ''' <summary>
        ''' 解答を取得します。
        ''' </summary>
        ''' <returns>解答リスト。</returns>
        Function GetAnswer() As List(Of ABNFAnalysisItem)

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <returns>マッチ結果。</returns>
        Function Match(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, shift As Integer)

        ''' <summary>
        ''' 次の位置へ移動を試みる。
        ''' </summary>
        ''' <param name="tr">位置調整バイト列。</param>
        ''' <param name="env">ABNF環境。</param>
        ''' <returns>移動結果。</returns>
        Function MoveNext(tr As PositionAdjustBytes, env As ABNFEnvironment) As (success As Boolean, shift As Integer)

    End Interface

End Namespace
