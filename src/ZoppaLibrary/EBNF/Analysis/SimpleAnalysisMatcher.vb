Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace EBNF

    ''' <summary>
    ''' 単純な分析マッチャー。
    ''' </summary>
    Public NotInheritable Class SimpleAnalysisMatcher
        Implements IAnalysisMatcher

        ''' <summary>ルートノード。</summary>
        Private ReadOnly _root As AnalysisNode

        ''' <summary>ルール名。</summary>
        Private ReadOnly _ruleName As String

        ''' <summary>解答リスト。</summary>
        Private ReadOnly _answer As New List(Of EBNFAnalysisItem)()

        ''' <summary>Match が呼び出されたかどうか。</summary>
        Private _matchCalled As Boolean = False

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="root">ルートノード。</param>
        Public Sub New(root As AnalysisNode, ruleName As String)
            Me._root = root
            Me._ruleName = ruleName
        End Sub

        ''' <summary>
        ''' キャッシュをクリアする。
        ''' </summary>
        Public Sub ClearCache() Implements IAnalysisMatcher.ClearCache
            ' キャッシュはないので何もしない
        End Sub

        ''' <summary>
        ''' マッチを試みる。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <returns>解析が成功した場合に True を返します。</returns>
        Public Function Match(tr As IPositionAdjustReader,
                              env As EBNFEnvironment) As (success As Boolean, shift As Integer) Implements IAnalysisMatcher.Match
            Dim startPosition = tr.Position
            Dim node = Me._root
            Me._answer.Clear()
            Me._matchCalled = True

            ' ルートを順次試行、一致を確認
            Do While True
                Dim nextNode = node.Routes(0).NextNode

                ' 対象ノードが一致するか判定
                Dim matched = nextNode.Match(tr, env, Me._ruleName)
                If matched.success Then
                    ' 次のノードへ進む
                    Me._answer.Add(matched.answer)

                    ' 最終ノードに到達した場合は成功
                    If nextNode.Routes.Count = 0 Then
                        Return (True, 0)
                    End If

                    ' 次のノードへ進む
                    node = nextNode
                Else
                    ' ノードが一致しなかった場合は失敗
                    Return (False, 0)
                End If
            Loop
        End Function

        ''' <summary>
        ''' 次の解析ステップを実行する。
        ''' </summary>
        ''' <param name="tr">位置調整リーダー。</param>
        ''' <param name="env">解析環境。</param>
        ''' <returns>次の解析ステップがないため Falseを返す。</returns>
        Public Function MoveNext(tr As IPositionAdjustReader,
                                 env As EBNFEnvironment) As (success As Boolean, shift As Integer) Implements IAnalysisMatcher.MoveNext
            If Me._matchCalled Then
                ' Match が呼び出された後は次のステップはない
                Return (False, 0)
            Else
                ' Match が呼び出されていない場合は Match を実行する
                Return Me.Match(tr, env)
            End If
        End Function

        ''' <summary>
        ''' 解析結果を取得する。
        ''' </summary>
        ''' <returns>解析結果リスト。</returns>
        Public Function GetAnswer() As List(Of EBNFAnalysisItem) Implements IAnalysisMatcher.GetAnswer
            Dim res As New List(Of EBNFAnalysisItem)()
            For Each item In Me._answer
                If item IsNot Nothing Then
                    res.Add(item)
                End If
            Next
            Return res
        End Function

    End Class

End Namespace
