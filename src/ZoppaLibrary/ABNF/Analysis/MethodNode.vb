Option Explicit On
Option Strict On

Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' メソッドノード。
    ''' </summary>
    NotInheritable Class MethodNode
        Inherits AnalysisNode

        ''' <summary>
        ''' メソッド。
        ''' </summary>
        Private ReadOnly _method As Func(Of PositionAdjustBytes, Boolean)

        ''' <summary>
        ''' 名前を取得します。
        ''' </summary>
        Public ReadOnly Property Name As String

        ''' <summary>
        ''' コンストラクタ。
        ''' </summary>
        ''' <param name="id">ID。</param>
        ''' <param name="name">名前。</param>
        ''' <param name="method">マッチ対象を判定する関数。</param>
        Public Sub New(id As Integer, name As String, method As Func(Of PositionAdjustBytes, Boolean))
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
        Public Overrides Function Match(tr As PositionAdjustBytes, env As ABNFEnvironment, ruleName As String) As (success As Boolean, answer As ABNFAnalysisItem)
            Dim snapPos = tr.MemoryPosition()
            Dim startPos = tr.Position

            Dim hit = True

            If Me._method(tr) Then
                ' 成功した場合は真を返す
                Return (True, New ABNFAnalysisItem(Me.Name, New List(Of ABNFAnalysisItem), tr, startPos, tr.Position))
            Else
                ' 失敗情報を設定
                env.SetFailureInformation(ruleName, tr, startPos, Me.Range)

                ' 一致しない場合は偽を返す
                snapPos.Restore()
                Return (False, Nothing)
            End If
        End Function

    End Class

End Namespace
