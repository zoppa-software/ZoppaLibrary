Option Explicit On
Option Strict On

Imports System.CodeDom.Compiler
Imports System.Globalization
Imports System.Text
Imports ZoppaLibrary.BNF

Namespace ABNF

    ''' <summary>
    ''' 文字値ノード（大文字、小文字を区別しない）
    ''' </summary>
    NotInheritable Class CharCaseInsensitiveNode
        Inherits AnalysisNode

        ''' <summary>
        ''' 比較するリテラルバイト列。
        ''' </summary>
        Private ReadOnly _literal As Byte()

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
        ''' <param name="range">式範囲。</param>
        Public Sub New(id As Integer, range As ExpressionRange)
            MyBase.New(id)
            Me._literal = Encoding.UTF8.GetBytes(range.SubRanges(0).ToString())
            Me.Range = range
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

            ' バイト配列を読み込み、リテラルと比較する
            Dim buffer = New Byte(Me._literal.Length - 1) {}
            tr.Read(buffer, 0, buffer.Length)
            Dim unmatched As Integer
            If EqualBytes(Me._literal, buffer, unmatched) Then
                Return (True, New ABNFAnalysisItem("literal", New List(Of ABNFAnalysisItem), tr, startPos, tr.Position))
            End If

            ' 失敗情報を設定
            env.SetFailureInformation(ruleName, tr, startPos + unmatched, Me.Range)

            ' 一致しない場合は偽を返す
            snapPos.Restore()
            Return (False, Nothing)
        End Function

        ''' <summary>
        ''' 2つのバイト配列が等しいかどうかを判定する。
        ''' </summary>
        ''' <param name="literal">バイト配列 A。</param>
        ''' <param name="read">バイト配列 B。</param>
        ''' <param name="unmatched">不一致となった位置のインデックス。</param>
        ''' <returns>等しい場合に True を返します。</returns>
        Private Shared Function EqualBytes(literal As Byte(), read As Byte(), ByRef unmatched As Integer) As Boolean
            If literal.Length <> read.Length Then
                Return False
            End If
            For i As Integer = 0 To literal.Length - 1
                If literal(i) <> read(i) Then
                    unmatched = i
                    Return False
                End If
            Next
            Return True
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
        Public Overrides Function MoveNext(tr As PositionAdjustBytes,
                                           env As ABNFEnvironment) As (success As Boolean, answer As ABNFAnalysisItem)
            Return (False, Nothing)
        End Function

    End Class

End Namespace
