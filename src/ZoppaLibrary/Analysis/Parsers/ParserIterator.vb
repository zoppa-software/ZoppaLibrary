Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' パーサー用のイテレーターを表すクラスです。
    ''' このクラスは、解析中の単語、ブロックのイテレーションを管理します。
    ''' </summary>
    ''' <typeparam name="T">イテレーションする要素の型。</typeparam>
    ''' <remarks>
    ''' このクラスは、単語、ブロックリストをイテレートし、次の要素を取得するためのメソッドを提供します。
    ''' </remarks>
    Public NotInheritable Class ParserIterator(Of T)

        ' 要素、ブロックのリストを保持する
        Private ReadOnly _items As T()

        ' 現在のインデックスを保持する
        Private _currentIndex As Integer

        ' 要素数
        Private ReadOnly _itemCount As Integer

        ''' <summary>現在のインデックスを取得します。</summary>
        ''' <returns>現在のインデックス。</returns>
        Public ReadOnly Property CurrentIndex() As Integer
            Get
                Return _currentIndex
            End Get
        End Property

        ''' <summary>現在の要素を取得します。</summary>
        ''' <returns>現在の要素。</returns>
        ''' <remarks>
        ''' 現在のインデックスに対応する要素を返します。
        ''' インデックスが範囲外の場合は、Nothingを返します。
        ''' </remarks>
        Public ReadOnly Property Current As T
            Get
                If _currentIndex < _itemCount Then
                    Return _items(_currentIndex)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="items">イテレーションする要素の配列。</param>
        Public Sub New(items As T())
            Me._items = items
            Me._currentIndex = 0
            Me._itemCount = items.Length
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="items">イテレーションする要素の配列。</param>
        ''' <param name="start">開始位置。</param>
        ''' <param name="count">要素数。</param>
        ''' <remarks>
        ''' 指定された範囲内の要素をイテレートします。
        ''' </remarks>
        Private Sub New(items As T(), start As Integer, count As Integer)
            Me._items = items
            Me._currentIndex = start
            Me._itemCount = count
        End Sub

        ''' <summary>
        ''' 次の要素が存在するかどうかを確認します。
        ''' </summary>
        ''' <returns>次の要素が存在する場合はTrue、それ以外はFalse。</returns>
        ''' <remarks>
        ''' 現在のインデックスがリストの範囲内であるかをチェックします。
        ''' </remarks>
        Public Function HasNext() As Boolean
            Return _currentIndex < _itemCount
        End Function

        ''' <summary>次の要素を取得します。</summary>
        ''' <returns>次の要素。</returns>
        ''' <remarks>
        ''' 現在のインデックスをインクリメントし、次の要素を返します。
        ''' インデックスが範囲外の場合は、Nothingを返します。
        ''' </remarks>
        Public Function [Next]() As T
            If _currentIndex < _itemCount Then
                Dim res = _items(_currentIndex)
                _currentIndex += 1
                Return res
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' 指定された範囲のイテレーターを取得します。
        ''' </summary>
        ''' <param name="startIndex">開始インデックス。</param>
        ''' <param name="endIndex">終了インデックス。</param>
        ''' <returns>指定された範囲のイテレーター。</returns>
        ''' <remarks>
        ''' 指定された範囲内の要素をイテレートする新しいイテレーターを返します。
        ''' </remarks>
        Public Function GetRangeIterator(startIndex As Integer, endIndex As Integer) As ParserIterator(Of T)
            Return New ParserIterator(Of T)(_items, startIndex, endIndex)
        End Function

    End Class

End Namespace
