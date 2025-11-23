Option Strict On
Option Explicit On

Imports ZoppaLibrary.Collections
Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 変数のコレクションを表すクラスです。
    ''' このクラスは、変数の名前と値を格納し、名前で検索することができます。
    ''' </summary>
    Public NotInheritable Class Variables

        ' 変数リスト
        Private ReadOnly _variables As BPlusTree(Of Entry)

        ' スコープ
        Private ReadOnly _scope As New Stack(Of Scope)()

        ''' <summary>
        ''' コンストラクタ。
        ''' 変数のコレクションを空のB木として初期化します。
        ''' </summary>
        Public Sub New()
            Me._variables = New BPlusTree(Of Entry)()
        End Sub

        ''' <summary>
        ''' 変数を登録します。
        ''' 既に同じ名前の変数が存在する場合は、値を更新します。
        ''' </summary>
        ''' <param name="name">登録する変数名。</param>
        Public Sub Register(name As U8String, value As IVariable)
            ' エントリキーを作成
            Dim entry As New Entry(name, value)

            If _scope.Count > 0 AndAlso Not _scope.Peek()._names.Contains(name) Then
                ' 現在のスコープに変数名が存在しない場合、スコープに追加します。
                _scope.Peek()._names.Insert(name)
                Me._variables.Insert(entry)
            Else
                ' 変数を更新します。
                Dim serd As Entry = Me._variables.Search(entry)
                If serd Is Nothing Then
                    Me._variables.Insert(entry)
                Else
                    serd.Value = value
                End If
            End If
        End Sub

        ''' <summary>
        ''' 変数を登録します。
        ''' 既に同じ名前の変数が存在する場合は、値を更新します。
        ''' </summary>
        ''' <param name="name">登録する変数名。</param>
        ''' <param name="value">登録する変数の値。</param>
        ''' <exception cref="ArgumentNullException">valueがnullの場合にスローされます。</exception>
        Public Sub Register(name As String, value As IVariable)
            Me.Register(U8String.NewString(name), value)
        End Sub

        ''' <summary>
        ''' 変数を取得します。
        ''' 指定した名前の変数が存在しない場合は例外をスローします。
        ''' </summary>
        ''' <param name="name">変数名。</param>
        ''' <returns>指定した名前の変数。</returns>
        ''' <exception cref="KeyNotFoundException">指定した名前の変数がない。</exception>
        Public Function [Get](name As U8String) As IVariable
            Dim entry As Entry = Me._variables.Search(New Entry(name, Nothing))
            If entry IsNot Nothing Then
                Return entry.Value
            Else
                Throw New KeyNotFoundException($"変数 '{name}' は登録されていません。")
            End If
        End Function

        ''' <summary>
        ''' 変数を取得します。
        ''' 指定した名前の変数が存在しない場合は例外をスローします。
        ''' </summary>
        ''' <param name="name">変数名。</param>
        ''' <returns>指定した名前の変数。</returns>
        ''' <exception cref="KeyNotFoundException">指定した名前の変数がない。</exception>
        Public Function [Get](name As String) As IVariable
            Return Me.[Get](U8String.NewString(name))
        End Function

        ''' <summary>
        ''' 変数を登録解除します。
        ''' 指定した名前の変数が存在する場合は削除します。
        ''' </summary>
        ''' <param name="name">登録解除する変数名。</param>
        Public Sub Unregister(name As U8String)
            ' 現在のスコープに変数名が存在する場合、スコープから削除します。
            If _scope.Count > 0 AndAlso _scope.Peek()._names.Contains(name) Then
                _scope.Peek()._names.Remove(name)
            End If

            ' 変数を検索して削除します。
            Dim entry As Entry = Me._variables.Search(New Entry(name, Nothing))
            If entry IsNot Nothing Then
                Me._variables.Remove(entry)
            End If
        End Sub

        ''' <summary>
        ''' 変数を登録解除します。
        ''' 指定した名前の変数が存在する場合は削除します。
        ''' </summary>
        ''' <param name="name">登録解除する変数名。</param>
        Public Sub Unregister(name As String)
            Me.Unregister(U8String.NewString(name))
        End Sub

        ''' <summary>
        ''' 指定した名前の変数が存在するかどうかを確認します。
        ''' </summary>
        ''' <param name="key">変数名。</param>
        ''' <returns>変数が存在する場合はTrue、存在しない場合はFalse。</returns>
        ''' <remarks>
        ''' 変数が現在のスコープまたはグローバルスコープに存在するかどうかを確認します。
        ''' </remark
        Public Function Contains(key As U8String) As Boolean
            Return Me._variables.Contains(New Entry(key, Nothing))
        End Function

        ''' <summary>
        ''' 新しいスコープを開始します。
        ''' 現在のスコープをスタックにプッシュし、新しいスコープを作成します。
        ''' </summary>
        ''' <returns></returns>
        Public Function GetScope() As Scope
            Dim scp As New Scope(Me)
            _scope.Push(scp)
            Return scp
        End Function

        ''' <summary>変数エントリ。</summary>
        Private NotInheritable Class Entry
            Implements IComparable(Of Entry)

            ' 変数の名前
            Public ReadOnly Name As U8String

            ' 変数の値
            Public Value As IVariable

            '''' <summary>
            ''' コンストラクタ。
            ''' 変数の名前と値を指定してエントリを初期化します。
            ''' </summary>
            ''' <param name="name">変数の名前。</param>
            ''' <param name="value">変数の値。</param>
            Public Sub New(name As U8String, value As IVariable)
                Me.Name = name
                Me.Value = value
            End Sub

            ''' <summary>
            ''' 変数の名前でエントリを比較します。
            ''' </summary>
            ''' <param name="other">比較対象のエントリ。</param>
            ''' <returns>比較結果。名前が同じ場合は0、異なる場合は名前の辞書順で比較した結果。</returns>
            Public Function CompareTo(other As Entry) As Integer Implements IComparable(Of Entry).CompareTo
                Return Me.Name.CompareTo(other.Name)
            End Function

        End Class

        ''' <summary>
        ''' 変数のスコープを表すクラスです。
        ''' このクラスは、スコープ内での変数の管理を行います。
        ''' </summary>
        Public NotInheritable Class Scope
            Implements IDisposable

            ' 変数のコレクション
            Private ReadOnly _variables As Variables

            ' スコープ内の変数名のコレクション
            Public _names As New Btree(Of U8String)()

            ''' <summary>
            ''' コンストラクタ。
            ''' 変数のコレクションを指定してスコープを初期化します。
            ''' </summary>
            ''' <param name="vars">変数のコレクション。</param>
            Public Sub New(vars As Variables)
                Me._variables = vars
            End Sub

            ''' <summary>スコープを終了し、変数を登録解除します。</summary>
            Public Sub Dispose() Implements IDisposable.Dispose
                Me._variables._scope.Pop()
                For Each name In _names
                    Me._variables.Unregister(name)
                Next
            End Sub

        End Class

    End Class

End Namespace
