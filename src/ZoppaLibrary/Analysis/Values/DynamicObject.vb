Option Strict On
Option Explicit On

Imports ZoppaLibrary.Collections

Namespace Analysis

    ''' <summary>
    ''' 動的オブジェクトを表すクラスです。
    ''' このクラスは、プロパティを動的に追加・取得できる機能を提供します。
    ''' </summary>
    ''' <remarks>
    ''' 動的オブジェクトは、名前と値のペアでプロパティを管理します。
    ''' プロパティは、名前でアクセスでき、存在しない場合は新たに追加されます。
    ''' </remarks>
    Public NotInheritable Class DynamicObject

        Public Const DEFAULT_TEXT_PROPERTY_NAME As String = "XMLValueText"

        ''' <summary>プロパティを表すBtreeコレクション。</summary>
        ''' <remarks>
        ''' Btreeは、プロパティ名でソートされ、効率的な検索と挿入を提供します。
        ''' </remarks>
        Private ReadOnly _properties As Btree(Of PropEntry)

        ''' <summary>
        ''' プロパティが空かどうかを示すプロパティです。
        ''' このプロパティは、動的オブジェクトにプロパティが存在しない場合にTrueを返します。
        ''' プロパティが1つも存在しない場合、IsEmptyはTrueになります。
        ''' </summary>
        ''' <returns>プロパティが空かどうか。</returns>
        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return _properties.Count = 0
            End Get
        End Property

        ''' <summary>
        ''' 指定された名前のプロパティを取得または設定します。
        ''' </summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <returns>プロパティの値。</returns>
        ''' <remarks>
        ''' プロパティが存在しない場合は新たに追加されます。
        ''' </remarks>
        Default Public Property Item(name As String) As Object
            Get
                Dim entry = _properties.Search(New PropEntry(name, Nothing))
                Return entry?.Value
            End Get
            Set(value As Object)
                Dim entry As PropEntry = _properties.Search(New PropEntry(name, Nothing))
                If entry IsNot Nothing Then
                    entry.Value = value
                Else
                    _properties.Insert(New PropEntry(name, value))
                End If
            End Set
        End Property

        ''' <summary>動的オブジェクトのコンストラクタ。</summary>
        Public Sub New()
            _properties = New Btree(Of PropEntry)()
        End Sub

        ''' <summary>
        ''' エントリのイテレータを取得します。
        ''' 
        ''' このメソッドは、プロパティエントリのコレクションを列挙するために使用されます。
        ''' プロパティ名でソートされた順序でエントリを取得できます。
        ''' </summary>
        ''' <returns>エントリのイテレータ。</returns>
        Public Function GetEntries() As IEnumerator(Of PropEntry)
            Return _properties.GetEnumerator()
        End Function

        ''' <summary>
        ''' 動的オブジェクトの文字列表現を取得します。
        ''' 
        ''' このメソッドは、"Text"プロパティが存在する場合、その値を返します。
        ''' "Text"プロパティが存在しない場合は、基底クラスのToStringメソッドを呼び出します。
        ''' </summary>
        ''' <returns>動的オブジェクトの文字列表現。</returns>
        Public Overrides Function ToString() As String
            Dim entry As PropEntry = _properties.Search(New PropEntry(DEFAULT_TEXT_PROPERTY_NAME, Nothing))
            If entry IsNot Nothing Then
                Return entry.Value.ToString()
            Else
                Return MyBase.ToString()
            End If
        End Function

        ''' <summary>
        ''' プロパティエントリを表すクラスです。
        ''' このクラスは、プロパティ名と値を保持し、比較可能なインターフェイスを実装します。
        ''' プロパティ名でソートされるため、Btreeなどのコレクションで使用できます。
        ''' </summary>
        Public NotInheritable Class PropEntry
            Implements IComparable(Of PropEntry)

            ''' <summary>プロパティ名。</summary>
            ''' <remarks>
            ''' プロパティ名は、プロパティを一意に識別するために使用されます。
            ''' </remarks>
            Public ReadOnly Property Name As String

            ''' <summary>プロパティの値。</summary>
            Public Property Value As Object

            ''' <summary>
            ''' プロパティエントリのコンストラクタ。
            ''' </summary>
            ''' <param name="name">プロパティ名。</param>
            ''' <param name="value">プロパティの値。</param>
            ''' <remarks>
            ''' プロパティ名と値を指定して、プロパティエントリを初期化します。
            ''' </remarks>
            Public Sub New(name As String, value As Object)
                Me.Name = name
                Me.Value = value
            End Sub

            ''' <summary>プロパティ名で比較を行います。</summary>
            ''' <param name="other">比較対象のプロパティエントリ。</param>
            ''' <returns>比較結果。名前が同じ場合は0、異なる場合は名前の順序に基づく整数値。</returns>
            Public Function CompareTo(other As PropEntry) As Integer Implements IComparable(Of PropEntry).CompareTo
                Return String.Compare(Me.Name, other.Name, StringComparison.Ordinal)
            End Function

        End Class

    End Class

End Namespace
