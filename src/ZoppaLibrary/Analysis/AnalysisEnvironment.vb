Option Strict On
Option Explicit On

Imports System.Reflection
Imports ZoppaLibrary.Collections
Imports ZoppaLibrary.Strings

Namespace Analysis

    ''' <summary>
    ''' 変数の環境を表すクラスです。
    ''' 変数の定義や管理を行います。
    ''' </summary>
    ''' <remarks>
    ''' このクラスは、変数のスコープや値の管理を行うために使用されます。
    ''' </remarks>
    Public NotInheritable Class AnalysisEnvironment

        ' 変数階層
        Private ReadOnly _hierarchy As Variables

        ' 関数リスト
        Private ReadOnly _functions As Btree(Of FuncEntry)

        ' プロパティリスト
        Private ReadOnly _properties As Dictionary(Of Type, Btree(Of PropEntry))

        ''' <summary>コンストラクタ。</summary>
        Public Sub New()
            Me._hierarchy = New Variables()
            Me._functions = New Btree(Of FuncEntry)()
            Me._functions.Insert(
                New FuncEntry(U8String.NewString("now"), Nothing, GetType(AnalysisEnvironment).GetMethod("Now", BindingFlags.NonPublic Or BindingFlags.Static))
            )
            Me._properties = New Dictionary(Of Type, Btree(Of PropEntry))()
        End Sub

        ''' <summary>現在の日時を文字列として取得します。</summary>
        ''' <returns>現在の日時を表す文字列。</returns>
        ''' <remarks>
        ''' このメソッドは、"yyyy/MM/dd HH:mm:ss"形式で現在の日時を返します。
        ''' </remarks>
        Private Shared Function Now() As IValue
            Return New DateTimeValue(DateTime.Now)
        End Function

        ''' <summary>指定したキーと値の変数を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="value">登録する変数の値。</param>
        Public Sub Register(key As U8String, value As IVariable)
            Me._hierarchy.Register(key, value)
        End Sub

        ''' <summary>指定したキーと値の式を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="value">登録する式。</param>
        Sub RegisterExpr(key As U8String, value As IExpression)
            Me._hierarchy.Register(key, New ExprVariable(value))
        End Sub

        ''' <summary>指定したキーと値の式を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="command">登録する式。</param>
        Public Sub RegisterExpr(key As String, command As String)
            Dim expr As IExpression = ParserModule.Translate(command).Expression
            Me._hierarchy.Register(U8String.NewString(key), New ExprVariable(expr))
        End Sub

        Public Sub RegisterExpr(key As String, value As IExpression)
            Me._hierarchy.Register(U8String.NewString(key), New ExprVariable(value))
        End Sub

        ''' <summary>指定したキーと値の数値変数を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="value">登録する変数の値。</param>
        Public Sub RegisterNumber(key As String, value As Double)
            Me._hierarchy.Register(U8String.NewString(key), New NumberVariable(value))
        End Sub

        ''' <summary>指定したキーと値の文字列変数を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="value">登録する変数の値。</param>
        Public Sub RegisterStr(key As String, value As String)
            Me._hierarchy.Register(U8String.NewString(key), New StringVariable(U8String.NewString(value)))
        End Sub

        ''' <summary>指定したキーと値の真偽値変数を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="value">登録する変数の値。</param>
        Public Sub RegisterBool(key As String, value As Boolean)
            Me._hierarchy.Register(U8String.NewString(key), If(value, BooleanVariable.TrueValue, BooleanVariable.FalseValue))
        End Sub

        ''' <summary>指定したキーと値の配列を登録します。</summary>
        ''' <typeparam name="T">配列の値。</typeparam>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="value">登録する配列。</param>
        Public Sub RegisterArray(Of T)(key As String, ParamArray value As T())
            ' 配列の値をIExpression型に変換
            Dim exprArray = New IExpression(value.Length - 1) {}
            For i As Integer = 0 To value.Length - 1
                exprArray(i) = ConvertToExpression(value(i))
            Next
            Me._hierarchy.Register(U8String.NewString(key), New ArrayVariable(exprArray))
        End Sub

        ''' <summary>指定したキーと値のオブジェクト変数を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="obj">登録するオブジェクト。</param>
        Public Sub RegisterObject(key As String, obj As Object)
            Me._hierarchy.Register(U8String.NewString(key), New ObjectVariable(obj))
        End Sub

        ''' <summary>指定したキーと値の日時変数を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="value">登録する日時。</param>
        Public Sub RegisterDateTime(key As String, value As DateTime)
            Me._hierarchy.Register(U8String.NewString(key), New DateTimeVariable(value))
        End Sub

        ''' <summary>指定したキーと値の時間間隔変数を登録します。</summary>
        ''' <param name="key">登録する変数のキー。</param>
        ''' <param name="value">登録する時間間隔。</param>
        Public Sub RegisterTimeSpan(key As String, value As TimeSpan)
            Me._hierarchy.Register(U8String.NewString(key), New TimeSpanVariable(value))
        End Sub

        ''' <summary>
        ''' 指定したキーの変数を取得します。
        ''' </summary>
        ''' <param name="key">取得する変数のキー。</param>
        ''' <returns>指定されたキーの変数。</returns>
        ''' <exception cref="KeyNotFoundException">指定されたキーの変数が存在しない場合にスローされます。</exception>
        ''' <remarks>
        ''' このメソッドは、現在の階層から指定されたキーの変数を検索します。
        ''' もし見つからない場合は、例外をスローします。
        ''' </remarks>
        Function [Get](key As U8String) As IVariable
            Dim result = Me._hierarchy.Get(key)
            If result Is Nothing Then
                Throw New KeyNotFoundException($"変数 '{key}' は存在しません")
            End If
            Return result
        End Function

        ''' <summary>
        ''' 指定したキーの変数を取得します。
        ''' </summary>
        ''' <param name="key">取得する変数のキー。</param>
        ''' <returns>指定されたキーの変数。</returns>
        ''' <exception cref="KeyNotFoundException">指定されたキーの変数が存在しない場合にスローされます。</exception>
        ''' <remarks>
        ''' このメソッドは、現在の階層から指定されたキーの変数を検索します。
        ''' もし見つからない場合は、例外をスローします。
        ''' </remarks>
        Public Function [Get](key As String) As IVariable
            Return Me.Get(U8String.NewString(key))
        End Function

        ''' <summary>指定したキーの変数が存在するかどうかを確認します。</summary>
        ''' <param name="key">確認する変数のキー。</param>
        ''' <returns>指定されたキーが存在する場合はTrue、存在しない場合はFalse。</returns>
        Public Function Contains(key As U8String) As Boolean
            Return Me._hierarchy.Contains(key)
        End Function

        ''' <summary>
        ''' 指定したキーの変数を削除します。
        ''' </summary>
        ''' <param name="key">削除する変数のキー。</param>
        ''' <remarks>
        ''' このメソッドは、現在の階層から指定されたキーの変数を削除します。
        ''' </remarks>
        Public Sub Unregister(key As String)
            Me._hierarchy.Unregister(U8String.NewString(key))
        End Sub

        ''' <summary>
        ''' 新しいスコープを開始します。
        ''' 現在のスコープをスタックにプッシュし、新しいスコープを作成します。
        ''' </summary>
        ''' <returns></returns>
        Public Function GetScope() As Variables.Scope
            Return _hierarchy.GetScope()
        End Function

        ''' <summary>
        ''' 関数の登録を行います。
        ''' このメソッドは、関数名と関数の実行方法を登録します。
        ''' 登録された関数は、後で呼び出すことができます。
        ''' </summary>
        ''' <param name="name">関数名。</param>
        ''' <param name="func">関数の実行方法。</param>
        Public Sub AddFunction(name As U8String, func As Func(Of IValue))
            Dim entry As New FuncEntry(name, func.Target, func.Method)
            Me._functions.Insert(entry)
        End Sub

        ''' <summary>
        ''' 関数の登録を行います。
        ''' このメソッドは、関数名と関数の実行方法を登録します。
        ''' 登録された関数は、後で呼び出すことができます。
        ''' </summary>
        ''' <param name="name">関数名。</param>
        ''' <param name="func">関数の実行方法。</param>
        Public Sub AddFunction(name As U8String, func As Func(Of IValue, IValue))
            Dim entry As New FuncEntry(name, func.Target, func.Method)
            Me._functions.Insert(entry)
        End Sub

        ''' <summary>
        ''' 関数の登録を行います。
        ''' このメソッドは、関数名と関数の実行方法を登録します。
        ''' 登録された関数は、後で呼び出すことができます。
        ''' </summary>
        ''' <param name="name">関数名。</param>
        ''' <param name="func">関数の実行方法。</param>
        Public Sub AddFunction(name As U8String, func As Func(Of IValue, IValue, IValue))
            Dim entry As New FuncEntry(name, func.Target, func.Method)
            Me._functions.Insert(entry)
        End Sub

        ''' <summary>
        ''' 関数の登録を行います。
        ''' このメソッドは、関数名と関数の実行方法を登録します。
        ''' 登録された関数は、後で呼び出すことができます。
        ''' </summary>
        ''' <param name="name">関数名。</param>
        ''' <param name="func">関数の実行方法。</param>
        Public Sub AddFunction(name As U8String, func As Func(Of IValue(), IValue))
            Dim entry As New FuncEntry(name, func.Target, func.Method)
            Me._functions.Insert(entry)
        End Sub

        ''' <summary>
        ''' 関数を呼び出します。
        ''' 指定された関数名と引数で関数を実行し、結果を返します。
        ''' </summary>
        ''' <param name="name">呼び出す関数名。</param>
        ''' <param name="parameter">関数に渡す引数の配列。</param>
        ''' <returns>関数の実行結果。</returns>
        ''' <remarks>
        ''' このメソッドは、登録された関数を呼び出し、結果を返します。
        ''' </remarks>
        Friend Function CallFunction(name As U8String, parameter() As IValue) As IValue
            Dim entry = Me._functions.Search(New FuncEntry(name, Nothing, Nothing))
            Return CType(entry.callfunc?.Invoke(entry.callins, parameter), IValue)
        End Function

        ''' <summary>
        ''' オブジェクトのプロパティを登録します。
        ''' 指定されたオブジェクトのプロパティを取得し、変数階層に登録します。
        ''' </summary>
        ''' <param name="params">プロパティを取得するオブジェクト。</param>
        ''' <remarks>
        ''' このメソッドは、オブジェクトのプロパティを変数階層に登録します。
        ''' プロパティ名はU8Stringとして保存されます。
        ''' </remarks>
        Public Sub RegisterReflectObject(params As Object)
            If params Is Nothing Then
                Throw New ArgumentNullException(NameOf(params))
            End If

            ' オブジェクトの型を取得
            Dim typeKey = params.GetType()

            ' プロパティのB木が存在しない場合は新規作成
            If Not _properties.ContainsKey(typeKey) Then
                ' プロパティのB木を作成
                Dim propTree As New Btree(Of PropEntry)()
                _properties.Add(typeKey, propTree)

                ' プロパティを取得してB木に登録
                For Each prop In typeKey.GetProperties()
                    If prop.CanRead Then
                        Dim entry As New PropEntry(U8String.NewString(prop.Name), prop)
                        propTree.Insert(entry)
                    End If
                Next
            End If

            ' プロパティを変数階層に登録
            For Each prop In _properties(typeKey)
                Dim value = prop.CallProp.GetValue(params, Nothing)
                Me._hierarchy.Register(prop.PropName, ConvertToVariable(value))
            Next
        End Sub

        ''' <summary>関数エントリ。</summary>
        Private Structure FuncEntry
            Implements IComparable(Of FuncEntry)

            ''' <summary>関数名。</summary>
            Public ReadOnly name As U8String

            ''' <summary>関数のインスタンス。</summary>
            Public ReadOnly callins As Object

            ''' <summary>実行する関数。</summary>
            Public ReadOnly callfunc As MethodInfo

            ''' <summary>関数エントリのコンストラクタ。</summary>
            ''' <param name="name">関数名。</param>
            ''' <param name="callins">関数のインスタンス。</param>
            ''' <param name="callfunc">実行する関数。</param>
            Public Sub New(name As U8String, callins As Object, callfunc As MethodInfo)
                Me.name = name
                Me.callins = callins
                Me.callfunc = callfunc
            End Sub

            ''' <summary>
            ''' 関数エントリを比較します。
            ''' このメソッドは、関数名を基準にしてエントリを比較します。
            ''' </summary>
            ''' <param name="other">比較対象。</param>
            ''' <returns>比較結果。</returns>
            Public Function CompareTo(other As FuncEntry) As Integer Implements IComparable(Of FuncEntry).CompareTo
                Return name.CompareTo(other.name)
            End Function

        End Structure

        ''' <summary>プロパティエントリ。</summary>
        ''' <remarks>
        ''' プロパティ名とプロパティ情報を保持します。
        ''' </remarks>
        Private Structure PropEntry
            Implements IComparable(Of PropEntry)

            ''' <summary>プロパティ名を取得。</summary>
            Public ReadOnly Property PropName As U8String

            ''' <summary>プロパティ参照を取得。</summary>
            Public ReadOnly Property CallProp As PropertyInfo

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="propname">プロパティ名。</param>
            ''' <param name="callprop">プロパティ参照。</param>
            Public Sub New(propname As U8String, callprop As PropertyInfo)
                Me.propname = propname
                Me.callprop = callprop
            End Sub

            ''' <summary>
            ''' プロパティエントリを比較します。
            ''' このメソッドは、プロパティ名を基準にしてエントリを比較します。
            ''' </summary>
            ''' <param name="other">比較対象。</param>
            ''' <returns>比較結果。</returns>
            Public Function CompareTo(other As PropEntry) As Integer Implements IComparable(Of PropEntry).CompareTo
                Return propname.CompareTo(other.propname)
            End Function
        End Structure

    End Class

End Namespace
