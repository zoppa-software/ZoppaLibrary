Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>
    ''' 埋め込み式の種類を定義する列挙型です。
    ''' この列挙型は、埋め込み式のタイプを表します。
    ''' </summary>
    Public Enum EmbeddedType

        ''' <summary>非埋込。</summary>
        None

        ''' <summary>埋込展開。</summary>
        Unfold

        ''' <summary>エスケープ無し展開。</summary>
        NoEscapeUnfold

        ''' <summary>変数定義。</summary>
        VariableDefine

        ''' <summary>ifブロック。</summary>
        IfStatement

        ''' <summary>ifブロック。</summary>
        IfBlock

        ''' <summary>else ifブロック。</summary>
        ElseIfBlock

        ''' <summary>elseブロック。</summary>
        ElseBlock

        ''' <summary>end ifブロック。</summary>
        EndIfBlock

        ''' <summary>forブロック。</summary>
        ForBlock

        ''' <summary>end forブロック。</summary>
        EndForBlock

        ''' <summary>selectブロック。</summary>
        SelectBlock

        ''' <summary>select caseブロック。</summary>
        SelectCaseBlock

        ''' <summary>select defaultブロック。</summary>
        SelectDefaultBlock

        ''' <summary>end selectブロック。</summary>
        EndSelectBlock

        ''' <summary>setブロック。</summary>
        SetBlock

        ''' <summary>brブロック。</summary>
        BrBlock

        ''' <summary>仮想brブロック。</summary>
        VlBrBlock

        ''' <summary>trimブロック。</summary>
        TrimBlock

        ''' <summary>end trimブロック。</summary>
        EndTrimBlock

        ''' <summary>remブロック。</summary>
        RemoveBlock

        ''' <summary>end remブロック。</summary>
        EndRemoveBlock

        ''' <summary>空ブロック。</summary>
        EmptyBlock

    End Enum

End Namespace
