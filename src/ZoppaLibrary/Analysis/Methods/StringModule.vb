Option Strict On
Option Explicit On

Namespace Analysis

    ''' <summary>文字列モジュール。</summary>
    Public Module StringModule

        ''' <summary>文字列をUpperSnakeCaseに変換します。</summary>
        ''' <param name="str">対象の文字列。</param>
        ''' <returns>変換後の文字列。</returns>
        Function ChangeUpperSnakeCase(str As IValue) As IValue
            Dim s = str.Str().ToString().ToUpper().Replace("-", "_").Replace(" ", "_")
            Return s.ToStringValue()
        End Function

        ''' <summary>文字列をSnakeCaseに変換します。</summary>
        ''' <param name="str">対象の文字列。</param>
        ''' <returns>変換後の文字列。</returns>
        Function ChangeSnakeCase(str As IValue) As IValue
            Dim s = str.Str().ToString().ToLower().Replace("-", "_").Replace(" ", "_")
            Return s.ToStringValue()
        End Function

        ''' <summary>文字列の先頭文字を小文字に変換します。</summary>
        ''' <param name="str">対象の文字列。</param>
        ''' <returns>変換後の文字列。</returns>
        Function ChangeFirstCharLower(str As IValue) As IValue
            Dim s = str.Str().ToString()
            If s.Length > 0 Then
                s = Char.ToLower(s(0)) & s.Substring(1)
            End If
            Return s.ToStringValue()
        End Function

        ''' <summary>文字列の先頭文字を大文字に変換します。</summary>
        ''' <param name="str">対象の文字列。</param>
        ''' <returns>変換後の文字列。</returns>
        Function ChangeFirstCharUpper(str As IValue) As IValue
            Dim s = str.Str().ToString()
            If s.Length > 0 Then
                s = Char.ToUpper(s(0)) & s.Substring(1)
            End If
            Return s.ToStringValue()
        End Function

        ''' <summary>日付を指定されたフォーマットで文字列に変換します。</summary>
        ''' <param name="dt">対象日付。</param>
        ''' <param name="format">書式。</param>
        ''' <returns>変換後の文字列。</returns>
        Function FormatDate(dt As IValue, format As IValue) As IValue
            Dim d = dt.ToDate()
            Dim f = format.Str().ToString()
            Return d.ToString(f).ToStringValue()
        End Function

    End Module

End Namespace
