Namespace FsUtil

    ''' <summary>
    ''' テキスト処理に関する機能を提供します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TextUtil

#Region "Text"

        ''' <summary>
        ''' 文字列を指定した回数つなぎ合わせる。
        ''' </summary>
        ''' <param name="value">文字列</param>
        ''' <param name="count">回数</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RepeatString(ByVal value As String, ByVal count As Integer)

            Dim result As String = Nothing
            Dim i As Integer

            For i = 0 To count - 1
                result &= value
            Next

            Return result

        End Function

        ''' <summary>
        ''' 改行。
        ''' </summary>
        ''' <param name="count">改行回数</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NewLine(Optional ByVal count As Integer = 1) As String

            Return RepeatString(vbNewLine, count)

        End Function

        ''' <summary>
        ''' 文字列を追加先文字列に改行して追加する。
        ''' </summary>
        ''' <param name="addValue">追加文字列</param>
        ''' <param name="srcValue">追加先文字列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddStringWithNewLine(ByVal addValue As String, ByVal srcValue As String) As String

            Dim result As String = srcValue

            If Not result = Nothing Then result &= vbNewLine
            result &= addValue

            Return result

        End Function

        ''' <summary>
        ''' 文字列を区切り文字と共に追加する。
        ''' </summary>
        ''' <param name="addValue">追加文字列</param>
        ''' <param name="srcValue">追加先文字列</param>
        ''' <param name="delimiter">区切り文字</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddStringWithDelimiter(ByVal addValue As String, ByVal srcValue As String, ByVal delimiter As String)

            Dim result As String = srcValue

            If Not result = Nothing Then result &= delimiter
            result &= addValue

            Return result

        End Function

        ''' <summary>
        ''' TextStringで取得した文字列をdelimiterで区切り、配列として返す。IgnoreStringで指定された文字列は取り込まないものとする
        ''' </summary>
        ''' <param name="target"></param>
        ''' <param name="delimiter"></param>
        ''' <param name="ignoreString"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetArrayFromString(ByVal target As String, ByVal delimiter As String, Optional ByVal ignoreString As String = Nothing) As String()

            Dim Result() As String = target.Split(delimiter)

            If Not ignoreString = Nothing Then
                Dim i As Integer
                For i = 0 To Result.Length - 1
                    Replace(Result(i), ignoreString, Nothing)
                Next
            End If

            Return Result

        End Function

        ''' <summary>
        ''' 指定した文字列から数値のみを取り出します。
        ''' 例外文字を指定した場合は該当する文字のみ含めて返します。
        ''' </summary>
        ''' <param name="target">指定文字列。</param>
        ''' <param name="escapeCharacter">例外文字。リストで複数指定可能。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumericRevice(ByVal target As String, ByVal escapeCharacter As List(Of String)) As String

            Dim result As String = Nothing
            Dim i As Integer
            Dim c As String

            For i = 0 To Len(target) - 1
                c = target.Substring(i, 1)
                If IsNumeric(c) = True Then
                    result &= c
                Else
                    If Not escapeCharacter Is Nothing Then
                        If escapeCharacter.Contains(c) = True Then
                            result &= c
                        End If
                    End If
                End If
            Next

            Return result

        End Function

        ''' <summary>
        ''' 指定した文字列から数値のみを取り出します。
        ''' 例外文字を指定した場合は該当する文字のみ含めて返します。
        ''' </summary>
        ''' <param name="target">指定文字列。</param>
        ''' <param name="escapeCharacter">例外文字。配列で複数指定可能。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumericRevice(ByVal target As String, ByVal escapeCharacter As String()) As String

            Dim escs As New List(Of String)
            For Each el In escapeCharacter
                escs.Add(el)
            Next

            Return NumericRevice(target, escs)

        End Function

        ''' <summary>
        ''' 指定した文字列から数値のみを取り出します。
        ''' 例外文字を指定した場合は該当する文字のみ含めて返します。
        ''' </summary>
        ''' <param name="target">指定文字列。</param>
        ''' <param name="escapePoint">ドット(.)を例外とするかどうかを設定します。
        ''' 小数が含まれるかもしれない場合はTrueにする必要があります。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumericRevice(ByVal target As String, ByVal escapePoint As Boolean) As String

            Dim escs As New List(Of String)
            escs.Add(".")

            Return NumericRevice(target, escs)

        End Function

        ''' <summary>
        ''' 指定した文字列から数値のみを取り出します。
        ''' </summary>
        ''' <param name="target">指定文字列。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumericRevice(ByVal target As String) As String

            Dim escs As New List(Of String)
            escs.Add(".")

            Return NumericRevice(target, False)

        End Function

        ''' <summary>
        ''' 指定した文字列から電話番号として識別できる文字のみを取り出す。
        ''' </summary>
        ''' <param name="target">指定文字列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumericReviceForTelephone(ByVal target As String) As String
            Dim escape As String() = {"#", "*", "-", "(", ")"}
            Return NumericRevice(target, New List(Of String)(escape))
        End Function

        ''' <summary>
        ''' 数字範囲を含む文字列から開始数と終了数と成功か不成功かを返す（例："1000-2000" > 1000,2000）
        ''' </summary>
        ''' <param name="TargetString"></param>
        ''' <param name="StartNumber"></param>
        ''' <param name="EndNumber"></param>
        ''' <param name="delimiter"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumericRangeByString(ByVal targetString As String, ByRef startNumber As Decimal, ByRef endNumber As Decimal, Optional ByVal delimiter As String = "-") As Boolean

            Dim i As Integer
            Dim str1 As String = Replace(targetString, " ", Nothing)

            startNumber = 0
            endNumber = 0

            If IsNumeric(str1) = True Then
                If Decimal.Parse(str1) > 0 Then
                    startNumber = Decimal.Parse(str1)
                    endNumber = Decimal.Parse(str1)
                Else
                    startNumber = Decimal.Parse(str1)
                    endNumber = Decimal.Parse(str1)
                End If
                Return True
            Else
                For i = 0 To Len(str1) - 1
                    If str1.Substring(i, 1) = delimiter Then
                        If IsNumeric(str1.Substring(0, i)) = True And IsNumeric(str1.Substring(i + 1, Len(str1) - (i + 1))) = True Then
                            startNumber = Decimal.Parse(str1.Substring(0, i))
                            endNumber = Decimal.Parse(str1.Substring(i + 1, Len(str1) - (i + 1)))
                            Return True
                        End If
                    End If
                Next
            End If

            Return False

        End Function

        ''' <summary>
        ''' 時間を秒単位で取得し、時間＋分＋秒形式の文字列を返す
        ''' </summary>
        ''' <param name="SecondValue">秒</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertSecondToTime(ByVal secondValue As Integer) As String

            If secondValue = 0 Then Return "0秒"
            Dim value As Integer = Math.Abs(secondValue)
            Dim result As String = Nothing

            Dim jHour As Integer = 0
            Dim jMin As Integer = 0
            Dim jSec As Integer = 0

            jMin = Math.DivRem(value, 60, jSec)
            jHour = Math.DivRem(jMin, 60, jMin)

            If jHour > 0 Then
                result &= jHour & "時間"
            End If
            If jMin > 0 Then
                result &= jMin & "分"
            End If
            If jSec > 0 Then
                result &= jSec & "秒"
            End If

            If secondValue < 0 Then
                result = "-" & result
            End If

            Return result

        End Function

        ''' <summary>
        ''' targetに指定した文字列中にwordに指定した文字列が含まれるかどうかを判別する
        ''' </summary>
        ''' <param name="target"></param>
        ''' <param name="word"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HasString(ByVal target As String, ByVal word As String) As Boolean
            If target = Nothing Or word = Nothing Then
                Return False
            Else
                Return target.Contains(word)
            End If
        End Function

        ''' <summary>
        ''' targetに指定した文字列を文字数がlengthに達するまでfillWordで補填する。
        ''' 半角全角は識別しない。
        ''' </summary>
        ''' <param name="target">元になる文字列</param>
        ''' <param name="length">文字数</param>
        ''' <param name="isByte">Trueの場合、文字数をバイト単位で数える</param>
        ''' <param name="fillWord">補填文字(1文字のみ)</param>
        ''' <param name="fillRight">Trueの場合、文字を右側に補填する</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FillString(ByVal target As String, ByVal length As Integer, Optional ByVal isByte As Boolean = False, _
                                   Optional ByVal fillWord As String = " ", Optional ByVal fillRight As Boolean = False) As String

            Dim result As String = target
            If fillWord = Nothing Then fillWord = " "
            Dim f As String = fillWord.Substring(0, 1)

            If result.Length >= length Then
                Return result.Substring(0, length)
            End If

            If isByte = False Then
                Do Until result.Length = length
                    If fillRight = False Then
                        result = f & result
                    Else
                        result = result & f
                    End If
                Loop
            Else
                Do Until GetTextLengthAsByte(result) = length
                    If fillRight = False Then
                        result = f & result
                    Else
                        result = result & f
                    End If
                Loop
            End If

            Return result

        End Function

        ''' <summary>
        ''' データベースなどから読み取ったフィールドがNullであるかどうか判別する。
        ''' Nullの場合は指定の文字もしくは文字列で返し、そうでない場合は文字列で返す。
        ''' </summary>
        ''' <param name="value">対象となるオブジェクト</param>
        ''' <param name="replaceString">Nullの場合に返す文字もしくは文字列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EscapeNull(ByVal value As Object, Optional ByVal replaceString As String = Nothing) As String

            If value Is Nothing Then
                Return replaceString
            Else
                If value Is DBNull.Value Then
                    Return replaceString
                Else
                    Return value.ToString
                End If
            End If

        End Function

        ''' <summary>
        ''' 指定した文字列が空白かどうかを判別し，空白であれば代わりの文字列を返します。
        ''' </summary>
        ''' <param name="value">指定する文字列。</param>
        ''' <param name="replaceValue">代わりの文字列。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsNothing(ByVal value As String, Optional ByVal replaceValue As String = Nothing) As String
            If value = Nothing OrElse value = "" Then
                Return replaceValue
            Else
                Return value
            End If
        End Function

        ''' <summary>
        ''' 指定した文字列を指定した文字列で囲む。
        ''' </summary>
        ''' <param name="value">対象文字列</param>
        ''' <param name="quotationString">囲いに使用する文字列。デフォルトはシングルクォーテーション</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddQuotation(ByVal value As String, Optional ByVal quotationString As String = "'") As String

            Return quotationString & value & quotationString

        End Function

#End Region

#Region "Randam"

        ''' <summary>
        ''' ランダムな文字列を生成します。
        ''' </summary>
        ''' <param name="intKeyLen">文字列の文字数を指定します。</param>
        ''' <returns>ランダムな文字列（0～9、A～Z、a～zの組み合わせ）。</returns>
        Public Shared Function CreateRandomString(ByVal intKeyLen As Integer) As String

            CreateRandomString = ""

            '指定の文字数になるまでランダムな文字を生成
            Dim strKey As String = ""
            Do Until Len(strKey) >= intKeyLen
                'ランダムな文字を生成
                Dim strKeyChar As String = Chr(GetRollDice(122 - 47) + 47)
                '数字・英字の範囲かチェック
                Select Case strKeyChar
                    Case "0" To "9", "A" To "Z", "a" To "z"
                        strKey = strKey & strKeyChar
                End Select
            Loop

            CreateRandomString = strKey

        End Function

        ''' <summary>
        ''' ランダムな文字列を生成します。
        ''' </summary>
        ''' <param name="length">生成する文字列の長さを指定します。</param>
        ''' <param name="enableChars">文字列に使用できる文字を指定します。
        ''' 指定しない場合は"0123456789abcdefghijklmnopqrstuvwxyz"を使用します。</param>
        ''' <returns>指定された文字を使用して生成された文字列。</returns>
        Public Shared Function CreateRandomString(ByVal length As Integer, ByVal enableChars As String) As String

            If enableChars = Nothing Then enableChars = "0123456789abcdefghijklmnopqrstuvwxyz"

            Dim sb As New System.Text.StringBuilder(length)
            Dim r As New Random()

            For i As Integer = 0 To length - 1
                '文字の位置をランダムに選択
                Dim pos As Integer = r.Next(enableChars.Length)
                '選択された位置の文字を取得
                Dim c As Char = enableChars(pos)
                '文字列に追加
                sb.Append(c)
            Next

            Return sb.ToString()
        End Function


        ''' <summary>
        ''' 暗号サービス プロバイダの暗号乱数ジェネレータを使っての乱数を生成します。
        ''' </summary>
        ''' <param name="NumSides">出力値の最大値</param>
        ''' <returns>乱数（1～指定した最大値）</returns>
        Private Shared Function GetRollDice(ByVal numSides As Integer) As Integer
            ' Create a byte array to hold the random value.
            Dim randomNumber(0) As Byte

            ' Create a new instance of the RNGCryptoServiceProvider.
            Dim Gen As New System.Security.Cryptography.RNGCryptoServiceProvider()

            ' Fill the array with a random value.
            Gen.GetBytes(randomNumber)

            ' Convert the byte to an integer value to make the modulus operation easier.
            Dim rand As Integer = Convert.ToInt32(randomNumber(0))

            ' Return the random number mod the number
            ' of sides.  The possible values are zero-
            ' based, so we add one.
            Return rand Mod numSides + 1
        End Function 'RollDice

#End Region

#Region "Byte"

        ''' <summary>
        ''' 指定された文字列のバイト数を返します。
        ''' </summary>
        ''' <param name="target">バイト数取得の対象となる文字列。</param>
        ''' <param name="enc">エンコードの指定。</param>
        ''' <returns>バイト数。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextLengthAsByte(ByVal target As String, ByVal enc As System.Text.Encoding) As Integer
            Return enc.GetByteCount(target)
        End Function


        ''' <summary>
        ''' 指定された文字列のバイト数を返します。
        ''' エンコードにはシフトJISを使用します。
        ''' </summary>
        ''' <param name="target">バイト数取得の対象となる文字列。</param>
        ''' <returns>バイト数。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextLengthAsByte(ByVal target As String) As Integer
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
            Return GetTextLengthAsByte(target, enc)
        End Function

        ''' <summary>
        ''' バイト配列から文字列を取得します。
        ''' コードは自動判別します。
        ''' エラーの場合はNothingを返します。
        ''' </summary>
        ''' <param name="value">バイト配列。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextFromBytes(ByVal value As Byte()) As String

            Dim result As String = Nothing
            Dim enc As System.Text.Encoding = GetEncode(value)

            If Not enc Is Nothing Then
                Try
                    result = enc.GetString(value)
                Catch ex As Exception

                End Try
            End If

            Return result

        End Function

        ''' <summary>
        ''' 文字列の指定されたバイト位置以降のすべての文字列を返します。
        ''' </summary>
        ''' <param name="stTarget">取り出す元になる文字列。</param>
        ''' <param name="iStart">取り出しを開始する位置。</param>
        ''' <param name="enc">エンコードの指定。</param>
        ''' <returns>指定されたバイト位置以降のすべての文字列。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextAsByteCount(ByVal stTarget As String, ByVal iStart As Integer, ByVal enc As System.Text.Encoding) As String

            Dim btBytes As Byte() = enc.GetBytes(stTarget)
            Return enc.GetString(btBytes, iStart - 1, btBytes.Length - iStart + 1)

        End Function

        ''' <summary>
        ''' 文字列の指定されたバイト位置以降のすべての文字列を返します。
        ''' エンコードにはシフトJISを使用します。
        ''' </summary>
        ''' <param name="stTarget">取り出す元になる文字列。</param>
        ''' <param name="iStart">取り出しを開始する位置。</param>
        ''' <returns>指定されたバイト位置以降のすべての文字列。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextAsByteCount(ByVal stTarget As String, ByVal iStart As Integer) As String

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
            Return GetTextAsByteCount(stTarget, iStart, enc)

        End Function

        ''' <summary>
        ''' 文字列の指定されたバイト位置から、指定されたバイト数分の文字列を返します。
        ''' </summary>
        ''' <param name="stTarget">取り出す元になる文字列。</param>
        ''' <param name="iStart">取り出しを開始する位置。</param>
        ''' <param name="iByteSize">取り出すバイト数。</param>
        ''' <param name="enc">エンコードの指定。</param>
        ''' <returns>指定されたバイト位置から指定されたバイト数分の文字列。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextAsByteCount(ByVal stTarget As String, ByVal iStart As Integer, ByVal iByteSize As Integer, ByVal enc As System.Text.Encoding) As String

            Dim btBytes As Byte() = enc.GetBytes(stTarget)
            Dim count As Integer = 0
            If iByteSize > btBytes.Count Then count = btBytes.Count
            Return enc.GetString(btBytes, iStart - 1, count)

        End Function

        ''' <summary>
        ''' 文字列の指定されたバイト位置から、指定されたバイト数分の文字列を返します。
        ''' エンコードにはシフトJISを使用します。
        ''' </summary>
        ''' <param name="stTarget">取り出す元になる文字列。</param>
        ''' <param name="iStart">取り出しを開始する位置。</param>
        ''' <param name="iByteSize">取り出すバイト数。</param>
        ''' <returns>指定されたバイト位置から指定されたバイト数分の文字列。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextAsByteCount(ByVal stTarget As String, ByVal iStart As Integer, ByVal iByteSize As Integer) As String

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
            Return GetTextAsByteCount(stTarget, iStart, iByteSize, enc)

        End Function

        ''' <summary>
        ''' 文字列の左端から指定したバイト数分の文字列を返します。
        ''' </summary>
        ''' <param name="stTarget">取り出す元になる文字列。</param>
        ''' <param name="iByteSize">取り出すバイト数。</param>
        ''' <param name="enc">エンコードの指定。</param>
        ''' <returns>左端から指定されたバイト数分の文字列。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextAsByteCountFromLeft(ByVal stTarget As String, ByVal iByteSize As Integer, ByVal enc As System.Text.Encoding) As String
            Return GetTextAsByteCount(stTarget, 1, iByteSize, enc)
        End Function

        ''' <summary>
        ''' 文字列の左端から指定したバイト数分の文字列を返します。
        ''' エンコードにはシフトJISを使用します。
        ''' </summary>
        ''' <param name="stTarget">取り出す元になる文字列。</param>
        ''' <param name="iByteSize">取り出すバイト数。</param>
        ''' <returns>左端から指定されたバイト数分の文字列。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextAsByteCountFromLeft(ByVal stTarget As String, ByVal iByteSize As Integer) As String
            Return GetTextAsByteCount(stTarget, 1, iByteSize)
        End Function

        ''' <summary>
        ''' 文字列の右端から指定されたバイト数分の文字列を返します。
        ''' </summary>
        ''' <param name="stTarget">取り出す元になる文字列。</param>
        ''' <param name="iByteSize">取り出すバイト数。</param>
        ''' <param name="enc">エンコードの指定。</param>
        ''' <returns>右端から指定されたバイト数分の文字列。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextAsByteCountFromRight(ByVal stTarget As String, ByVal iByteSize As Integer, ByVal enc As System.Text.Encoding) As String

            Dim btBytes As Byte() = enc.GetBytes(stTarget)
            Dim count As Integer = 0
            If iByteSize > btBytes.Count Then count = btBytes.Count
            Return enc.GetString(btBytes, btBytes.Length - iByteSize, count)

        End Function

        ''' <summary>
        ''' 文字列の右端から指定されたバイト数分の文字列を返します。
        ''' エンコードにはシフトJISを使用します。
        ''' </summary>
        ''' <param name="stTarget">取り出す元になる文字列。</param>
        ''' <param name="iByteSize">取り出すバイト数。</param>
        ''' <returns>右端から指定されたバイト数分の文字列。</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextAsByteCountFromRight(ByVal stTarget As String, ByVal iByteSize As Integer) As String

            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
            Return GetTextAsByteCountFromRight(stTarget, iByteSize, enc)

        End Function

#End Region

#Region "Encoding"

        ''' <summary>
        ''' 文字コードを判別します。
        ''' </summary>
        ''' <remarks></remarks>
        ''' <param name="bytes">文字コードを調べるデータ</param>
        ''' <returns>適当と思われるEncodingオブジェクト。
        ''' 判断できなかった時はnull。</returns>
        Public Shared Function GetEncode(ByVal bytes As Byte()) As System.Text.Encoding
            Const bEscape As Byte = &H1B
            Const bAt As Byte = &H40
            Const bDollar As Byte = &H24
            Const bAnd As Byte = &H26
            Const bOpen As Byte = &H28 ''('
            Const bB As Byte = &H42
            Const bD As Byte = &H44
            Const bJ As Byte = &H4A
            Const bI As Byte = &H49

            Dim len As Integer = bytes.Length
            Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte

            'Encode::is_utf8 は無視

            Dim isBinary As Boolean = False
            Dim i As Integer
            For i = 0 To len - 1
                b1 = bytes(i)
                If b1 <= &H6 OrElse b1 = &H7F OrElse b1 = &HFF Then
                    ''binary'
                    isBinary = True
                    If b1 = &H0 AndAlso i < len - 1 AndAlso bytes(i + 1) <= &H7F Then
                        'smells like raw unicode
                        Return System.Text.Encoding.Unicode
                    End If
                End If
            Next
            If isBinary Then
                Return Nothing
            End If

            'not Japanese
            Dim notJapanese As Boolean = True
            For i = 0 To len - 1
                b1 = bytes(i)
                If b1 = bEscape OrElse &H80 <= b1 Then
                    notJapanese = False
                    Exit For
                End If
            Next
            If notJapanese Then
                Return System.Text.Encoding.ASCII
            End If

            For i = 0 To len - 3
                b1 = bytes(i)
                b2 = bytes(i + 1)
                b3 = bytes(i + 2)

                If b1 = bEscape Then
                    If b2 = bDollar AndAlso b3 = bAt Then
                        'JIS_0208 1978
                        'JIS
                        Return System.Text.Encoding.GetEncoding(50220)
                    ElseIf b2 = bDollar AndAlso b3 = bB Then
                        'JIS_0208 1983
                        'JIS
                        Return System.Text.Encoding.GetEncoding(50220)
                    ElseIf b2 = bOpen AndAlso (b3 = bB OrElse b3 = bJ) Then
                        'JIS_ASC
                        'JIS
                        Return System.Text.Encoding.GetEncoding(50220)
                    ElseIf b2 = bOpen AndAlso b3 = bI Then
                        'JIS_KANA
                        'JIS
                        Return System.Text.Encoding.GetEncoding(50220)
                    End If
                    If i < len - 3 Then
                        b4 = bytes(i + 3)
                        If b2 = bDollar AndAlso b3 = bOpen AndAlso b4 = bD Then
                            'JIS_0212
                            'JIS
                            Return System.Text.Encoding.GetEncoding(50220)
                        End If
                        If i < len - 5 AndAlso _
                            b2 = bAnd AndAlso b3 = bAt AndAlso b4 = bEscape AndAlso _
                            bytes(i + 4) = bDollar AndAlso bytes(i + 5) = bB Then
                            'JIS_0208 1990
                            'JIS
                            Return System.Text.Encoding.GetEncoding(50220)
                        End If
                    End If
                End If
            Next

            'should be euc|sjis|utf8
            'use of (?:) by Hiroki Ohzaki <ohzaki@iod.ricoh.co.jp>
            Dim sjis As Integer = 0
            Dim euc As Integer = 0
            Dim utf8 As Integer = 0
            For i = 0 To len - 2
                b1 = bytes(i)
                b2 = bytes(i + 1)
                If ((&H81 <= b1 AndAlso b1 <= &H9F) OrElse _
                    (&HE0 <= b1 AndAlso b1 <= &HFC)) AndAlso _
                    ((&H40 <= b2 AndAlso b2 <= &H7E) OrElse _
                     (&H80 <= b2 AndAlso b2 <= &HFC)) Then
                    'SJIS_C
                    sjis += 2
                    i += 1
                End If
            Next
            For i = 0 To len - 2
                b1 = bytes(i)
                b2 = bytes(i + 1)
                If ((&HA1 <= b1 AndAlso b1 <= &HFE) AndAlso _
                    (&HA1 <= b2 AndAlso b2 <= &HFE)) OrElse _
                    (b1 = &H8E AndAlso (&HA1 <= b2 AndAlso b2 <= &HDF)) Then
                    'EUC_C
                    'EUC_KANA
                    euc += 2
                    i += 1
                ElseIf i < len - 2 Then
                    b3 = bytes(i + 2)
                    If b1 = &H8F AndAlso (&HA1 <= b2 AndAlso b2 <= &HFE) AndAlso _
                        (&HA1 <= b3 AndAlso b3 <= &HFE) Then
                        'EUC_0212
                        euc += 3
                        i += 2
                    End If
                End If
            Next
            For i = 0 To len - 2
                b1 = bytes(i)
                b2 = bytes(i + 1)
                If (&HC0 <= b1 AndAlso b1 <= &HDF) AndAlso _
                    (&H80 <= b2 AndAlso b2 <= &HBF) Then
                    'UTF8
                    utf8 += 2
                    i += 1
                ElseIf i < len - 2 Then
                    b3 = bytes(i + 2)
                    If (&HE0 <= b1 AndAlso b1 <= &HEF) AndAlso _
                        (&H80 <= b2 AndAlso b2 <= &HBF) AndAlso _
                        (&H80 <= b3 AndAlso b3 <= &HBF) Then
                        'UTF8
                        utf8 += 3
                        i += 2
                    End If
                End If
            Next
            'M. Takahashi's suggestion
            'utf8 += utf8 / 2;

            System.Diagnostics.Debug.WriteLine( _
                String.Format("sjis = {0}, euc = {1}, utf8 = {2}", sjis, euc, utf8))
            If euc > sjis AndAlso euc > utf8 Then
                'EUC
                Return System.Text.Encoding.GetEncoding(51932)
            ElseIf sjis > euc AndAlso sjis > utf8 Then
                'SJIS
                Return System.Text.Encoding.GetEncoding(932)
            ElseIf utf8 > euc AndAlso utf8 > sjis Then
                'UTF8
                Return System.Text.Encoding.UTF8
            End If

            Return Nothing
        End Function


#End Region

#Region "XML"

        ''' <summary>
        ''' XDocumentの内容を文字列で取得します。
        ''' </summary>
        ''' <param name="xdoc"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextFromXmlDocument(ByVal xdoc As XDocument) As String

            Dim sb As New System.Text.StringBuilder
            Dim tr As New System.IO.StringWriter(sb)
            xdoc.Save(tr)
            Return sb.ToString

        End Function

#End Region

    End Class

End Namespace


