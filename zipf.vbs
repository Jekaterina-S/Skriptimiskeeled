Option Explicit

' Check arguments
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript zipf.vbs <PathToTextFile> <NumberOfMostPopularWords>"
    WScript.Quit
End If

Dim objFSO, objFile, dictWords, dictShortForms, strLine, arrWords, word, strFilePath, topN

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set dictWords = CreateObject("Scripting.Dictionary")
Set dictShortForms = CreateObject("Scripting.Dictionary")

strFilePath = WScript.Arguments(0)
topN = Int(WScript.Arguments(1))

Set objFile = objFSO.OpenTextFile(strFilePath, 1)

' Regular expression
Dim regEx
Set regEx = New RegExp
regEx.Pattern = "[^a-zA-Z0-9']"
regEx.Global = True

Dim baseForms
' Base forms for replacement
' Source for forms: https://dictionary.cambridge.org/grammar/british-grammar/
Set baseForms = CreateObject("Scripting.Dictionary")
' Pronouns
baseForms.Add "i'm", "i be"
baseForms.Add "him", "he"
baseForms.Add "his", "he"
baseForms.Add "her", "she"
baseForms.Add "hers", "she"
baseForms.Add "its", "it"
' Be forms
baseForms.Add "am", "be"
baseForms.Add "is", "be"
baseForms.Add "are", "be"
baseForms.Add "being", "be"
baseForms.Add "been", "be"
baseForms.Add "was", "be"
baseForms.Add "were", "be"
' Do forms
baseForms.Add "does", "do"
baseForms.Add "did", "do"
' Not forms
baseForms.Add "don't", "do not"
baseForms.Add "doesn't", "do not"
baseForms.Add "didn't", "do not"
baseForms.Add "can't", "can not"
baseForms.Add "cannot", "can not"
baseForms.Add "couldn't", "could not"
baseForms.Add "isn't", "be not"
baseForms.Add "aren't", "be not"
baseForms.Add "wasn't", "be not"
baseForms.Add "weren't", "be not"
baseForms.Add "won't", "will not"
' Have forms
baseForms.Add "has", "have"
baseForms.Add "had", "have"
baseForms.Add "hasn't", "have not"
baseForms.Add "hadn't", "have not"
baseForms.Add "haven't", "have not"
baseForms.Add "'ve", "have"
' Extra
baseForms.Add "an'", "and"
baseForms.Add "o'clock", "of the clock"

' Read file
Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    strLine = regEx.Replace(strLine, " ")
    arrWords = Split(LCase(strLine))

    For Each word In arrWords
        If word <> "" Then
            Dim baseWord, prefix, suffix

            ' Check for base form replacements
            If baseForms.Exists(word) Then
                ' For cases like is not, can not etc
                Dim replacementWords, replacementWord
                replacementWords = Split(baseForms(word))
                For Each replacementWord In replacementWords
                    If dictWords.Exists(replacementWord) Then
                        dictWords(replacementWord) = dictWords(replacementWord) + 1
                    Else
                        dictWords.Add replacementWord, 1
                    End If
                Next
            ' Check for apostrophes
            ElseIf InStr(word, "'") > 0 Then
                prefix = Split(word, "'")(0)
                suffix = "'" & Split(word, "'")(1)

                ' Check for short form 've
                If baseForms.Exists(suffix) Then
                    If dictWords.Exists(prefix) Then
                        dictWords(prefix) = dictWords(prefix) + 1
                    Else
                        dictWords.Add prefix, 1
                    End If

                    baseWord = baseForms(suffix)
                    If dictWords.Exists(baseWord) Then
                        dictWords(baseWord) = dictWords(baseWord) + 1
                    Else
                        dictWords.Add baseWord, 1
                    End If
                Else
                    ' Count the whole short form as is
                    If dictShortForms.Exists(word) Then
                        dictShortForms(word) = dictShortForms(word) + 1
                    Else
                        dictShortForms.Add word, 1
                    End If
                End If
            ' Otherwise, count the word
            Else
                If dictWords.Exists(word) Then
                    dictWords(word) = dictWords(word) + 1
                Else
                    dictWords.Add word, 1
                End If
            End If
        End If
    Next
Loop

objFile.Close

' Dictionary sort 
Function SortDictionary(dict)
    Dim sortedKeys(), sortedCounts(), i, j, tempKey, tempValue, key
    ReDim sortedKeys(dict.Count - 1)
    ReDim sortedCounts(dict.Count - 1)

    i = 0
    For Each key In dict.Keys
        sortedKeys(i) = key
        sortedCounts(i) = dict(key)
        i = i + 1
    Next

    ' Bubble sort
    For i = LBound(sortedCounts) To UBound(sortedCounts) - 1
        For j = i + 1 To UBound(sortedCounts)
            If sortedCounts(i) < sortedCounts(j) Then
                tempValue = sortedCounts(i)
                sortedCounts(i) = sortedCounts(j)
                sortedCounts(j) = tempValue

                tempKey = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = tempKey
            End If
        Next
    Next

    SortDictionary = Array(sortedKeys, sortedCounts)
End Function

WScript.Echo "CHECKING THE ZIPF'S LAW" & vbCrLf
WScript.Echo "The first column is the number of corresponding words in the text and the second column is the number of words which should occur in the text according to the Zipf's law." & vbCrLf

' Sort word dictionary
Dim sortedWordResults, sortedShortFormResults, rank, topFreq, i
sortedWordResults = SortDictionary(dictWords)
topFreq = sortedWordResults(1)(0)
rank = 1

' Popular words output with Zipf's prediction
WScript.Echo "The most popular words in " & strFilePath & " are:" & vbCrLf
For i = 0 To topN - 1
    If i > UBound(sortedWordResults(0)) Then Exit For
    WScript.Echo sortedWordResults(0)(i) & vbTab & sortedWordResults(1)(i) & Space(2) & Round(topFreq / rank)
    rank = rank + 1
Next

WScript.Echo

' Short form output
WScript.Echo "The most popular short forms are:" & vbCrLf
sortedShortFormResults = SortDictionary(dictShortForms)
For i = 0 To topN - 1
    If i > UBound(sortedShortFormResults(0)) Then Exit For
    WScript.Echo sortedShortFormResults(0)(i) & vbTab & sortedShortFormResults(1)(i)
Next

' Clean up objects
Set objFSO = Nothing
Set dictWords = Nothing
Set dictShortForms = Nothing
Set regEx = Nothing
Set baseForms = Nothing
