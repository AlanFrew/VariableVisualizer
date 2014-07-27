Imports System.IO

Module Module1
    Sub Main()
        Dim regexTest = "Console.WriteLine(""Local Variables:"")"
        Dim testText As String =
            " Imports System.Text.RegularExpressions" & vbCrLf &
" " & vbCrLf &
" Public Class LocalVariableVisualizer" & vbCrLf &
"     Public Shared Function Visualize(ByVal functionText As String) As String" & vbCrLf &
"         Console.WriteLine(""Local Variables:"")" & vbCrLf &
" " & vbCrLf &
"         Dim lines As String() = functionText.Split(Environment.NewLine)" & vbCrLf &
"         Dim foundLocals As New Dictionary(Of String, LocalVariableProperties)" & vbCrLf &
"         Dim lineNumber As Integer = 0" & vbCrLf &
"         Dim fullMatchRegex As String = String.Empty" & vbCrLf &
" " & vbCrLf &
"         For Each line As String In lines" & vbCrLf &
"             line = line.Replace(Constants.vbLf, """")" & vbCrLf &
"             If fullMatchRegex <> String.Empty Then" & vbCrLf &
"                 For Each usage As Match In Regex.Matches(line, fullMatchRegex)" & vbCrLf &
"                     foundLocals(usage.Groups(0).ToString()).endingLineNum = lineNumber" & vbCrLf &
"                 Next" & vbCrLf &
"             End If" & vbCrLf &
" " & vbCrLf &
"             Dim match = Regex.Match(line, ""Dim\s+(\w+)\s.*"")" & vbCrLf &
"             If match.Length <> 0 Then" & vbCrLf &
"                 Dim matchName = match.Groups(1).ToString()" & vbCrLf &
"                 foundLocals.Add(matchName, New LocalVariableProperties() With" & vbCrLf &
"                                        {.name = matchName, .startingLine = match.Groups(0).ToString(), .startingLineNum = lineNumber})" & vbCrLf &
" this is a line of nothinnnnng" & vbCrLf &
"                 fullMatchRegex = (fullMatchRegex & ""|("" & matchName & ""(?=[^\w]))|"" & matchName & ""$"").Trim(""|"")" & vbCrLf &
"             End If" & vbCrLf &
"             lineNumber += 1" & vbCrLf &
"         Next" & vbCrLf &
" " & vbCrLf &
"         For Each localVariable In foundLocals" & vbCrLf &
"             If localVariable.Value.endingLineNum = 0 Then" & vbCrLf &
"                 localVariable.Value.endingLineNum = lineNumber" & vbCrLf &
"             End If" & vbCrLf &
" " & vbCrLf &
"             Console.WriteLine(localVariable.Key & "" ("" & localVariable.Value.startingLineNum & ""-"" & localVariable.Value.endingLineNum & "")"")" & vbCrLf &
"         Next" & vbCrLf &
"         Console.WriteLine()" & vbCrLf &
" " & vbCrLf &
"         Console.WriteLine(""Sections:"")" & vbCrLf &
"         Dim scopeSections As New List(Of SectionOfScope)" & vbCrLf &
"         Dim startingAndEndingLines As New SortedSet(Of Integer)" & vbCrLf &
"         For Each localvar As LocalVariableProperties In foundLocals.Values" & vbCrLf &
"             startingAndEndingLines.Add(localvar.startingLineNum)" & vbCrLf &
"             startingAndEndingLines.Add(localvar.endingLineNum)" & vbCrLf &
"         Next" & vbCrLf &
" " & vbCrLf &
"         For i As Integer = 0 To startingAndEndingLines.Count - 2" & vbCrLf &
"             scopeSections.Add(New SectionOfScope With {.startingLineNum = startingAndEndingLines(i), .endingLineNum = startingAndEndingLines(i + 1)})" & vbCrLf &
"         Next" & vbCrLf &
" " & vbCrLf &
"         For Each section In scopeSections" & vbCrLf &
"             For Each localVar In foundLocals.Values" & vbCrLf &
"                 If localVar.startingLineNum <= section.startingLineNum AndAlso localVar.endingLineNum >= section.endingLineNum Then" & vbCrLf &
"                     section.variables.Add(localVar.name)" & vbCrLf &
"                 End If" & vbCrLf &
"             Next" & vbCrLf &
" " & vbCrLf &
"             If section.variables.Any() Then" & vbCrLf &
"                 Console.Write(""Section ("" & section.startingLineNum & ""-"" & section.endingLineNum & "") "")" & vbCrLf &
"                 For Each variableName In section.variables" & vbCrLf &
"                     Console.Write(variableName & "" "")" & vbCrLf &
"                 Next" & vbCrLf &
"                 Console.WriteLine()" & vbCrLf &
"             Else" & vbCrLf &
"                 Console.WriteLine(""*************"")" & vbCrLf &
"             End If" & vbCrLf &
"         Next" & vbCrLf &
"         Return String.Empty" & vbCrLf &
"     End Function" & vbCrLf &
" End Class" & vbCrLf &
" " & vbCrLf &
" Friend Class LocalVariableProperties" & vbCrLf &
"     Friend name As String" & vbCrLf &
"     Friend startingLineNum As Integer" & vbCrLf &
"     Friend startingLine As String" & vbCrLf &
"     Friend endingLineNum As Integer" & vbCrLf &
" End Class" & vbCrLf &
" " & vbCrLf &
" Friend Class SectionOfScope" & vbCrLf &
"     Friend startingLineNum As Integer" & vbCrLf &
"     Friend endingLineNum As Integer" & vbCrLf &
"     Friend variables As New List(Of String)" & vbCrLf &
" End Class"

        LocalVariableVisualizer.Visualize(testText, 1, 7)
        Console.ReadLine()
    End Sub

End Module
