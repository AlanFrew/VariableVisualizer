Imports System.Text.RegularExpressions

Public Class LocalVariableVisualizer
    ''' <summary>
    ''' Outputs a list of local variables, followed by a line-wise summary of what local variables are in use
    ''' </summary>
    ''' <param name="suppressSubsetsOfSize">Causes sections of the specified size or less to be condensed into a '.'</param>
    ''' <remarks>Does not support multiple variable declarations per line</remarks>
    Public Shared Sub Visualize(ByVal functionText As String, Optional ByVal suppressSubsetsOfSize As Integer = 0, Optional ByVal startingLineNumber As Integer = 0)
        Console.WriteLine()

        Dim localVariables = FindLocalVariables(functionText, startingLineNumber)
        PrintLocalVariables(localVariables)
        Dim sections = FindSections(localVariables)
        PrintSections(sections, suppressSubsetsOfSize)

        Console.WriteLine()
    End Sub

    Private Shared Function FindLocalVariables(ByVal functionText As String, Optional ByVal startingLineNumber As Integer = 0) As Dictionary(Of String, LocalVariableProperties)
        Dim lines As String() = functionText.Split(Environment.NewLine)
        Dim foundLocals As New Dictionary(Of String, LocalVariableProperties)
        Dim lineNumber As Integer = startingLineNumber
        Dim fullMatchRegex As String = String.Empty

        For Each line As String In lines
            line = line.Replace(Constants.vbLf, "")
            If fullMatchRegex <> String.Empty Then
                For Each usage As Match In Regex.Matches(line, fullMatchRegex)
                    foundLocals(usage.Groups(0).ToString()).endingLineNum = lineNumber
                Next
            End If

            Dim match = Regex.Match(line, "Dim\s+(\w+)\s.*")
            If match.Length <> 0 Then
                Dim matchName = match.Groups(1).ToString()
                foundLocals.Add(matchName, New LocalVariableProperties() With
                                       {.name = matchName, .startingLine = match.Groups(0).ToString(), .startingLineNum = lineNumber})
                fullMatchRegex = (fullMatchRegex & "|(" & matchName & "(?=[^\w]))|" & matchName & "$").Trim("|")
            End If
            lineNumber += 1
        Next

        For Each localVariable In foundLocals
            If localVariable.Value.endingLineNum = 0 Then
                localVariable.Value.endingLineNum = lineNumber
            End If
        Next

        Return foundLocals
    End Function

    Private Shared Sub PrintLocalVariables(ByVal localVariables As Dictionary(Of String, LocalVariableProperties))
        Console.WriteLine("Local Variables:")

        For Each localVariable In localVariables
            Console.WriteLine(localVariable.Key & " (" & localVariable.Value.startingLineNum & "-" & localVariable.Value.endingLineNum & ")")
        Next

        Console.WriteLine()
    End Sub

    Private Shared Function FindSections(ByVal localVariables As Dictionary(Of String, LocalVariableProperties)) As List(Of SectionOfScope)
        Dim scopeSections As New List(Of SectionOfScope)
        Dim startingAndEndingLinesSet As New SortedSet(Of Integer)

        'Calculate starting and ending lines
        For Each localvar As LocalVariableProperties In localVariables.Values
            startingAndEndingLinesSet.Add(localvar.startingLineNum)
            startingAndEndingLinesSet.Add(localvar.endingLineNum)
        Next

        'Convert starting and ending lines into sections
        Dim startingAndEndingLines(startingAndEndingLinesSet.Count) As Integer
        startingAndEndingLinesSet.CopyTo(startingAndEndingLines)
        For i As Integer = 0 To startingAndEndingLines.Count - 2
            scopeSections.Add(New SectionOfScope With {.startingLineNum = startingAndEndingLines(i), .endingLineNum = startingAndEndingLines(i + 1)})
        Next

        'Populate local variables for each section
        For Each section In scopeSections
            For Each localVar In localVariables.Values
                If localVar.startingLineNum <= section.startingLineNum AndAlso localVar.endingLineNum >= section.endingLineNum Then
                    section.variables.Add(localVar.name)

                End If
            Next
        Next

        Return scopeSections
    End Function

    Private Shared Sub PrintSections(ByVal sections As List(Of SectionOfScope), Optional ByVal suppressSubsetsOfSize As Integer = 0)
        Console.WriteLine("Sections:")

        Dim lastSectionSuppressed As Boolean = False

        For Each section In sections
            If section.variables.Any() Then
                If section.endingLineNum - section.startingLineNum > suppressSubsetsOfSize Then
                    PrintSectionExpanded(section, lastSectionSuppressed)

                    lastSectionSuppressed = False
                Else
                    Console.Write(".")

                    lastSectionSuppressed = True
                End If
            Else
                Console.WriteLine("*********")
            End If
        Next
    End Sub

    Private Shared Sub PrintSectionExpanded(ByVal section As SectionOfScope, ByVal lastSuppressed As Boolean)
        If lastSuppressed Then
            Console.WriteLine()
        End If

        Console.Write("(" & section.startingLineNum & "-" & section.endingLineNum & ") ")
        For Each variableName In section.variables
            Console.Write(variableName & " ")
        Next
        Console.WriteLine()
    End Sub
End Class

Friend Class LocalVariableProperties
    Friend name As String
    Friend startingLineNum As Integer
    Friend startingLine As String
    Friend endingLineNum As Integer
End Class

Friend Class SectionOfScope
    Friend startingLineNum As Integer
    Friend endingLineNum As Integer
    Friend variables As New List(Of String)
End Class