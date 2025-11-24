Sub CapitalizarAutores_TodoDocumento()

    Dim p As Paragraph
    Dim posAbreParen As Long
    Dim rAutores As Range
    Dim txt As String
    Dim s As String

    Application.ScreenUpdating = False

    ' Percorre todos os parágrafos do documento
    For Each p In ActiveDocument.Paragraphs

        txt = p.Range.Text

        ' remove o caractere de fim de parágrafo (¶) da análise
        If Len(txt) > 0 Then
            If Right$(txt, 1) = Chr(13) Or Right$(txt, 1) = Chr(7) Then
                txt = Left$(txt, Len(txt) - 1)
            End If
        End If

        ' Procura o primeiro "(" no parágrafo
        posAbreParen = InStr(txt, "(")

        ' Se encontrou, tudo antes dele são autores (+ possivelmente " (Eds)")
        If posAbreParen > 1 Then

            ' Range da parte dos autores (antes do primeiro "(")
            Set rAutores = p.Range.Duplicate
            rAutores.End = rAutores.Start + posAbreParen - 2

            ' Trabalha em string
            s = rAutores.Text

            ' 1) tudo minúsculo
            s = LCase$(s)
            ' 2) Title Case (primeira letra de cada palavra)
            s = StrConv(s, vbProperCase)

            ' 3) Correções finas:

            ' "And" -> "and"
            s = Replace(s, " And ", " and ")
            s = Replace(s, " And,", " and,")
            s = Replace(s, " And;", " and;")
            ' se começar com "And "
            If Left$(s, 4) = "And " Then Mid$(s, 1, 3) = "and"

            ' "Et Al." / "Et Al" -> "et al."
            s = Replace(s, " Et Al.", " et al.")
            s = Replace(s, " Et Al,", " et al,")
            s = Replace(s, " Et Al;", " et al;")
            s = Replace(s, " Et Al", " et al")

            ' devolve o texto corrigido para o documento
            rAutores.Text = s
        End If

    Next p

    Application.ScreenUpdating = True

End Sub
