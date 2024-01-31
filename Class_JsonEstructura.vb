

Public Class Class_JsonEstructura

    Public Structure ListaValores
        Dim Clave As String
        Dim Valor As String
    End Structure

    Public Valores As New List(Of ListaValores) 'New ListaValores() {}

    Friend Sub ObtenerElementos(ByRef Texto As String)

        Texto = Replace(Texto, """", "")

        Dim Elementos() As String = New String() {}
        'ReDim Valores(Elementos.Length - 1)

        Dim PosicionObjetos As UInteger = 0


        If 1 = CInt(Texto.IndexOf("{")) Then


            For Indice As UInteger = 0 To Texto.Length - 1

                Dim PrimerDelimitador As UInteger = Texto.IndexOf(":", CInt(PosicionObjetos))
                If Texto.Chars(PrimerDelimitador + 1) = "{" Then
                    Try
                        Dim PosBrazo As UInteger = Texto.IndexOf("}", CInt(PrimerDelimitador))
                        Dim NuevoPar As New ListaValores
                        NuevoPar.Clave = Replace(Texto.Substring(PosicionObjetos, PrimerDelimitador - PosicionObjetos), ",", "")
                        Dim Cantidad As Integer = PosBrazo - (PrimerDelimitador)
                        Dim Posicion As Integer = PrimerDelimitador + 1
                        NuevoPar.Valor = Texto.Substring(Posicion, Cantidad)
                        Valores.Add(NuevoPar)
                        PosicionObjetos = PosBrazo + 1
                        Indice = PosBrazo + 1
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try
                ElseIf Texto.Chars(PrimerDelimitador + 1) = "[" Then
                    Try
                        Dim PosBrazo As UInteger = Texto.IndexOf("]", CInt(PrimerDelimitador))
                        Dim NuevoPar As New ListaValores
                        NuevoPar.Clave = Replace(Texto.Substring(PosicionObjetos, PrimerDelimitador - PosicionObjetos), ",", "")
                        Dim Cantidad As Integer = PosBrazo - (PrimerDelimitador)
                        Dim Posicion As Integer = PrimerDelimitador + 1
                        NuevoPar.Valor = Texto.Substring(Posicion, Cantidad)
                        Valores.Add(NuevoPar)
                        PosicionObjetos = PosBrazo + 1
                        Indice = PosBrazo + 1
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try
                Else
                    Try
                        Dim Par() As String = Split(Elementos(Indice), ":")
                        Dim NuevoPar As New ListaValores
                        NuevoPar.Clave = Replace(Par(0), "{", "")
                        NuevoPar.Valor = Replace(Par(1), "}", "")
                        Valores.Add(NuevoPar)
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try

                End If

            Next

        Else
            Elementos = Split(Texto, ",")
            'ReDim Valores(Elementos.Length)


            For Indice As UInteger = 0 To Elementos.Length - 1

                Dim PrimerDelimitador As UInteger = Texto.IndexOf(":", CInt(PosicionObjetos))
                If Texto.Chars(PrimerDelimitador + 1) = "{" Then
                    Try
                        Dim PosBrazo As UInteger = Texto.IndexOf("}", CInt(PrimerDelimitador))
                        Dim NuevoPar As New ListaValores
                        NuevoPar.Clave = Replace(Texto.Substring(PosicionObjetos, PrimerDelimitador - PosicionObjetos), ",", "")
                        Dim Cantidad As Integer = PosBrazo - (PrimerDelimitador)
                        Dim Posicion As Integer = PrimerDelimitador + 1
                        NuevoPar.Valor = Texto.Substring(Posicion, Cantidad)
                        Valores.Add(NuevoPar)
                        PosicionObjetos = PosBrazo + 1
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try
                ElseIf Texto.Chars(PrimerDelimitador + 1) = "[" Then
                    Try
                        Dim PosBrazo As UInteger = Texto.IndexOf("]", CInt(PrimerDelimitador))
                        Dim NuevoPar As New ListaValores
                        NuevoPar.Clave = Replace(Texto.Substring(PosicionObjetos, PrimerDelimitador - PosicionObjetos), ",", "")
                        Dim Cantidad As Integer = PosBrazo - (PrimerDelimitador)
                        Dim Posicion As Integer = PrimerDelimitador + 1
                        NuevoPar.Valor = Texto.Substring(Posicion, Cantidad)
                        Valores.Add(NuevoPar)
                        PosicionObjetos = PosBrazo + 1
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try
                Else
                    Try
                        Dim Par() As String = Split(Elementos(Indice), ":")
                        Dim NuevoPar As New ListaValores
                        NuevoPar.Clave = Replace(Par(0), "{", "")
                        NuevoPar.Valor = Replace(Par(1), "}", "")
                        Valores.Add(NuevoPar)
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try

                End If
                If PosicionObjetos >= Texto.Length Then
                    Exit For
                End If

            Next


        End If


    End Sub

    Private m_StringValue As String = String.Empty

    Public Property StringValue() As String
        Get
            Return m_StringValue
        End Get
        Set(ByVal value As String)
            m_StringValue = value
            GetNodes(m_StringValue)
        End Set
    End Property

    Public Property Header As String = String.Empty

    Friend Property BracesOrBrackets As String = String.Empty  ' array=Brackets, Braces=objeto
    Friend Property ContainsHijos As Boolean = False

    Public Property BracketsList As New List(Of Class_JsonEstructura)
    Public Property BracesList As New List(Of Class_JsonEstructura)

    Private Sub GetNodes(ByRef Input As String)
        Dim braceCount As Integer = 0
        Dim cketsCount As Integer = 0
        Dim braceStart As Integer = 0
        Dim cketsStart As Integer = 0
        Dim lastStart As Integer = 0

        For i As Integer = 0 To Input.Length - 1
            If cketsCount > 1 Or braceCount > 1 Then
                Me.ContainsHijos = True
            End If

            If (Input(i) = "{") And ((BracesOrBrackets = "") Or (BracesOrBrackets = "Braces")) Then


                braceCount += 1
                If (BracesOrBrackets = "") Then
                    braceStart = i
                    BracesOrBrackets = "Braces"
                End If

            ElseIf (Input(i) = "[") And ((BracesOrBrackets = "") Or (BracesOrBrackets = "Brackets")) Then


                cketsCount += 1
                If (BracesOrBrackets = "") Then
                    cketsStart = i
                    BracesOrBrackets = "Brackets"
                End If

            ElseIf (Input(i) = "}") And (BracesOrBrackets = "Braces") Then
                braceCount -= 1
                If braceCount = 0 Then
                    Dim node As New Class_JsonEstructura()
                    Try
                        Dim headerStart As Integer = Input.LastIndexOf(":", 1) + 1
                        Dim headerEnd As Integer = braceStart - 1
                        node.Header = GetHeader(Input.Substring(lastStart + 1, (i - braceStart)).Trim())
                    Catch ex As Exception
                        ' Manejar el error de manera adecuada, por ejemplo, imprimir un mensaje
                        Console.WriteLine("Error obteniendo el nombre del nodo: " & ex.Message)
                    End Try
                    Dim Texto As String = Input.Substring(braceStart + 1, i - braceStart - 1)
                    node.StringValue = Input.Substring(braceStart + 1, (i - braceStart) - 1)
                    node.ObtenerElementos(node.StringValue)
                    BracesOrBrackets = ""
                    BracesList.Add(node)
                    braceStart = -1
                    lastStart = i
                    braceStart = i + 1
                End If

            ElseIf (Input(i) = "]") And (BracesOrBrackets = "Brackets") Then
                cketsCount -= 1
                If cketsCount = 0 Then
                    Dim node As New Class_JsonEstructura()
                    Try
                        node.Header = GetHeader(Mid(Input, lastStart + 1, i - cketsStart - 1))
                    Catch ex As Exception
                        ' Manejar el error de manera adecuada, por ejemplo, imprimir un mensaje
                        Console.WriteLine("Error obteniendo el nombre del nodo: " & ex.Message)
                    End Try
                    node.StringValue = Input.Substring(cketsStart + 1, (i - cketsStart) - 1)
                    node.ObtenerElementos(node.StringValue)
                    BracesOrBrackets = ""
                    BracketsList.Add(node)
                    cketsStart = -1
                    lastStart = i
                End If

            ElseIf Input(i) = ":" Then
                'lastColonIndex = i
            End If
        Next

        Dim aaasdf As Boolean = False
    End Sub

    Private Function GetHeader(ByRef Input As String) As String
        Dim TextoSalida As String = ""
        Dim Comillas As String = """"
        Dim Llave As String = "{"
        Dim Coma As String = ","
        Dim PosicinInicio As UInteger = 0
        Dim Texto As String = Replace(Input, """", "")
        Texto = Replace(Texto, " ", "")
        For i As Integer = 0 To Input.Length - 1
            Try
                If Input(i) = Coma Then
                    PosicinInicio = i
                End If
                If Input(i) = ":" Then

                    TextoSalida = Input.Substring(PosicinInicio, i - PosicinInicio - 1).ToString
                    TextoSalida = Replace(TextoSalida, ":", "").Trim
                    TextoSalida = Replace(TextoSalida, ",", "").Trim
                    TextoSalida = Replace(TextoSalida, """", "").Trim
                    Exit For
                End If
            Catch ex As Exception

            End Try
        Next
        Return TextoSalida
    End Function

End Class


