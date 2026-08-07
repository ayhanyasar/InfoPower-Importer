<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.spec.asp
' Part    : 1
'====================================================================
'Option Explicit

Class Specification

    Private mItems

    Private Sub Class_Initialize()

        Set mItems = Server.CreateObject("Scripting.Dictionary")
        mItems.CompareMode = 1

    End Sub

    Private Sub Class_Terminate()

        If IsObject(mItems) Then
            mItems.RemoveAll
            Set mItems = Nothing
        End If

    End Sub



    Public Sub Add(Name, Value)

        Name = Normalize(Name)
        Value = Trim(CStr(Value))

        If Len(Name) = 0 Then Exit Sub

        If mItems.Exists(Name) Then
            mItems(Name) = Value
        Else
            mItems.Add Name, Value
        End If

    End Sub



    Public Function Item(Name)

        Name = Normalize(Name)

        If mItems.Exists(Name) Then
            Item = mItems(Name)
        Else
            Item = ""
        End If

    End Function



    Public Function Exists(Name)

        Exists = mItems.Exists(Normalize(Name))

    End Function



    Public Function Count()

        Count = mItems.Count

    End Function



    Public Function Keys()

        Keys = mItems.Keys

    End Function



    Public Function Values()

        Values = mItems.Items

    End Function



    Public Function Dictionary()

        Set Dictionary = mItems

    End Function



    Public Sub Remove(Name)

        Name = Normalize(Name)

        If mItems.Exists(Name) Then
            mItems.Remove Name
        End If

    End Sub



    Public Sub Clear()

        mItems.RemoveAll

    End Sub



    Public Function LoadFromDictionary(ByVal D)

        Dim K

        If D Is Nothing Then

            LoadFromDictionary = False
            Exit Function

        End If

        For Each K In D.Keys

            Add K, D(K)

        Next

        LoadFromDictionary = True

    End Function



    Public Function ParseTable(TableNode)

        Dim Rows
        Dim Row
        Dim Cells

        ParseTable = 0

        If TableNode Is Nothing Then Exit Function

        On Error Resume Next

        Set Rows = TableNode.getElementsByTagName("tr")

        If Err.Number <> 0 Then

            Err.Clear
            Exit Function

        End If

        For Each Row In Rows

            Set Cells = Row.getElementsByTagName("td")

            If Cells.Length >= 2 Then

                Add _
                    Cells(0).innerText, _
                    Cells(1).innerText

                ParseTable = ParseTable + 1

            End If

        Next

    End Function
 
     Public Function ParseDefinitionList(Node)

        Dim DTs
        Dim DDs
        Dim i
        Dim C

        C = 0

        If Node Is Nothing Then

            ParseDefinitionList = 0
            Exit Function

        End If

        Set DTs = Node.getElementsByTagName("dt")
        Set DDs = Node.getElementsByTagName("dd")

        If DTs Is Nothing Then

            ParseDefinitionList = 0
            Exit Function

        End If

        For i = 0 To DTs.Length - 1

            If i < DDs.Length Then

                Add DTs(i).innerText, DDs(i).innerText

                C = C + 1

            End If

        Next

        ParseDefinitionList = C

    End Function



    Public Function Merge(ByVal Spec)

        Dim K

        If Spec Is Nothing Then

            Merge = False
            Exit Function

        End If

        For Each K In Spec.Dictionary.Keys

            Add K, Spec.Item(K)

        Next

        Merge = True

    End Function



    Public Function ToJson()

        Dim K
        Dim S

        S = "{"

        For Each K In mItems.Keys

            If Right(S,1) <> "{" Then
                S = S & ","
            End If

            S = S & _
                """" & EscapeJson(K) & """:" & _
                """" & EscapeJson(mItems(K)) & """"

        Next

        S = S & "}"

        ToJson = S

    End Function



    Public Function ToText()

        Dim K
        Dim S

        S = ""

        For Each K In mItems.Keys

            S = S & K & ": " & mItems(K) & vbCrLf

        Next

        ToText = S

    End Function



    Public Function Clone()

        Dim O
        Dim K

        Set O = New Specification

        For Each K In mItems.Keys

            O.Add K, mItems(K)

        Next

        Set Clone = O

    End Function
 
     Private Function Normalize(Value)

        Dim S

        S = Trim(CStr(Value))

        S = Replace(S, vbCr, "")
        S = Replace(S, vbLf, "")
        S = Replace(S, vbTab, " ")

        Do While InStr(S, "  ") > 0
            S = Replace(S, "  ", " ")
        Loop

        Normalize = LCase(S)

    End Function



    Private Function EscapeJson(Value)

        Dim S

        S = CStr(Value)

        S = Replace(S, "\", "\\")
        S = Replace(S, """", "\""")
        S = Replace(S, "/", "\/")
        S = Replace(S, vbCrLf, "\n")
        S = Replace(S, vbCr, "\n")
        S = Replace(S, vbLf, "\n")
        S = Replace(S, vbTab, "\t")

        EscapeJson = S

    End Function



    Public Function SaveToProduct(ByRef Product)

        Dim K

        If Product Is Nothing Then

            SaveToProduct = False
            Exit Function

        End If

        For Each K In mItems.Keys

            Product.AddSpecification K, mItems(K)

        Next

        SaveToProduct = True

    End Function



    Public Function LoadFromProduct(ByVal Product)

        Dim Specs
        Dim K

        If Product Is Nothing Then

            LoadFromProduct = False
            Exit Function

        End If

        Set Specs = Product.Specifications()

        For Each K In Specs.Keys

            Add K, Specs(K)

        Next

        LoadFromProduct = True

    End Function

End Class
%>