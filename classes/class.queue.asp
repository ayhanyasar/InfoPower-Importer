<%
'====================================================================
' InfoPower Importer PRO v3.0
' File    : classes/class.queue.asp
' Part    : 1
'====================================================================
Option Explicit

Class Queue

    Private mItems
    Private mPosition

    Private Sub Class_Initialize()

        ReDim mItems(-1)

        mPosition = 0

    End Sub



    Private Sub Class_Terminate()

        Clear

    End Sub



    Public Sub Enqueue(Item)

        Dim n

        n = UBound(mItems) + 1

        ReDim Preserve mItems(n)

        Set mItems(n) = Item

    End Sub



    Public Function Dequeue()

        Dim i
        Dim Obj

        If Count() = 0 Then

            Set Dequeue = Nothing
            Exit Function

        End If

        Set Obj = mItems(0)

        For i = 1 To UBound(mItems)

            Set mItems(i-1) = mItems(i)

        Next

        ReDim Preserve mItems(UBound(mItems)-1)

        Set Dequeue = Obj

    End Function



    Public Function Peek()

        If Count() = 0 Then

            Set Peek = Nothing

        Else

            Set Peek = mItems(0)

        End If

    End Function



    Public Function Item(Index)

        If Index < 0 Then

            Set Item = Nothing
            Exit Function

        End If

        If Index > UBound(mItems) Then

            Set Item = Nothing
            Exit Function

        End If

        Set Item = mItems(Index)

    End Function



    Public Function Count()

        If UBound(mItems) < 0 Then
            Count = 0
        Else
            Count = UBound(mItems) + 1
        End If

    End Function



    Public Function IsEmpty()

        IsEmpty = (Count() = 0)

    End Function



    Public Function First()

        Set First = Peek()

    End Function



    Public Function Last()

        If Count() = 0 Then

            Set Last = Nothing

        Else

            Set Last = mItems(UBound(mItems))

        End If

    End Function
 
     Public Function Contains(ByVal Obj)

        Dim i

        Contains = False

        If Obj Is Nothing Then Exit Function

        For i = 0 To UBound(mItems)

            If Obj Is mItems(i) Then

                Contains = True
                Exit Function

            End If

        Next

    End Function



    Public Function IndexOf(ByVal Obj)

        Dim i

        IndexOf = -1

        If Obj Is Nothing Then Exit Function

        For i = 0 To UBound(mItems)

            If Obj Is mItems(i) Then

                IndexOf = i
                Exit Function

            End If

        Next

    End Function



    Public Function Remove(Index)

        Dim i

        Remove = False

        If Count() = 0 Then Exit Function

        If Index < 0 Then Exit Function
        If Index > UBound(mItems) Then Exit Function

        For i = Index To UBound(mItems) - 1
            Set mItems(i) = mItems(i + 1)
        Next

        ReDim Preserve mItems(UBound(mItems) - 1)

        Remove = True

    End Function



    Public Function MoveNext()

        If mPosition >= Count() Then

            Set MoveNext = Nothing
            Exit Function

        End If

        Set MoveNext = mItems(mPosition)

        mPosition = mPosition + 1

    End Function



    Public Sub Reset()

        mPosition = 0

    End Sub



    Public Function EOF()

        EOF = (mPosition >= Count())

    End Function



    Public Function ToArray()

        ToArray = mItems

    End Function
 
     Public Function Clone()

        Dim Q
        Dim i

        Set Q = New Queue

        If Count() > 0 Then

            For i = 0 To UBound(mItems)
                Q.Enqueue mItems(i)
            Next

        End If

        Set Clone = Q

    End Function



    Public Sub Reverse()

        Dim i
        Dim j
        Dim T

        j = UBound(mItems)

        For i = 0 To (Count() \ 2) - 1

            Set T = mItems(i)
            Set mItems(i) = mItems(j)
            Set mItems(j) = T

            j = j - 1

        Next

    End Sub



    Public Sub Clear()

        ReDim mItems(-1)

        mPosition = 0

    End Sub



    Public Function Capacity()

        Capacity = Count()

    End Function



    Public Function CurrentPosition()

        CurrentPosition = mPosition

    End Function



    Public Property Let Position(Value)

        If Value < 0 Then Value = 0

        If Value > Count() Then Value = Count()

        mPosition = CLng(Value)

    End Property



    Public Property Get Position()

        Position = mPosition

    End Property



    Public Function HasNext()

        HasNext = (mPosition < Count())

    End Function

End Class
%>