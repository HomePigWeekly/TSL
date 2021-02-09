Public Class GameMap

    Public Class Player
        Public X, Y As UInteger
        Public Sub New(X As UInteger, Y As UInteger)
            Me.X = X
            Me.Y = Y
        End Sub
    End Class

    Public Enum States As Integer
        None = 0
        MoveP1 = 1
        MoveP2 = 2
        WinP1 = -1
        WinP2 = -2
    End Enum

    Public Enum Directions As Integer
        Left = 1
        Right = -1
        Up = 2
        Down = -2
    End Enum

    Private mState As States
    Private mapX, mapY As UInteger
    Private mP1, mP2 As Player
    Private mArrivable, mChecked As Boolean(,)
    Private mHLine, mVLine As Boolean(,)

    Private Sub PrepareMap(X As UInteger, Y As UInteger)
        If X = 0 Or Y = 0 Then
            Err().Raise(vbObjectError + 1000,, "Illegal Length")
            Exit Sub
        End If
        Me.mapX = X
        Me.mapY = Y
        ReDim Me.mArrivable(X, Y)
        ReDim Me.mChecked(X, Y)
        ReDim Me.mHLine(X, Y)
        ReDim Me.mVLine(X, Y)
    End Sub

    Private Sub ClearBooleanArray2(Array As Boolean(,))
        For i1 = 0 To Array.GetLength(1)
            For i2 = 0 To Array.GetLength(2)
                Array(i1, i2) = Nothing
            Next
        Next
    End Sub

    Public Sub New()
        PrepareMap(5, 5)
    End Sub

    Public Sub New(ByVal Length As UInteger)
        PrepareMap(Length, Length)
    End Sub

    Public Sub New(ByVal XLength As UInteger, ByVal YLength As UInteger)
        PrepareMap(XLength, YLength)
    End Sub

    Public Sub Start(ByVal xP1 As UInteger, ByVal yP1 As UInteger,
                     ByVal xP2 As UInteger, ByVal yP2 As UInteger)
        If xP1 >= mapX Or xP2 >= mapX Or
            yP1 >= mapY Or yP2 >= mapY Then
            Err().Raise(vbObjectError + 1001,, "Illegal X or Y")
        End If
        mP1 = New Player(xP1, yP1)
        mP2 = New Player(xP2, yP2)
    End Sub

    Public Sub Clear()
        mState = States.None
        mP1 = Nothing
        mP2 = Nothing
        ClearBooleanArray2(mArrivable)
        ClearBooleanArray2(mChecked)
        ClearBooleanArray2(mHLine)
        ClearBooleanArray2(mVLine)
    End Sub

    Public ReadOnly Property State As States
        Get
            Return mState
        End Get
    End Property

    ReadOnly Property XLength
        Get
            Return mapX
        End Get
    End Property

    ReadOnly Property YLength
        Get
            Return mapY
        End Get
    End Property

    ReadOnly Property P1 As Player
        Get
            Return New Player(mP1.X, mP1.Y)
        End Get
    End Property

    ReadOnly Property P2 As Player
        Get
            Return New Player(mP2.X, mP2.Y)
        End Get
    End Property
End Class
