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

    Private Sub PrepareMap(ByVal X As UInteger, ByVal Y As UInteger)
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

    Private Sub LoadArrivable(Player As Player)
        ClearBooleanArray2(mArrivable)
        ClearBooleanArray2(mChecked)
        mArrivable(Player.X, Player.Y) = True
        Dim out As Boolean
        Do
            For i1 = 0 To mArrivable.GetLength(1)
                For i2 = 0 To mArrivable.GetLength(2)
                    If mArrivable(i1, i2) Then
                        mChecked(i1, i2) = True
                        If Not IsLined(Directions.Left, i1, i2) Then
                            mArrivable(i1 - 1, i2) = True
                        End If
                        If Not IsLined(Directions.Right, i1, i2) Then
                            mArrivable(i1 + 1, i2) = True
                        End If
                        If Not IsLined(Directions.Up, i1, i2) Then
                            mArrivable(i1, i2 - 1) = True
                        End If
                        If Not IsLined(Directions.Down, i1, i2) Then
                            mArrivable(i1, i2 + 1) = True
                        End If
                    End If
                Next
            Next
            out = True
            For i1 = 0 To mArrivable.GetLength(1)
                For i2 = 0 To mArrivable.GetLength(2)
                    If mArrivable(i1, i2) <> mChecked(i1, i2) Then out = False : Exit For
                Next
                If Not out Then Exit For
            Next
        Loop Until out
    End Sub

    Private Sub ClearBooleanArray2(ByRef Array As Boolean(,))
        For i1 = 0 To Array.GetLength(1)
            For i2 = 0 To Array.GetLength(2)
                Array(i1, i2) = Nothing
            Next
        Next
    End Sub

    Public Function IsIllegalXY(ByVal X As UInteger, ByVal Y As UInteger) As Boolean
        If X >= mapX Or Y >= mapY Then
            Return True
        Else
            Return False
        End If
    End Function

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
        If IsIllegalXY(xP1, yP1) Or IsIllegalXY(xP2, yP2) Then
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

    ReadOnly Property GetHLine() As Boolean(,)
        Get
            Return mHLine
        End Get
    End Property
    ReadOnly Property GetVLine As Boolean(,)
        Get
            Return mVLine
        End Get
    End Property

    ReadOnly Property IsLined(ByVal Direction As Directions, ByVal X As UInteger, ByVal Y As UInteger) As Boolean
        Get
            If IsIllegalXY(X, Y) Then
                Err().Raise(vbObjectError + 1001,, "Illegal X or Y")
            End If
            Select Case Direction
                Case Directions.Left
                    Return mVLine(X, Y)
                Case Directions.Right
                    If X + 1 = mapX Then
                        Return True
                    Else
                        Return mVLine(X + 1, Y)
                    End If
                Case Directions.Up
                    Return mHLine(X, Y)
                Case Directions.Down
                    If Y + 1 = mapY Then
                        Return True
                    Else
                        Return mHLine(X, Y + 1)
                    End If
                Case Else
                    Err().Raise(vbObjectError + 1002,, "Illegal Direction")
                    Return True
            End Select
        End Get
    End Property
End Class
