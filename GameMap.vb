Public Class GameMap

    Public Class Player
        Public X, Y As UInt16
        Public Name As String
        Public Sub New(X As UInt16, Y As UInt16, Name As String)
            Me.X = X
            Me.Y = Y
            Me.Name = Name
        End Sub
    End Class

    Public Class Results
        Public Winner As String
        Public Loser As String
        Public WinnerSize As UInteger
        Public LoserSize As UInteger
        Public Sub New(ByVal Winner As String, ByVal Loser As String,
                       ByVal WinnerSize As UInteger, ByVal LoserSize As UInteger)
            Me.Winner = Winner
            Me.Loser = Loser
            Me.WinnerSize = WinnerSize
            Me.LoserSize = LoserSize
        End Sub
    End Class

    Public Enum States As Integer
        None = 0
        MoveP1 = 1
        MoveP2 = 2
        WinP1 = -1
        WinP2 = -2
        NoWinner = -3
    End Enum

    Public Enum Directions As Integer
        Left = 1
        Right = -1
        Up = 2
        Down = -2
    End Enum

    Private mState As States
    Private mapX, mapY As UInt16
    Private mP1, mP2 As Player
    Private mArrivable, mChecked As Boolean(,)
    Private mHLine, mVLine As Boolean(,)
    Private mSize1, mSize2 As UInteger
    Private mResult As Results

    Private Sub PrepareMap(ByVal X As UInt16, ByVal Y As UInt16)
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

    Private Sub LoadSize()
        mSize1 = 0
        LoadArrivable(mP1)
        For i1 = 0 To mArrivable.GetLength(1)
            For i2 = 0 To mArrivable.GetLength(2)
                If mArrivable(i1, i2) Then mSize1 += 1
            Next
        Next
        mSize2 = 0
        LoadArrivable(mP2)
        For i1 = 0 To mArrivable.GetLength(1)
            For i2 = 0 To mArrivable.GetLength(2)
                If mArrivable(i1, i2) Then mSize2 += 1
            Next
        Next
    End Sub

    Private Sub CheckWin()
        If Not (mArrivable(mP1.X, mP1.Y) And mArrivable(mP2.X, mP2.Y)) Then
            LoadSize()
            If mSize1 > mSize2 Then
                mState = States.WinP1
                mResult = New Results(P1.Name, P2.Name, mSize1, mSize2)
            ElseIf mSize2 > mSize1 Then
                mState = States.WinP2
                mResult = New Results(P2.Name, P1.Name, mSize2, mSize1)
            Else
                mState = States.NoWinner
                mResult = New Results("None", "None", mSize1, mSize1)
            End If
        End If
    End Sub

    Private Sub ClearBooleanArray2(ByRef Array As Boolean(,))
        For i1 = 0 To Array.GetLength(1)
            For i2 = 0 To Array.GetLength(2)
                Array(i1, i2) = Nothing
            Next
        Next
    End Sub

    Public Function IsIllegalXY(ByVal X As UInt16, ByVal Y As UInt16) As Boolean
        If X >= mapX Or Y >= mapY Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function BlockID2X(ByVal BlockID As UInteger) As UInt16
        Return (BlockID - 1) Mod mapX
    End Function

    Public Function BlockID2Y(ByVal BlockID As UInteger) As UInt16
        Return (BlockID - 1) \ mapX
    End Function

    Public Sub New()
        PrepareMap(5, 5)
    End Sub

    Public Sub New(ByVal Length As UInt16)
        PrepareMap(Length, Length)
    End Sub

    Public Sub New(ByVal XLength As UInt16, ByVal YLength As UInt16)
        PrepareMap(XLength, YLength)
    End Sub

    Public Sub Start(ByVal xP1 As UInt16, ByVal yP1 As UInt16,
                     ByVal xP2 As UInt16, ByVal yP2 As UInt16,
                     Optional ByVal nameP1 As String = "Player1",
                     Optional ByVal nameP2 As String = "Player2")
        If IsIllegalXY(xP1, yP1) Or IsIllegalXY(xP2, yP2) Then
            Err().Raise(vbObjectError + 1001,, "Illegal X or Y")
        End If
        mP1 = New Player(xP1, yP1, nameP1)
        mP2 = New Player(xP2, yP2, nameP2)
        mState = States.MoveP1
    End Sub

    Public Sub Clear()
        mState = States.None
        mResult = Nothing
        mP1 = Nothing
        mP2 = Nothing
        mSize1 = 0
        mSize2 = 0
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

    ReadOnly Property XLength As UInt16
        Get
            Return mapX
        End Get
    End Property

    ReadOnly Property YLength As UInt16
        Get
            Return mapY
        End Get
    End Property

    ReadOnly Property MaxBlockID As UInteger
        Get
            Return mapX * mapY
        End Get
    End Property

    ReadOnly Property P1 As Player
        Get
            Return mP1
        End Get
    End Property

    ReadOnly Property P2 As Player
        Get
            Return mP2
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

    ReadOnly Property Result As Results
        Get
            Return mResult
        End Get
    End Property

    ReadOnly Property IsLined(ByVal Direction As Directions, ByVal X As UInt16, ByVal Y As UInt16) As Boolean
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

    Public Function SetLine(ByVal Direction As Directions, ByVal X As UInt16, ByVal Y As UInt16) As Boolean
        If mState <> States.None Then
            Err().Raise(vbObjectError + 1003,, "Illegal Operate")
            Return False
        End If
        If IsIllegalXY(X, Y) Then
            Err().Raise(vbObjectError + 1001,, "Illegal X or Y")
        End If
        Select Case Direction
            Case Directions.Left
                mVLine(X, Y) = True
            Case Directions.Right
                If X + 1 <> mapX Then
                    mVLine(X + 1, Y) = True
                End If
            Case Directions.Up
                mHLine(X, Y) = True
            Case Directions.Down
                If Y + 1 <> mapY Then
                    mHLine(X, Y + 1) = True
                End If
            Case Else
                Err().Raise(vbObjectError + 1002,, "Illegal Direction")
                Return False
        End Select
        Return True
    End Function

    Public Sub SetLines(ByVal HLines As Boolean(,), ByVal VLines As Boolean(,))
        mHLine = HLines
        mVLine = VLines
    End Sub

    Public Function Move(ByVal Direction As Directions) As Boolean
        Dim X As UInt16, Y As UInt16
        Select Case mState
            Case States.MoveP1
                X = mP1.X
                Y = mP1.Y
            Case States.MoveP2
                X = mP2.X
                Y = mP2.Y
            Case Else
                Err().Raise(vbObjectError + 1003,, "Illegal Operate")
                Return False
        End Select
        If IsLined(Direction, X, Y) Then
            Return False
        End If
        Select Case Direction
            Case Directions.Left
                X -= 1
            Case Directions.Right
                X += 1
            Case Directions.Up
                Y -= 1
            Case Directions.Down
                Y += 1
            Case Else
                Err().Raise(vbObjectError + 1002,, "Illegal Direction")
                Return True
        End Select
        Select Case mState
            Case States.MoveP1
                mP1.X = X
                mP1.Y = Y
            Case States.MoveP2
                mP2.X = X
                mP2.Y = Y
            Case Else
                Err().Raise(vbObjectError + 1003,, "Illegal Operate")
                Return False
        End Select
        Return True
    End Function

    Public Function MoveTo(ByVal X As UInt16, ByVal Y As UInt16) As Boolean
        If mState <> States.MoveP1 And mState = States.MoveP2 Then
            Err().Raise(vbObjectError + 1003,, "Illegal Operate")
        End If
        If IsIllegalXY(X, Y) Then
            Err().Raise(vbObjectError + 1001,, "Illegal X or Y")
        End If
        If mArrivable(X, Y) Then
            If mState = States.MoveP1 Then
                mP1.X = X
                mP1.Y = Y
            Else
                mP2.X = X
                mP2.Y = Y
            End If
            Return True
        Else
            Return False
        End If
    End Function

    Public Function MoveTo(ByVal BlockID As UInteger) As Boolean
        Return MoveTo(BlockID2X(BlockID), BlockID2Y(BlockID))
    End Function

    Public Function Line(ByVal Direction As Directions) As Boolean
        Dim X As UInt16, Y As UInt16
        Select Case mState
            Case States.MoveP1
                X = mP1.X
                Y = mP1.Y
            Case States.MoveP2
                X = mP2.X
                Y = mP2.Y
            Case Else
                Err().Raise(vbObjectError + 1003,, "Illegal Operate")
                Return False
        End Select
        If IsLined(Direction, X, Y) Then
            Return False
        End If
        Select Case Direction
            Case Directions.Left
                mVLine(X, Y) = True
            Case Directions.Right
                If X + 1 <> mapX Then
                    mVLine(X + 1, Y) = True
                End If
            Case Directions.Up
                mHLine(X, Y) = True
            Case Directions.Down
                If Y + 1 <> mapY Then
                    mHLine(X, Y + 1) = True
                End If
            Case Else
                Err().Raise(vbObjectError + 1002,, "Illegal Direction")
                Return False
        End Select
        Return True
    End Function

    Public Function MoveToLine(ByVal Direction As Directions, ByVal X As UInt16, ByVal Y As UInt16) As Boolean
        Return Not (IsLined(Direction, X, Y)) AndAlso MoveTo(X, Y) AndAlso Line(Direction)
    End Function

    Public Function MoveToLine(ByVal Direction As Directions, ByVal BlockID As UInteger) As Boolean
        Return MoveToLine(Direction, BlockID2X(BlockID), BlockID2Y(BlockID))
    End Function
End Class
