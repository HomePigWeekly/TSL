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

    Private mState As States
    Private mapX, mapY As UInteger
    Private mP1, mP2 As Player
    Private mArrivable, mChecked As Boolean(,)
    Private mHLine, mVLine As Boolean(,)

    Private Sub PrepareMap(X As UInteger, Y As UInteger)
        Me.mapX = X
        Me.mapY = Y
        ReDim Me.mArrivable(X, Y)
        ReDim Me.mChecked(X, Y)
        ReDim Me.mHLine(X, Y)
        ReDim Me.mVLine(X, Y)
    End Sub

    Public Sub New()
        PrepareMap(5, 5)
        Me.mP1 = New Player(0, 0)
        Me.mP2 = New Player(4, 4)
    End Sub

    Public Sub New(ByVal Length As UInteger)
        PrepareMap(Length, Length)
        Me.mP1 = New Player(0, 0)
        Me.mP2 = New Player(Length - 1, Length - 1)
    End Sub

    Public Sub New(ByVal XLength As UInteger, ByVal YLength As UInteger)
        PrepareMap(XLength, YLength)
        Me.mP1 = New Player(0, 0)
        Me.mP2 = New Player(XLength - 1, YLength - 1)
    End Sub

    Public Sub New(ByVal xP1 As UInteger, ByVal yP1 As UInteger, ByVal xP2 As UInteger, ByVal yP2 As UInteger)
        PrepareMap(5, 5)
        If xP1 >= 5 Or
           xP2 >= 5 Or
           yP1 >= 5 Or
           yP2 >= 5 Then
            Err().Raise(vbObjectError + 513,, "非法位置")
        End If
        Me.mP1 = New Player(xP1, yP1)
        Me.mP2 = New Player(xP2, yP2)
    End Sub

    Public Sub New(ByVal Length As UInteger, ByVal xP1 As UInteger, ByVal yP1 As UInteger, ByVal xP2 As UInteger, ByVal yP2 As UInteger)
        PrepareMap(Length, Length)
        If xP1 >= Length Or
           xP2 >= Length Or
           yP1 >= Length Or
           yP2 >= Length Then
            Err().Raise(vbObjectError + 513,, "非法位置")
        End If
        Me.mP1 = New Player(xP1, yP1)
        Me.mP2 = New Player(xP2, yP2)
    End Sub

    Public Sub New(ByVal XLength As UInteger, ByVal YLength As UInteger, ByVal xP1 As UInteger, ByVal yP1 As UInteger, ByVal xP2 As UInteger, ByVal yP2 As UInteger)
        PrepareMap(XLength, YLength)
        If xP1 >= XLength Or
           xP2 >= XLength Or
           yP1 >= YLength Or
           yP2 >= YLength Then
            Err().Raise(vbObjectError + 513,, "非法位置")
        End If
        Me.mP1 = New Player(xP1, yP1)
        Me.mP2 = New Player(xP2, yP2)
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


End Class
