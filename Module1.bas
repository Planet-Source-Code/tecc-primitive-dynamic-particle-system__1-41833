Attribute VB_Name = "Module1"

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const MAX_GRAVOVERTAKE As Long = 33
Public Type nPoint
    X As Long
    Y As Long
    'X and Y coordinates for 2 dimensional
End Type
Public Type Particle
    Life As Long
        'How many frames the particle has rendered
        'through
    Loc As nPoint
        'particle's current location
    Force As nPoint
        'force can be negative on X or Y values,
        'thus an X force of -2 would push a particle
        'to the left. this simplifies the code and
        'makes it easier to understand
    RandomMovement As nPoint
        'the same as force, just random
    GFRC As Boolean
    GFV As Long
    Fall As Long
End Type
Public SourceP As nPoint
Public P_MaxLife As Long
Public FRC As nPoint
Public Parr() As Particle
Public pcol As OLE_COLOR

Public Sub InitParticles(ByVal MaxCount As Long, _
                        ByVal ForceX As Long, _
                        ByVal ForceY As Long, _
                        ByVal SourceX As Long, _
                        ByVal SourceY As Long, _
                        ByVal RNDX As Long, _
                        ByVal RNDY As Long, _
                        ByVal MaxLife As Long _
                        , Optional ByRef PCOLS As OLE_COLOR)
ReDim Parr(MaxCount)
pcol = PCOLS
P_MaxLife = MaxLife
SourceP.X = SourceX
SourceP.Y = SourceY
FRC.X = ForceX
FRC.Y = ForceY
For i = 0 To UBound(Parr)
Randomize
    With Parr(i)
        .Force = FRC
        .Life = Rnd * MaxLife
        .Loc = SourceP
        .Loc.Y = .Loc.Y + (-(Rnd * 20))
        .RandomMovement.X = RNDX
        .RandomMovement.X = RNDY
        .GFRC = False
        .GFV = Rnd * MAX_GRAVOVERTAKE
    End With
Next
End Sub
Public Sub MainLoop()
Dim NextPoint As nPoint
Dim Tcol As Long
Dim Tcol1 As Long
Dim RX As Long
Dim RY As Long
Dim a As Byte
Do
Form1.PicMax.Picture = Form1.Picture1.Image
Form1.PicMax.Cls 'clear the render target
For i = 0 To UBound(Parr)
    With Parr(i)

        If .Life < P_MaxLife Then
            'move the particle according to forces
            RY = IIf(Rnd * 10 <= 5, (Rnd * .RandomMovement.Y), -(Rnd * .RandomMovement.Y))
            RX = IIf(Rnd * 10 <= 5, (Rnd * .RandomMovement.X), -(Rnd * .RandomMovement.X))
            NextPoint.Y = .Loc.Y + .Force.Y + (.Fall / 50) + RY + IIf(Rnd * 10 <= 5, Rnd * 1, -(Rnd * 1))
            NextPoint.X = .Loc.X + .Force.X + RX
            'Check the color of the target point
            Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X, NextPoint.Y)
            If Tcol <> vbBlack And Tcol <> pcol Then
                'if we run into a black pixel, stop
                .Loc = NextPoint
                
            Else
                .Fall = 0
                '.Force.X = 0
                .GFRC = True
                'check the pixel next to the particle
                Select Case .Force.X
                    Case 0
                    'we have no force, so there is no
                    'initial path checking, just take
                    'a random path if both ways are open
                    'or if only one path is available,
                    'take it.
     
                    If RX < 0 Then
                        Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X - 1, .Loc.Y)
                        If Tcol <> vbBlack And Tcol <> pcol Then
                            .Loc.X = .Loc.X - 1
                            
                            .Force.X = -1
                            .GFRC = True
                        Else
                            Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X + 1, .Loc.Y)
                            If Tcol <> vbBlack And Tcol <> pcol Then
                            .Loc.X = .Loc.X + 1
                            .Fall = 0
                            .Loc.Y = .Loc.Y - (Rnd * (.Force.Y * 2))
                            .Force.X = 1
                            .GFRC = True
                            End If
                        End If
                    Else
                        Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X + 1, NextPoint.Y - 2)
                        If Tcol <> vbBlack And Tcol <> pcol Then
                            .Loc.X = .Loc.X + 1
                            .Loc.Y = .Loc.Y - (Rnd * 1)
                            .Fall = 0
                            .Force.X = 1
                            .GFRC = True
                        Else
                            Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X - 1, .Loc.Y)
                            If Tcol <> vbBlack And Tcol <> pcol Then
                            .Loc.X = .Loc.X - 1
                            .Loc.Y = .Loc.Y - (Rnd * 1)
                            .Fall = 0
                            .Force.X = -1
                            .GFRC = True
                            End If
                        End If
                    End If
                    
                    Case Is < 0
                    
                    ' we already have some negative X force,
                    'so check if there is a clear path to the
                    'Left First
                    Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X - 1, .Loc.Y)
                    If Tcol <> vbBlack And Tcol <> pcol Then
                    'ok, we can move to the left
                        .Loc.X = .Loc.X - 1
                        .Loc.Y = .Loc.Y - (Rnd * 1)
                        .Fall = 0
                        .Force.X = -1
                        .GFRC = True
                    Else
                        'we cannot move to the left!, try the right
                        Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X + 1, .Loc.Y)
                        If Tcol <> vbBlack And Tcol <> pcol Then
                            'we can move to the right
                            .Loc.X = .Loc.X + 1
                            .Loc.Y = .Loc.Y - (Rnd * 1)
                            .Fall = 0
                            .Force.X = 1
                            .GFRC = True
                        Else
                            GoTo StuckX: 'were stuck as far as
                                         'left and right go
                        End If
                    End If
                    
                    Case Is > 0
                    'we have positive X force, check for
                    'a clear path to the right first
                    Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X + 1, .Loc.Y)
                    If Tcol <> vbBlack And Tcol <> pcol Then
                    'ok, we can move to the right
                        .Loc.X = .Loc.X + 1
                        .Loc.Y = .Loc.Y - (Rnd * 1)
                        .Fall = 0
                        .Force.X = 1
                        .GFRC = True
                    Else
                        'we cannot move to the Right!, try the left
                        Tcol = GetPixel(Form1.Picture1.hdc, NextPoint.X - 1, .Loc.Y)
                        If Tcol <> vbBlack And Tcol <> pcol Then
                            'we can move to the left
                            .Loc.X = .Loc.X - 1
                            .Loc.Y = .Loc.Y - (Rnd * 1)
                            .Fall = 0
                            .Force.X = -1
                            .GFRC = True
                        Else
                            GoTo StuckX: 'were stuck as far as
                                         'left and right go
                        End If
                    End If
                    
                    
                End Select
            End If
GoTo NotStk:
StuckX:

NotStk:
            'render the particle
            SetPixelV Form1.PicMax.hdc, .Loc.X, .Loc.Y, pcol
            .Life = .Life + 1
            .Fall = .Fall + 1
            If .Loc.Y >= Form1.PicMax.ScaleHeight Then
                GoTo REFs:
            End If
            If .GFRC Then
                If .GFV >= MAX_GRAVOVERTAKE Then
                    .GFRC = False
                    .GFV = Rnd * MAX_GRAVOVERTAKE
                    .Force.X = FRC.X
                Else
                    .GFV = .GFV + 1
                End If
            End If
        Else
            'reset the particle
REFs:
            .Life = 0
            .Loc = SourceP
            .Force = FRC
            .GFRC = False
            .Fall = 0
            .GFV = Rnd * MAX_GRAVOVERTAKE
        End If
    End With
Next

DoEvents 'so we dont freeze
Sleep 10 'limit the frame rate
Loop
End Sub
