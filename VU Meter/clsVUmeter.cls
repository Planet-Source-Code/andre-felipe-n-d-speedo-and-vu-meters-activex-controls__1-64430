VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVUmeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' V1.0.0 - 2004-10-12 by Alain van Hanegem
' E-mail: alain@decos.nl

Option Explicit

Private Const RefSizeX As Double = 60
Private Const RefSizeY As Double = 60

Private Const VUNEEDLE_MODE_SMOOTH As Integer = 0
Private Const VUNEEDLE_MODE_DIRECT As Integer = 1
Private Const VUNEEDLE_MODE_SLOW As Integer = 2

Private Type PictureProperties
    hDC As Long
    hPic As Long
    SizeX As Long
    SizeY As Long
End Type

Private Type VU_Position
    centerX As Long
    centery As Long
End Type

Private Type VU_Scale
    scaleX As Double
    scaleY As Double
End Type

Private Type VU_Font
    fontname As String
    fontsize As Long
    isbold As Boolean
    isitalic As Boolean
    txtformat As String
    Color As Long
End Type

Private Type VU_NeedleAnimation
    diff As Double
    speed As Double
End Type

Private Type VU_Label
    centerpos As VU_Position
    F As VU_Font
End Type

Private Type VU_Clipping
    enabled As Boolean
    y1 As Long
    y2 As Long
End Type

Private Type VU_Needle
    centerpos As VU_Position
    scale As VU_Scale
    needletype As Integer
    Color As Long
    operationmode As Integer
    Value As Double
    newvalue As Double
    A As VU_NeedleAnimation
End Type

Private Type VU_Border
    BorderStyle As Long
    BorderColor As Long
    borderthickness As Long
    startvalue As Double
    endvalue As Double
    valuestepsize As Double
    borderbandstartsize1 As Double
    borderbandstartsize2 As Double
    borderbandendsize1 As Double
    borderbandendsize2 As Double
    startRGB As Long
    endRGB As Long
End Type

Private Type VU_Screws
    enabled As Boolean
    posx1 As Long
    posy1 As Long
    posx2 As Long
    posy2 As Long
End Type

Private Type VU_Ticks
    centerpos As VU_Position
    scale As VU_Scale
    small_amp1 As Double
    small_amp2 As Double
    small_color As Long
    small_tickness As Long
    interval_small As Long
    large_amp1 As Double
    large_amp2 As Double
    large_color As Long
    large_tickness As Long
    interval_large As Long
    min_value As Double
    max_value As Double
    minpos_angle As Long
    maxpos_angle As Long
End Type

Private Type VU_ScaleTxt
    centerpos As VU_Position
    scale As VU_Scale
    interval As Long
    F As VU_Font
End Type

Private Type VUMeter
    pic As PictureProperties
    needle As VU_Needle
    ticks As VU_Ticks
    scale As VU_ScaleTxt
    lbltext As VU_Label
    lblvalue As VU_Label
    clipping As VU_Clipping
    border As VU_Border
    screws As VU_Screws
End Type

Private Const M_PI = 3.14159265358979
Private VUm As VUMeter
Private Knb(0 To 1, 0 To 11, 0 To 11) As Long

Public Sub SetVUDefaults(ByVal cx As Long, ByVal cy As Long)    ' 0 = speedo meter; 1 = VU meter
        Init_Border 2, RGB(0, 0, 0), 1
        Init_BorderFade m_FadeStart, m_FadeEnd, 1, 0.85, 1, 0.7, 1, m_FadeStartColor, m_FadeEndColor
        Init_Ticks cx, cy, (1.5 + (m_ControlSize * 0.5)), (0.75 + (m_ControlSize * 0.25)), m_MinValue, m_MaxValue, m_MinPosAngle, m_MaxPosAngle, m_TicksIntervalSmall, m_TicksIntervalBig, m_TickColorSmall, m_TickColorBig, 1, 3, 0.9, 1, 0.85, 1
        Init_Scale cx, cy, (1.425 + (m_ControlSize * 0.475)), (0.75 + (m_ControlSize * 0.25)), m_TicksIntervalBig, "Arial", 12, False, False, "0", RGB(255, 255, 255)
        Init_Needle cx, cy, (((m_NeedleSize * m_ControlSize * NeedleRef) + 0.25) * 2), ((m_NeedleSize * m_ControlSize * NeedleRef) + 0.25), m_NeedleStyle, 0, m_NeedleColor
        Init_Clipping 1, 0, cy - 10
        Init_Textlabel cx, cy - 25, m_TextFont.Name, m_TextFont.SIZE, m_TextFont.Bold, m_TextFont.Italic, m_TextLabel, m_TextColor
        Init_Valuelabel cx, cy, m_ValueFont.Name, m_ValueFont.SIZE, m_ValueFont.Bold, m_ValueFont.Italic, m_ValueLabel, m_ValueColor
        Init_Screws m_Screw, cx - 70, cy - 15, cx + 70, cy - 15
        SetNeedleValue 0
End Sub

Public Sub Init_Picture(ByVal hDC As Long, ByVal hPic As Long, ByVal SizeX As Long, ByVal SizeY As Long)
    VUm.pic.hDC = hDC
    VUm.pic.hPic = hPic
    VUm.pic.SizeX = SizeX
    VUm.pic.SizeY = SizeY
End Sub

Public Sub Init_Border(ByVal BorderStyle As Long, ByVal BorderColor As Long, ByVal borderthickness As Long)
    VUm.border.BorderStyle = BorderStyle
    VUm.border.BorderColor = BorderColor
    VUm.border.borderthickness = borderthickness
End Sub

Public Sub Init_BorderFade(ByVal startvalue As Double, ByVal endvalue As Double, ByVal valuestepsize As Long, ByVal borderbandstartsize1 As Double, ByVal borderbandstartsize2 As Double, ByVal borderbandendsize1 As Double, ByVal borderbandendsize2 As Double, ByVal startRGB As Long, ByVal endRGB As Long)
    VUm.border.startvalue = startvalue
    VUm.border.endvalue = endvalue
    VUm.border.valuestepsize = valuestepsize
    VUm.border.borderbandstartsize1 = borderbandstartsize1
    VUm.border.borderbandstartsize2 = borderbandstartsize2
    VUm.border.borderbandendsize1 = borderbandendsize1
    VUm.border.borderbandendsize2 = borderbandendsize2
    VUm.border.startRGB = startRGB
    VUm.border.endRGB = endRGB
End Sub

Public Sub Init_Ticks(ByVal centerX As Long, ByVal centery As Long, ByVal scaleX As Double, ByVal scaleY As Double, ByVal min_value As Double, ByVal max_value As Double, ByVal minpos_angle As Long, ByVal maxpos_angle As Long, ByVal tickinterval_small As Long, ByVal tickinterval_large As Long, ByVal smallcolor As Long, ByVal largecolor As Long, ByVal small_thickness As Long, ByVal large_thickness As Long, ByVal small_amp1 As Double, ByVal small_amp2 As Double, ByVal large_amp1 As Double, ByVal large_amp2 As Double)
    VUm.ticks.centerpos.centerX = centerX
    VUm.ticks.centerpos.centery = centery
    VUm.ticks.scale.scaleX = scaleX
    VUm.ticks.scale.scaleY = scaleY
    VUm.ticks.min_value = min_value
    VUm.ticks.max_value = max_value
    VUm.ticks.minpos_angle = minpos_angle
    VUm.ticks.maxpos_angle = maxpos_angle
    VUm.ticks.interval_small = tickinterval_small
    VUm.ticks.interval_large = tickinterval_large
    VUm.ticks.small_color = smallcolor
    VUm.ticks.large_color = largecolor
    VUm.ticks.small_tickness = small_thickness
    VUm.ticks.large_tickness = large_thickness
    VUm.ticks.small_amp1 = small_amp1
    VUm.ticks.small_amp2 = small_amp2
    VUm.ticks.large_amp1 = large_amp1
    VUm.ticks.large_amp2 = large_amp2
End Sub

Public Sub Init_Needle(ByVal centerX As Long, ByVal centery As Long, ByVal scaleX As Double, ByVal scaleY As Double, ByVal needletype As Integer, ByVal operationmode As Integer, ByVal NeedleColor As Long)
    VUm.needle.centerpos.centerX = centerX
    VUm.needle.centerpos.centery = centery
    VUm.needle.scale.scaleX = scaleX
    VUm.needle.scale.scaleY = scaleY
    VUm.needle.needletype = needletype
    VUm.needle.operationmode = operationmode
    VUm.needle.Color = NeedleColor
End Sub

Public Sub Init_Scale(ByVal centerX As Long, ByVal centery As Long, ByVal scaleX As Double, ByVal scaleY As Double, ByVal interval As Long, ByVal fontname As String, ByVal fontsize As Long, ByVal fontisbold As Boolean, ByVal fontisitalic As Boolean, ByVal txtformat As String, ByVal TextColor As Long)
    VUm.scale.centerpos.centerX = centerX
    VUm.scale.centerpos.centery = centery
    VUm.scale.scale.scaleX = scaleX
    VUm.scale.scale.scaleY = scaleY
    VUm.scale.interval = interval
    VUm.scale.F.fontname = fontname
    VUm.scale.F.fontsize = fontsize
    VUm.scale.F.isbold = fontisbold
    VUm.scale.F.isitalic = fontisitalic
    VUm.scale.F.txtformat = txtformat
    VUm.scale.F.Color = TextColor
End Sub

Public Sub Init_Textlabel(ByVal centerX As Long, ByVal centery As Long, ByVal fontname As String, ByVal fontsize As Long, ByVal fontisbold As Boolean, ByVal fontisitalic As Boolean, ByVal txtformat As String, ByVal TextColor As Long)
    VUm.lbltext.centerpos.centerX = centerX
    VUm.lbltext.centerpos.centery = centery
    VUm.lbltext.F.fontname = fontname
    VUm.lbltext.F.fontsize = fontsize
    VUm.lbltext.F.isbold = fontisbold
    VUm.lbltext.F.isitalic = fontisitalic
    VUm.lbltext.F.txtformat = txtformat
    VUm.lbltext.F.Color = TextColor
End Sub

Public Sub Init_Valuelabel(ByVal centerX As Long, ByVal centery As Long, ByVal fontname As String, ByVal fontsize As Long, ByVal fontisbold As Boolean, ByVal fontisitalic As Boolean, ByVal txtformat As String, ByVal TextColor As Long)
    VUm.lblvalue.centerpos.centerX = centerX
    VUm.lblvalue.centerpos.centery = centery
    VUm.lblvalue.F.fontname = fontname
    VUm.lblvalue.F.fontsize = fontsize
    VUm.lblvalue.F.isbold = fontisbold
    VUm.lblvalue.F.isitalic = fontisitalic
    VUm.lblvalue.F.txtformat = txtformat
    VUm.lblvalue.F.Color = TextColor
End Sub

Public Sub Init_Clipping(ByVal turnedon As Boolean, ByVal y1 As Long, ByVal y2 As Long)
    VUm.clipping.enabled = turnedon
    VUm.clipping.y1 = y1
    VUm.clipping.y1 = y2
End Sub

Public Sub Init_Screws(ByVal enabled As Boolean, ByVal posx1 As Long, ByVal posy1 As Long, ByVal posx2 As Long, ByVal posy2 As Long)
    VUm.screws.enabled = enabled
    VUm.screws.posx1 = posx1
    VUm.screws.posy1 = posy1
    VUm.screws.posx2 = posx2
    VUm.screws.posy2 = posy2
End Sub

Public Sub SetNeedleValue(ByVal Value As Double)
    VUm.needle.newvalue = Value
    
    If VUm.needle.operationmode = VUNEEDLE_MODE_DIRECT Then
        VUm.needle.Value = VUm.needle.newvalue
    End If
End Sub

Public Sub SetNeedleValueDirect(ByVal Value As Double)
    VUm.needle.newvalue = Value
    VUm.needle.Value = Value
    VUm.needle.A.speed = 0
    VUm.needle.A.diff = 0
End Sub

Public Function GetNeedleValue() As Double
    GetNeedleValue = VUm.needle.newvalue
End Function

Public Function GetNeedleAnimationValue() As Double
    GetNeedleAnimationValue = VUm.needle.Value
End Function

Public Function AdjustC(ByVal c As Double) As Double
    AdjustC = (c + 90) Mod 360
End Function

Public Function GetXPos(ByVal c As Double, ByVal r As Double) As Double
    GetXPos = (Sin(AdjustC(c) * M_PI / 180) * r)
End Function

Public Function GetYPos(ByVal c As Double, ByVal r As Double) As Double
    GetYPos = (Cos(AdjustC(c) * M_PI / 180) * r)
End Function

Public Sub Draw()
    DrawRectangle VUm.pic.hDC, 0, 0, VUm.pic.SizeX, VUm.pic.SizeY, 1, m_BackColor, m_BackColor
    DrawVUScrews
    DrawVUTicks
    DrawVUScale
    DrawVULabels
    DrawVUNeedle
End Sub

Public Sub AnimationLoop()
    If VUm.needle.operationmode = VUNEEDLE_MODE_SMOOTH Then
        VUm.needle.A.diff = VUm.needle.newvalue - VUm.needle.Value
        
        VUm.needle.A.speed = VUm.needle.A.speed + (VUm.needle.A.diff)
        VUm.needle.A.speed = VUm.needle.A.speed * 0.25
        
        VUm.needle.Value = VUm.needle.Value + VUm.needle.A.speed
    End If
    
    If VUm.needle.operationmode = VUNEEDLE_MODE_SLOW Then
        VUm.needle.A.diff = VUm.needle.newvalue - VUm.needle.Value
        
        VUm.needle.A.speed = VUm.needle.A.speed + (VUm.needle.A.diff / 25)
        VUm.needle.A.speed = VUm.needle.A.speed * 0.85
        
        VUm.needle.Value = VUm.needle.Value + VUm.needle.A.speed
    End If
End Sub

Private Function Interpolate(ByVal a1 As Double, ByVal a2 As Double, ByVal value1 As Double, ByVal value2 As Double, ByVal Value As Double) As Double
    Dim f1 As Double
    Dim f2 As Double
    Dim fc1 As Double
    Dim fc2 As Double
    Dim r As Double
    
    f1 = value1
    f2 = value2
    fc1 = a1
    fc2 = a2
    r = (((fc2 - fc1) / (f2 - f1)) * (Value - f1)) + fc1

    Interpolate = r
End Function

Private Sub DrawVUScrews()
    If VUm.screws.enabled Then
        DrawScrew VUm.pic.hDC, 0, VUm.screws.posx1, VUm.screws.posy1
        DrawScrew VUm.pic.hDC, 1, VUm.screws.posx2, VUm.screws.posy2
    End If
End Sub

Private Sub DrawVULabels()
    Dim x1 As Long, y1 As Long
    Dim escapement As Long
    Dim textsize As Long
    
    textsize = VUm.lbltext.F.fontsize
    x1 = VUm.lbltext.centerpos.centerX
    y1 = VUm.lbltext.centerpos.centery - (textsize \ 2)
    escapement = 0
    PlotRotatedText VUm.pic.hDC, VUm.lbltext.F.txtformat, x1, y1, VUm.lbltext.F.Color, RGB(0, 0, 0), RGB(0, 0, 0), False, False, VUm.lbltext.F.fontname, textsize, IIf(VUm.lbltext.F.isbold, FW_BOLD, FW_NORMAL), escapement, VUm.lbltext.F.isitalic, False, False, TA_CENTER

    textsize = VUm.lblvalue.F.fontsize
    x1 = VUm.lblvalue.centerpos.centerX
    y1 = VUm.lblvalue.centerpos.centery - (textsize \ 2)
    escapement = 0
    PlotRotatedText VUm.pic.hDC, Format(VUm.needle.newvalue, VUm.lblvalue.F.txtformat), x1, y1, VUm.lblvalue.F.Color, RGB(0, 0, 0), RGB(0, 0, 0), False, False, VUm.lblvalue.F.fontname, textsize, IIf(VUm.lblvalue.F.isbold, FW_BOLD, FW_NORMAL), escapement, VUm.lblvalue.F.isitalic, False, False, TA_CENTER
End Sub

Public Sub DeactivateClippingRegion()
    Dim p(0 To 3) As POINTAPI
    
    p(0).x = 0: p(0).y = 0
    p(1).x = VUm.pic.SizeX: p(1).y = 0
    p(2).x = VUm.pic.SizeX: p(2).y = VUm.pic.SizeY
    p(3).x = 0: p(3).y = VUm.pic.SizeY

    BeginPath VUm.pic.hDC
    DrawFilledPolygon VUm.pic.hDC, 4, p, 1, 0, 0
    EndPath VUm.pic.hDC
    
    SelectClipPath VUm.pic.hDC, RGN_COPY
End Sub

Public Sub ActivateClippingRegion()
    Dim p(0 To 3) As POINTAPI
    
    p(0).x = 0: p(0).y = VUm.clipping.y1
    p(1).x = VUm.pic.SizeX: p(1).y = VUm.clipping.y1
    p(2).x = VUm.pic.SizeX: p(2).y = VUm.clipping.y2
    p(3).x = 0: p(3).y = VUm.clipping.y2

    BeginPath VUm.pic.hDC
    DrawFilledPolygon VUm.pic.hDC, 4, p, 1, 0, 0
    EndPath VUm.pic.hDC
    
    SelectClipPath VUm.pic.hDC, RGN_COPY
End Sub

Private Sub DrawVUCenter()
    Dim sx As Long, sy As Long
    
    sx = RefSizeX * VUm.ticks.scale.scaleX * 0.1
    sy = RefSizeY * VUm.ticks.scale.scaleY * 0.1
    DrawEllipse VUm.pic.hDC, VUm.needle.centerpos.centerX - sx, VUm.needle.centerpos.centery - sy, VUm.needle.centerpos.centerX + sx, VUm.needle.centerpos.centery + sy, 1, RGB(0, 0, 0), RGB(0, 0, 0)
End Sub

Private Sub DrawVUNeedle()
    Dim totalticks As Long
    Dim i As Integer
    Dim B As Double
    Dim angle As Double
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, X3 As Long, Y3 As Long, X4 As Long, Y4 As Long
    Dim sx As Long, sy As Long
    Dim P1(0 To 25) As POINTAPI
    Dim p2(0 To 25) As POINTAPI
    Dim totalpoints As Integer

    If VUm.clipping.enabled Then
        ActivateClippingRegion
    End If
    
    angle = Interpolate(VUm.ticks.minpos_angle, VUm.ticks.maxpos_angle, VUm.ticks.min_value, VUm.ticks.max_value, VUm.needle.Value)
    
    If VUm.needle.needletype < 3 Then  ' thin line
        DrawVUCenter
        
        x1 = VUm.needle.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.needle.scale.scaleX * 0.95)
        y1 = VUm.needle.centerpos.centery + GetYPos(angle, RefSizeY * VUm.needle.scale.scaleY * 0.95)
        
        x2 = VUm.needle.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.needle.scale.scaleX * -0.2)
        y2 = VUm.needle.centerpos.centery + GetYPos(angle, RefSizeY * VUm.needle.scale.scaleY * -0.2)
        
    
        Select Case VUm.needle.needletype
            Case 0:
                DrawLine VUm.pic.hDC, x1 + 3, y1 + 3, x2 + 3, y2 + 3, 1, RGB(128, 128, 128)
                DrawLine VUm.pic.hDC, x1, y1, x2, y2, 1, VUm.needle.Color
            Case 1:
                DrawLine VUm.pic.hDC, x1 + 3, y1 + 3, x2 + 3, y2 + 3, 2, RGB(128, 128, 128)
                DrawLine VUm.pic.hDC, x1, y1, x2, y2, 2, VUm.needle.Color
            Case 2:
                DrawLine VUm.pic.hDC, x1 + 3, y1 + 3, x2 + 3, y2 + 3, 4, RGB(128, 128, 128)
                DrawLine VUm.pic.hDC, x1, y1, x2, y2, 4, VUm.needle.Color
        End Select
        
    Else    ' >= 3 are polygons
        If VUm.needle.needletype = 3 Then   ' normal pointy
            totalpoints = 5
            P1(0).x = VUm.needle.centerpos.centerX + GetXPos(angle - 8, RefSizeX * VUm.needle.scale.scaleX * -0.2)
            P1(0).y = VUm.needle.centerpos.centery + GetYPos(angle - 8, RefSizeY * VUm.needle.scale.scaleY * -0.2)
            P1(1).x = VUm.needle.centerpos.centerX + GetXPos(angle + 8, RefSizeX * VUm.needle.scale.scaleX * -0.2)
            P1(1).y = VUm.needle.centerpos.centery + GetYPos(angle + 8, RefSizeY * VUm.needle.scale.scaleY * -0.2)
            P1(2).x = VUm.needle.centerpos.centerX + GetXPos(angle - 90, RefSizeX * VUm.needle.scale.scaleX * 0.04)
            P1(2).y = VUm.needle.centerpos.centery + GetYPos(angle - 90, RefSizeY * VUm.needle.scale.scaleY * 0.04)
            P1(3).x = VUm.needle.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.needle.scale.scaleX * 0.95)
            P1(3).y = VUm.needle.centerpos.centery + GetYPos(angle, RefSizeY * VUm.needle.scale.scaleY * 0.95)
            P1(4).x = VUm.needle.centerpos.centerX + GetXPos(angle + 90, RefSizeX * VUm.needle.scale.scaleX * 0.04)
            P1(4).y = VUm.needle.centerpos.centery + GetYPos(angle + 90, RefSizeY * VUm.needle.scale.scaleY * 0.04)
        End If
        
        If VUm.needle.needletype = 4 Then   ' sharp
            totalpoints = 5
            P1(0).x = VUm.needle.centerpos.centerX + GetXPos(angle - 2, RefSizeX * VUm.needle.scale.scaleX * -0.2)
            P1(0).y = VUm.needle.centerpos.centery + GetYPos(angle - 2, RefSizeY * VUm.needle.scale.scaleY * -0.2)
            P1(1).x = VUm.needle.centerpos.centerX + GetXPos(angle + 2, RefSizeX * VUm.needle.scale.scaleX * -0.2)
            P1(1).y = VUm.needle.centerpos.centery + GetYPos(angle + 2, RefSizeY * VUm.needle.scale.scaleY * -0.2)
            P1(2).x = VUm.needle.centerpos.centerX + GetXPos(angle - 90, RefSizeX * VUm.needle.scale.scaleX * 0.02)
            P1(2).y = VUm.needle.centerpos.centery + GetYPos(angle - 90, RefSizeY * VUm.needle.scale.scaleY * 0.02)
            P1(3).x = VUm.needle.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.needle.scale.scaleX * 0.95)
            P1(3).y = VUm.needle.centerpos.centery + GetYPos(angle, RefSizeY * VUm.needle.scale.scaleY * 0.95)
            P1(4).x = VUm.needle.centerpos.centerX + GetXPos(angle + 90, RefSizeX * VUm.needle.scale.scaleX * 0.02)
            P1(4).y = VUm.needle.centerpos.centery + GetYPos(angle + 90, RefSizeY * VUm.needle.scale.scaleY * 0.02)
        End If
        
        If VUm.needle.needletype = 5 Then   ' block
            totalpoints = 4
            P1(0).x = VUm.needle.centerpos.centerX + GetXPos(angle + 2, RefSizeX * VUm.needle.scale.scaleX * 0.95)
            P1(0).y = VUm.needle.centerpos.centery + GetYPos(angle + 2, RefSizeY * VUm.needle.scale.scaleY * 0.95)
            P1(1).x = VUm.needle.centerpos.centerX + GetXPos(angle - 2, RefSizeX * VUm.needle.scale.scaleX * 0.95)
            P1(1).y = VUm.needle.centerpos.centery + GetYPos(angle - 2, RefSizeY * VUm.needle.scale.scaleY * 0.95)
            P1(2).x = VUm.needle.centerpos.centerX + GetXPos(angle + 12, RefSizeX * VUm.needle.scale.scaleX * -0.2)
            P1(2).y = VUm.needle.centerpos.centery + GetYPos(angle + 12, RefSizeY * VUm.needle.scale.scaleY * -0.2)
            P1(3).x = VUm.needle.centerpos.centerX + GetXPos(angle - 12, RefSizeX * VUm.needle.scale.scaleX * -0.2)
            P1(3).y = VUm.needle.centerpos.centery + GetYPos(angle - 12, RefSizeY * VUm.needle.scale.scaleY * -0.2)
        End If
        
        
        For i = 0 To totalpoints        ' shift for shadow in p2
            p2(i).x = P1(i).x + 3
            p2(i).y = P1(i).y + 3
        Next i
        
        DrawFilledPolygon VUm.pic.hDC, totalpoints, p2(), 1, RGB(128, 128, 128), RGB(128, 128, 128)
        
        DrawVUCenter
            
        DrawFilledPolygon VUm.pic.hDC, totalpoints, P1(), 1, VUm.needle.Color, VUm.needle.Color
    End If
    
    If VUm.clipping.enabled Then
        DeactivateClippingRegion
    End If

End Sub

Private Sub DrawVUTicks()
    Dim totalticks As Long
    Dim i As Integer
    Dim B As Double
    Dim angle As Double, angle1 As Double, angle2 As Double
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim sx As Long, sy As Long
    Dim p(0 To 25) As POINTAPI
    Dim clr As Long
    Dim t As Long
    Dim cR As Long
    Dim cG As Long
    Dim cB As Long
    Dim value1 As Double, value2 As Double
    Dim amp1 As Double, amp2 As Double, amp3 As Double, amp4 As Double
    
    If VUm.border.BorderStyle = 1 Then      ' circle
        sx = RefSizeX * VUm.ticks.scale.scaleX * 1
        sy = RefSizeY * VUm.ticks.scale.scaleY * 1
        DrawEllipseNoFill VUm.pic.hDC, VUm.needle.centerpos.centerX - sx, VUm.needle.centerpos.centery - sy, VUm.needle.centerpos.centerX + sx, VUm.needle.centerpos.centery + sy, VUm.border.borderthickness, VUm.border.BorderColor
    End If
    
    If VUm.border.BorderStyle = 2 Then      ' faded band
        totalticks = (VUm.border.endvalue - VUm.border.startvalue) \ VUm.border.valuestepsize
        For i = 0 To totalticks - 1
            value1 = VUm.border.startvalue + (i * VUm.border.valuestepsize)
            value2 = VUm.border.startvalue + ((i + 1) * VUm.border.valuestepsize)
            
            angle1 = Interpolate(VUm.ticks.minpos_angle, VUm.ticks.maxpos_angle, VUm.ticks.min_value, VUm.ticks.max_value, value1)
            angle2 = Interpolate(VUm.ticks.minpos_angle, VUm.ticks.maxpos_angle, VUm.ticks.min_value, VUm.ticks.max_value, value2)
            
            amp1 = Interpolate(VUm.border.borderbandstartsize1, VUm.border.borderbandendsize1, VUm.border.startvalue, VUm.border.endvalue, value1)
            amp3 = Interpolate(VUm.border.borderbandstartsize2, VUm.border.borderbandendsize2, VUm.border.startvalue, VUm.border.endvalue, value1)
            amp4 = Interpolate(VUm.border.borderbandstartsize2, VUm.border.borderbandendsize2, VUm.border.startvalue, VUm.border.endvalue, value2)
            amp2 = Interpolate(VUm.border.borderbandstartsize1, VUm.border.borderbandendsize1, VUm.border.startvalue, VUm.border.endvalue, value2)
                        
            p(0).x = VUm.ticks.centerpos.centerX + GetXPos(angle1, RefSizeX * VUm.ticks.scale.scaleX * amp1)
            p(0).y = VUm.ticks.centerpos.centery + GetYPos(angle1, RefSizeY * VUm.ticks.scale.scaleY * amp1)
            p(1).x = VUm.ticks.centerpos.centerX + GetXPos(angle1, RefSizeX * VUm.ticks.scale.scaleX * amp3)
            p(1).y = VUm.ticks.centerpos.centery + GetYPos(angle1, RefSizeY * VUm.ticks.scale.scaleY * amp3)
            p(2).x = VUm.ticks.centerpos.centerX + GetXPos(angle2, RefSizeX * VUm.ticks.scale.scaleX * amp4)
            p(2).y = VUm.ticks.centerpos.centery + GetYPos(angle2, RefSizeY * VUm.ticks.scale.scaleY * amp4)
            p(3).x = VUm.ticks.centerpos.centerX + GetXPos(angle2, RefSizeX * VUm.ticks.scale.scaleX * amp2)
            p(3).y = VUm.ticks.centerpos.centery + GetYPos(angle2, RefSizeY * VUm.ticks.scale.scaleY * amp2)
            
            cR = Interpolate(VUm.border.startRGB Mod 256, VUm.border.endRGB Mod 256, VUm.border.startvalue, VUm.border.endvalue, value2)
            cG = Interpolate(VUm.border.startRGB \ 256 Mod 256, VUm.border.endRGB \ 255 Mod 256, VUm.border.startvalue, VUm.border.endvalue, value2)
            cB = Interpolate(VUm.border.startRGB \ 65536 Mod 256, VUm.border.endRGB \ 65536 Mod 256, VUm.border.startvalue, VUm.border.endvalue, value2)
            clr = RGB(cR, cG, cB)
            
            DrawFilledPolygon VUm.pic.hDC, 4, p(), 1, clr, clr
        Next i
    End If
    
    ' small ticks
    totalticks = (VUm.ticks.max_value - VUm.ticks.min_value) \ VUm.ticks.interval_small
    
    For i = 0 To totalticks
        B = VUm.ticks.min_value + (i * VUm.ticks.interval_small)
        angle = Interpolate(VUm.ticks.minpos_angle, VUm.ticks.maxpos_angle, VUm.ticks.min_value, VUm.ticks.max_value, B)
        
        x1 = VUm.ticks.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.ticks.scale.scaleX * VUm.ticks.small_amp1)
        y1 = VUm.ticks.centerpos.centery + GetYPos(angle, RefSizeY * VUm.ticks.scale.scaleY * VUm.ticks.small_amp1)
        
        x2 = VUm.ticks.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.ticks.scale.scaleX * VUm.ticks.small_amp2)
        y2 = VUm.ticks.centerpos.centery + GetYPos(angle, RefSizeY * VUm.ticks.scale.scaleY * VUm.ticks.small_amp2)

        DrawLine VUm.pic.hDC, x1, y1, x2, y2, VUm.ticks.small_tickness, VUm.ticks.small_color
    Next i

    ' large ticks
    totalticks = (VUm.ticks.max_value - VUm.ticks.min_value) \ VUm.ticks.interval_large

    For i = 0 To totalticks
        B = VUm.ticks.min_value + (i * VUm.ticks.interval_large)
        angle = Interpolate(VUm.ticks.minpos_angle, VUm.ticks.maxpos_angle, VUm.ticks.min_value, VUm.ticks.max_value, B)
        
        x1 = VUm.ticks.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.ticks.scale.scaleX * VUm.ticks.large_amp1)
        y1 = VUm.ticks.centerpos.centery + GetYPos(angle, RefSizeY * VUm.ticks.scale.scaleY * VUm.ticks.large_amp1)
        
        x2 = VUm.ticks.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.ticks.scale.scaleX * VUm.ticks.large_amp2)
        y2 = VUm.ticks.centerpos.centery + GetYPos(angle, RefSizeY * VUm.ticks.scale.scaleY * VUm.ticks.large_amp2)

        DrawLine VUm.pic.hDC, x1, y1, x2, y2, VUm.ticks.large_tickness, VUm.ticks.large_color
    Next i
    
End Sub

Private Sub DrawVUScale()
    Dim totalticks As Long
    Dim i As Integer
    Dim B As Double
    Dim angle As Double
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim escapement As Long
    Dim textsize As Long

    totalticks = (VUm.ticks.max_value - VUm.ticks.min_value) \ VUm.scale.interval
    textsize = VUm.scale.F.fontsize

    For i = 0 To totalticks
        B = VUm.ticks.min_value + (i * VUm.scale.interval)
        angle = Interpolate(VUm.ticks.minpos_angle, VUm.ticks.maxpos_angle, VUm.ticks.min_value, VUm.ticks.max_value, B)
        
        x1 = VUm.scale.centerpos.centerX + GetXPos(angle, RefSizeX * VUm.scale.scale.scaleX * 1.15)
        y1 = VUm.scale.centerpos.centery + GetYPos(angle, RefSizeY * VUm.scale.scale.scaleY * 1.15) - (textsize \ 2)

        escapement = 0
        PlotRotatedText VUm.pic.hDC, Format(B, VUm.scale.F.txtformat), x1, y1, RGB(0, 0, 0), RGB(0, 0, 0), RGB(0, 0, 0), False, False, "Arial", textsize, IIf(VUm.scale.F.isbold, FW_BOLD, FW_NORMAL), escapement, VUm.scale.F.isitalic, False, False, TA_CENTER
    Next i
End Sub

Public Function log10(ByVal x As Double) As Double
    If x = 0 Then log10 = 0 Else log10 = Log(x) / Log(10)
End Function

Public Function VU_MouseClick(ByVal x As Long, ByVal y As Long) As Double
    Dim r As Double
    Dim v As Double
        
    Dim rx As Double
    Dim ry As Double
        
    rx = ((x - VUm.needle.centerpos.centerX) / VUm.needle.scale.scaleX) + VUm.needle.centerpos.centerX
    ry = ((y - VUm.needle.centerpos.centery) / VUm.needle.scale.scaleY) + VUm.needle.centerpos.centery
        
    r = CalcNewAngle(VUm.needle.centerpos.centerX, VUm.needle.centerpos.centery, rx, ry)
    If r > 270 Then r = r - 360
    
    v = Interpolate(VUm.ticks.min_value, VUm.ticks.max_value, VUm.ticks.minpos_angle, VUm.ticks.maxpos_angle, r)
    
    If v < VUm.ticks.min_value Then v = VUm.ticks.min_value
    If v > VUm.ticks.max_value Then v = VUm.ticks.max_value
    
    VU_MouseClick = v
End Function

Private Function CalcNewAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
'         90
'     A2  |  A1
'         |
'180 ----------- 0
'         |
'     A3  |  A4
'        270
   
    Dim g As Double
    Dim h As Double
    Dim v As Double
            
    g = 0
    
    v = y2 - y1
    h = x2 - x1
    
    If (h = 0) And (v < 0) Then g = 90
    If (h = 0) And (v > 0) Then g = 270
    If (v = 0) And (h < 0) Then g = 180
    If (v = 0) And (h > 0) Then g = 0
        
    If (h > 0) And (v < 0) Then ' A1
        g = 90 - Abs(Atn(h / v) * (180 / Pi))
    End If
    
    If (h < 0) And (v < 0) Then ' A2
        g = 90 + Abs(Atn(h / v) * (180 / Pi))
    End If
    
    If (h < 0) And (v > 0) Then ' A3
        g = 180 + (90 - Abs(Atn(h / v) * (180 / Pi)))
    End If

    If (h > 0) And (v > 0) Then ' A4
        g = 270 + Abs(Atn(h / v) * (180 / Pi))
    End If

    CalcNewAngle = g
End Function

Public Sub DrawScrew(ByVal hDC As Long, ByVal id As Integer, ByVal posx As Integer, ByVal posy As Integer)
    Dim x As Integer
    Dim y As Integer
    
    For x = 0 To 11
        For y = 0 To 11
            SetPixel hDC, posx + x - 5, posy + y - 5, Knb(id, y, x)
        Next y
    Next x

End Sub

Public Sub InitScrew()
    Knb(0, 0, 0) = 16777215
    Knb(0, 0, 1) = 16777215
    Knb(0, 0, 2) = 16777215
    Knb(0, 0, 3) = 16777215
    Knb(0, 0, 4) = 16711679
    Knb(0, 0, 5) = 16645886
    Knb(0, 0, 6) = 16645886
    Knb(0, 0, 7) = 16711679
    Knb(0, 0, 8) = 16777215
    Knb(0, 0, 9) = 16777215
    Knb(0, 0, 10) = 16777215
    Knb(0, 0, 11) = 16777215
    Knb(0, 1, 0) = 16777215
    Knb(0, 1, 1) = 16777215
    Knb(0, 1, 2) = 16711679
    Knb(0, 1, 3) = 16054519
    Knb(0, 1, 4) = 14673381
    Knb(0, 1, 5) = 14607077
    Knb(0, 1, 6) = 14739176
    Knb(0, 1, 7) = 15133935
    Knb(0, 1, 8) = 16317436
    Knb(0, 1, 9) = 16777215
    Knb(0, 1, 10) = 16777215
    Knb(0, 1, 11) = 16777215
    Knb(0, 2, 0) = 16777215
    Knb(0, 2, 1) = 16777215
    Knb(0, 2, 2) = 16186105
    Knb(0, 2, 3) = 15660275
    Knb(0, 2, 4) = 15857399
    Knb(0, 2, 5) = 16055036
    Knb(0, 2, 6) = 15923708
    Knb(0, 2, 7) = 15726843
    Knb(0, 2, 8) = 14016482
    Knb(0, 2, 9) = 15791349
    Knb(0, 2, 10) = 16777215
    Knb(0, 2, 11) = 16777215
    Knb(0, 3, 0) = 16777215
    Knb(0, 3, 1) = 16514557
    Knb(0, 3, 2) = 16252412
    Knb(0, 3, 3) = 16580607
    Knb(0, 3, 4) = 16383999
    Knb(0, 3, 5) = 16383999
    Knb(0, 3, 6) = 16449535
    Knb(0, 3, 7) = 13753566
    Knb(0, 3, 8) = 10069417
    Knb(0, 3, 9) = 12700877
    Knb(0, 3, 10) = 16382971
    Knb(0, 3, 11) = 16777215
    Knb(0, 4, 0) = 16777215
    Knb(0, 4, 1) = 15659506
    Knb(0, 4, 2) = 16251642
    Knb(0, 4, 3) = 16382714
    Knb(0, 4, 4) = 16646143
    Knb(0, 4, 5) = 15594998
    Knb(0, 4, 6) = 11911105
    Knb(0, 4, 7) = 9016982
    Knb(0, 4, 8) = 11450810
    Knb(0, 4, 9) = 13950686
    Knb(0, 4, 10) = 14739174
    Knb(0, 4, 11) = 16645886
    Knb(0, 5, 0) = 16645886
    Knb(0, 5, 1) = 11383733
    Knb(0, 5, 2) = 9738906
    Knb(0, 5, 3) = 12040890
    Knb(0, 5, 4) = 12502726
    Knb(0, 5, 5) = 9477279
    Knb(0, 5, 6) = 8424589
    Knb(0, 5, 7) = 8687760
    Knb(0, 5, 8) = 9411226
    Knb(0, 5, 9) = 10134178
    Knb(0, 5, 10) = 12699333
    Knb(0, 5, 11) = 16514300
    Knb(0, 6, 0) = 16514300
    Knb(0, 6, 1) = 9411226
    Knb(0, 6, 2) = 7108472
    Knb(0, 6, 3) = 8818832
    Knb(0, 6, 4) = 9148312
    Knb(0, 6, 5) = 9806501
    Knb(0, 6, 6) = 9346204
    Knb(0, 6, 7) = 8227722
    Knb(0, 6, 8) = 8095878
    Knb(0, 6, 9) = 8292487
    Knb(0, 6, 10) = 12041402
    Knb(0, 6, 11) = 16514300
    Knb(0, 7, 0) = 16514556
    Knb(0, 7, 1) = 10793136
    Knb(0, 7, 2) = 9675424
    Knb(0, 7, 3) = 10267819
    Knb(0, 7, 4) = 10662321
    Knb(0, 7, 5) = 10464685
    Knb(0, 7, 6) = 10530477
    Knb(0, 7, 7) = 10464684
    Knb(0, 7, 8) = 9938339
    Knb(0, 7, 9) = 8555401
    Knb(0, 7, 10) = 12633796
    Knb(0, 7, 11) = 16580093
    Knb(0, 8, 0) = 16711679
    Knb(0, 8, 1) = 13621721
    Knb(0, 8, 2) = 11517373
    Knb(0, 8, 3) = 9281438
    Knb(0, 8, 4) = 8820629
    Knb(0, 8, 5) = 9477532
    Knb(0, 8, 6) = 9477532
    Knb(0, 8, 7) = 9214616
    Knb(0, 8, 8) = 8556428
    Knb(0, 8, 9) = 8555657
    Knb(0, 8, 10) = 14739172
    Knb(0, 8, 11) = 16711678
    Knb(0, 9, 0) = 16777215
    Knb(0, 9, 1) = 16251642
    Knb(0, 9, 2) = 11056050
    Knb(0, 9, 3) = 7307389
    Knb(0, 9, 4) = 6780273
    Knb(0, 9, 5) = 6779758
    Knb(0, 9, 6) = 6911344
    Knb(0, 9, 7) = 7240567
    Knb(0, 9, 8) = 8161156
    Knb(0, 9, 9) = 11976636
    Knb(0, 9, 10) = 16448763
    Knb(0, 9, 11) = 16777215
    Knb(0, 10, 0) = 16777215
    Knb(0, 10, 1) = 16777215
    Knb(0, 10, 2) = 16054262
    Knb(0, 10, 3) = 12042429
    Knb(0, 10, 4) = 9608085
    Knb(0, 10, 5) = 8686981
    Knb(0, 10, 6) = 8423811
    Knb(0, 10, 7) = 9476499
    Knb(0, 10, 8) = 12568772
    Knb(0, 10, 9) = 16251641
    Knb(0, 10, 10) = 16777215
    Knb(0, 10, 11) = 16777215
    Knb(0, 11, 0) = 16777215
    Knb(0, 11, 1) = 16777215
    Knb(0, 11, 2) = 16777215
    Knb(0, 11, 3) = 16448763
    Knb(0, 11, 4) = 14014677
    Knb(0, 11, 5) = 14212055
    Knb(0, 11, 6) = 15264744
    Knb(0, 11, 7) = 15725551
    Knb(0, 11, 8) = 16645885
    Knb(0, 11, 9) = 16777215
    Knb(0, 11, 10) = 16777215
    Knb(0, 11, 11) = 16777215

    Knb(1, 0, 0) = 16777215
    Knb(1, 0, 1) = 16777215
    Knb(1, 0, 2) = 16777215
    Knb(1, 0, 3) = 16777215
    Knb(1, 0, 4) = 16514557
    Knb(1, 0, 5) = 16382715
    Knb(1, 0, 6) = 16448507
    Knb(1, 0, 7) = 16580093
    Knb(1, 0, 8) = 16777215
    Knb(1, 0, 9) = 16777215
    Knb(1, 0, 10) = 16777215
    Knb(1, 0, 11) = 16777215
    Knb(1, 1, 0) = 16777215
    Knb(1, 1, 1) = 16777215
    Knb(1, 1, 2) = 16711679
    Knb(1, 1, 3) = 15922679
    Knb(1, 1, 4) = 14739177
    Knb(1, 1, 5) = 14607334
    Knb(1, 1, 6) = 14476776
    Knb(1, 1, 7) = 14739690
    Knb(1, 1, 8) = 15133932
    Knb(1, 1, 9) = 16645886
    Knb(1, 1, 10) = 16777215
    Knb(1, 1, 11) = 16777215
    Knb(1, 2, 0) = 16777215
    Knb(1, 2, 1) = 16711679
    Knb(1, 2, 2) = 15528178
    Knb(1, 2, 3) = 15462900
    Knb(1, 2, 4) = 16121342
    Knb(1, 2, 5) = 16187135
    Knb(1, 2, 6) = 16056319
    Knb(1, 2, 7) = 14609646
    Knb(1, 2, 8) = 10135208
    Knb(1, 2, 9) = 14608102
    Knb(1, 2, 10) = 16646143
    Knb(1, 2, 11) = 16777215
    Knb(1, 3, 0) = 16777215
    Knb(1, 3, 1) = 15725556
    Knb(1, 3, 2) = 14607589
    Knb(1, 3, 3) = 16580607
    Knb(1, 3, 4) = 16646143
    Knb(1, 3, 5) = 16646143
    Knb(1, 3, 6) = 15596540
    Knb(1, 3, 7) = 10267310
    Knb(1, 3, 8) = 11188410
    Knb(1, 3, 9) = 15990269
    Knb(1, 3, 10) = 15857655
    Knb(1, 3, 11) = 16711679
    Knb(1, 4, 0) = 16580350
    Knb(1, 4, 1) = 13621206
    Knb(1, 4, 2) = 13358035
    Knb(1, 4, 3) = 14937067
    Knb(1, 4, 4) = 15199212
    Knb(1, 4, 5) = 15396078
    Knb(1, 4, 6) = 11451585
    Knb(1, 4, 7) = 9609123
    Knb(1, 4, 8) = 12701135
    Knb(1, 4, 9) = 14543080
    Knb(1, 4, 10) = 14081757
    Knb(1, 4, 11) = 16383227
    Knb(1, 5, 0) = 16383227
    Knb(1, 5, 1) = 12240067
    Knb(1, 5, 2) = 10069156
    Knb(1, 5, 3) = 9279640
    Knb(1, 5, 4) = 8423819
    Knb(1, 5, 5) = 9147030
    Knb(1, 5, 6) = 10136240
    Knb(1, 5, 7) = 8622998
    Knb(1, 5, 8) = 7306622
    Knb(1, 5, 9) = 8358795
    Knb(1, 5, 10) = 10463142
    Knb(1, 5, 11) = 15988470
    Knb(1, 6, 0) = 16382971
    Knb(1, 6, 1) = 11845053
    Knb(1, 6, 2) = 9148824
    Knb(1, 6, 3) = 9083034
    Knb(1, 6, 4) = 8622483
    Knb(1, 6, 5) = 10662068
    Knb(1, 6, 6) = 10334388
    Knb(1, 6, 7) = 8952477
    Knb(1, 6, 8) = 8359567
    Knb(1, 6, 9) = 8227208
    Knb(1, 6, 10) = 9674139
    Knb(1, 6, 11) = 15922932
    Knb(1, 7, 0) = 16645886
    Knb(1, 7, 1) = 13423827
    Knb(1, 7, 2) = 8293515
    Knb(1, 7, 3) = 8688789
    Knb(1, 7, 4) = 10794679
    Knb(1, 7, 5) = 11715782
    Knb(1, 7, 6) = 9874090
    Knb(1, 7, 7) = 10136749
    Knb(1, 7, 8) = 9938854
    Knb(1, 7, 9) = 9017237
    Knb(1, 7, 10) = 10529447
    Knb(1, 7, 11) = 16054518
    Knb(1, 8, 0) = 16777215
    Knb(1, 8, 1) = 15857140
    Knb(1, 8, 2) = 8885139
    Knb(1, 8, 3) = 9412510
    Knb(1, 8, 4) = 11847108
    Knb(1, 8, 5) = 8623251
    Knb(1, 8, 6) = 8754834
    Knb(1, 8, 7) = 8096648
    Knb(1, 8, 8) = 7701376
    Knb(1, 8, 9) = 8753807
    Knb(1, 8, 10) = 13555412
    Knb(1, 8, 11) = 16580093
    Knb(1, 9, 0) = 16777215
    Knb(1, 9, 1) = 16711679
    Knb(1, 9, 2) = 14213085
    Knb(1, 9, 3) = 9938080
    Knb(1, 9, 4) = 7701375
    Knb(1, 9, 5) = 6451564
    Knb(1, 9, 6) = 6713964
    Knb(1, 9, 7) = 6384998
    Knb(1, 9, 8) = 7897469
    Knb(1, 9, 9) = 12765895
    Knb(1, 9, 10) = 16383227
    Knb(1, 9, 11) = 16777215
    Knb(1, 10, 0) = 16777215
    Knb(1, 10, 1) = 16777215
    Knb(1, 10, 2) = 16711422
    Knb(1, 10, 3) = 14607328
    Knb(1, 10, 4) = 9871256
    Knb(1, 10, 5) = 8818565
    Knb(1, 10, 6) = 8291965
    Knb(1, 10, 7) = 8686723
    Knb(1, 10, 8) = 11976119
    Knb(1, 10, 9) = 16251127
    Knb(1, 10, 10) = 16777215
    Knb(1, 10, 11) = 16777215
    Knb(1, 11, 0) = 16777215
    Knb(1, 11, 1) = 16777215
    Knb(1, 11, 2) = 16777215
    Knb(1, 11, 3) = 16777215
    Knb(1, 11, 4) = 15922418
    Knb(1, 11, 5) = 14014419
    Knb(1, 11, 6) = 15067364
    Knb(1, 11, 7) = 15659759
    Knb(1, 11, 8) = 16382713
    Knb(1, 11, 9) = 16777215
    Knb(1, 11, 10) = 16777215
    Knb(1, 11, 11) = 16777215
End Sub

Private Sub Class_Initialize()
    InitScrew
End Sub
