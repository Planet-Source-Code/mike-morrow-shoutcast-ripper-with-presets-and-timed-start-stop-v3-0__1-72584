Attribute VB_Name = "RoundCtrls"
Option Explicit

'+----------------------------------------------------+
'| 3D Rounded Controls 1.1 William W. October 9, 2009 |
'|           Contact: Me.theuser@yahoo.com            |
'+----------------------------------------------------+
'|            Free for NON-Commercial uses            |
'+----------------------------------------------------+
'| Place 'DO_CtrlOutline me' in the form load event   |
'| of any form you want the controls rounded in       |
'| If you don't want a control to be rounded          |
'| then put -1 in the tag property of the control     |
'+----------------------------------------------------+

'+-------------------------[Usage Reference]------------------------------+
'| DO_ColorCtrlOutline(ctrl As Control, lOutline As Long, lShadow As Long)|
'|  Changes the outline and shadow colors of the specified control        |
'|  to keep the current color supply -1 for the color                     |
'+------------------------------------------------------------------------+
'| DO_ColorCtrlOutlineALL(frm as form, lOutline As Long, lShadow As Long) |
'|  Changes the outline and shadow colors of all controls on the          |
'|  specified control form                                                |
'|  to keep the current color supply -1 for the color                     |
'+------------------------------------------------------------------------+
'| DO_CtrlOutline(frm As Form,Optional lOutline As Long, _                |
'|                            Optional lShadow As Long)                   |
'|  Rounds all controls on the specified form except those not supported  |
'|  and controls with a -1 Tag                                            |
'|  supplying color arguments overrides the machines color scheme         |
'+------------------------------------------------------------------------+
'| DO_HideShowCtrlOutline(ctrl As Control, Show As Boolean)               |
'|  Hide or Show the specified control outline                            |
'+------------------------------------------------------------------------+
'| DO_HideShowCtrlOutlineALL(frm As Form, Show As Boolean)                |
'|  Hide or Show the outline for every control on the specified form      |
'+------------------------------------------------------------------------+
'| DO_RemoveCtrlOutlines(frm As Object)                                   |
'|  Removes rounding controls and all rounding applied                    |
'|  to the original controls themselves of the specified form             |
'|    :Note: it is not required that you remove controls                  |
'|    before unloading the form                                           |
'+------------------------------------------------------------------------+
'| DO_UpdateCtrlOutline(ctrl As Control,Optional lOutline As Long, _      |
'|                                      Optional lShadow As Long)         |
'|  Updates the size, position, color, and visibility                     |
'|  of the specified control outline based on                             |
'|  the control it corresponds to.                                        |
'|   :Notes: Will round a control even if it hasn't                       |
'|   been rounded before. if you are only changing colors do not          |
'|   use this as its 2x slower than DO_ColorCtrlOutline                   |
'+------------------------------------------------------------------------+

Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
      ByVal X1 As Long, _
      ByVal Y1 As Long, _
      ByVal X2 As Long, _
      ByVal Y2 As Long, _
      ByVal X3 As Long, _
      ByVal Y3 As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" ( _
      ByVal hwnd As Long, _
      ByVal hRgn As Long, _
      ByVal bRedraw As Long) As Long


Public Sub DO_ColorCtrlOutline(ctrl As Control, ByVal lOutline As Long, ByVal lShadow As Long)

 'changes the outline and shadow colors for the outline of the specified control

  On Error GoTo Error_Handling
  Dim cIndex As String
  Dim Obj As Object

   cIndex = ctrl.Index

   If InStr(1, ctrl.Name, "cOutline") + InStr(1, ctrl.Name, "cShadow") + InStr(1, ctrl.Name, _
      "cLabel") <> 0 Or ctrl.Tag = -1 Then Exit Sub

   Set Obj = ctrl.Parent.Controls.Item(ctrl.Name & "_cShadow" & cIndex)

   If lShadow <> -1 And Obj Is Nothing = False Then
      If TypeName(ctrl) = "Frame" Then lShadow = lOutline
      Obj.BorderColor = lShadow
      Obj.BackColor = lShadow
      'Obj.ZOrder 0
      Set Obj = Nothing
   End If

   Set Obj = ctrl.Parent.Controls.Item(ctrl.Name & "_cOutline" & cIndex)

   If Obj Is Nothing = False Then
      If lShadow <> -1 Then Obj.BackColor = lShadow
      If lOutline <> -1 Then Obj.BorderColor = lOutline
      'Obj.ZOrder 0
      Set Obj = Nothing
   End If

   If TypeName(ctrl) = "Frame" Then
      Set Obj = ctrl.Parent.Controls.Item(ctrl.Name & "_cLabel" & cIndex)

      If Obj Is Nothing = False Then
         Obj.ForeColor = ctrl.ForeColor
         Obj.ZOrder 0
         Set Obj = Nothing
      End If

   End If
   ctrl.ZOrder 0 'front
   Exit Sub
Error_Handling:
   If Error_Handler(err) Then Resume Next

End Sub

Public Sub DO_ColorCtrlOutlineALL(frm As Form, lOutline As Long, lShadow As Long)

   'changes outline and shadow colors of all
   'controls on the specified form
  Dim a As Long

   For a = frm.Controls.Count - 1 To 0 Step -1
      DO_ColorCtrlOutline frm.Controls(a), lOutline, lShadow
   Next a

End Sub

Public Sub DO_CtrlOutline(frm As Form, _
                          Optional lOutline As Long = vbButtonText, _
                          Optional lShadow As Long = vb3DDKShadow)

 'Outlines all controls on the specified form that
 'don't have a tag of -1 and are of a recognized type

  Dim a As Long
  Dim sShow As Integer

  If frm.Visible = False Then
    sShow = -1
  Else
    sShow = 1
  End If
  
  'Returns control ZOrder from low to high
  
  For a = frm.Controls.Count - 1 To 0 Step -1
    OutlineCtrl frm.Controls(a), lOutline, lShadow, sShow, 0
  Next a

End Sub

Public Sub DO_HideShowCtrlOutline(ctrl As Control, Show As Boolean)

  'hides or shows the outline and shadow for the specified control
  On Error GoTo Error_Handling
  Dim cIndex As String
  Dim Obj As Object

   If InStr(1, ctrl.Name, "cOutline", vbBinaryCompare) + InStr(1, ctrl.Name, "cShadow", _
      vbBinaryCompare) + InStr(1, ctrl.Name, "cLabel", vbBinaryCompare) <> 0 Then Exit Sub
   cIndex = ctrl.Index

   Set Obj = ctrl.Parent.Controls.Item(ctrl.Name & "_cOutline" & cIndex)
   If Obj Is Nothing = False Then Obj.Visible = Show
   Set Obj = Nothing
   Set Obj = ctrl.Parent.Controls.Item(ctrl.Name & "_cShadow" & cIndex)
   If Obj Is Nothing = False Then Obj.Visible = Show
   Set Obj = Nothing

   If TypeName(ctrl) = "Frame" Then
      Set Obj = ctrl.Parent.Controls.Item(ctrl.Name & "_cLabel" & cIndex)
      If Obj Is Nothing = False Then Obj.Visible = Show
      Set Obj = Nothing

   End If

   ctrl.ZOrder 0
   Exit Sub
Error_Handling:
   If Error_Handler(err) Then Resume Next

End Sub

Public Sub DO_HideShowCtrlOutlineALL(frm As Form, Show As Boolean)

2
   'hide or shows all controls on the specified form
  Dim a As Long

   For a = frm.Controls.Count - 1 To 0 Step -1
      DO_HideShowCtrlOutline frm.Controls(a), Show
   Next a

End Sub

Public Sub DO_RemoveCtrlOutlines(frm As Object)

  'removes all rounding on the specified form and
  'makes the controls 'Standard' again
  On Error GoTo Error_Handling
  DO_HideShowCtrlOutlineALL frm, False
  DoEvents
  Dim ctrl As Control

   For Each ctrl In frm.Controls

      If ctrl.Tag <> -1 Then

         If InStr(1, ctrl.Name, "cOutline", vbBinaryCompare) + InStr(1, ctrl.Name, "cShadow", _
            vbBinaryCompare) + InStr(1, ctrl.Name, "cLabel", vbBinaryCompare) <> 0 Then

            frm.Controls.Remove ctrl
            Set ctrl = Nothing
          Else

            If TypeName(ctrl) = "Label" Then

               If ctrl.Tag = 1 Then
                  ' values were changed so reverse changes

                  With ctrl
                     .Top = .Top - 40
                     .Height = .Height + 40
                     .Width = .Width + 70
                     .Left = .Left - 40
                     .BorderStyle = 1
                     .Tag = 0
                  End With

               End If

            End If

            If TypeName(ctrl) = "Frame" Then

               If Val(ctrl.Tag) > 0 Then
                  ctrl.BorderStyle = ctrl.Tag - 1
                  ctrl.Tag = 0
               End If

            End If

            SetWindowRgn ctrl.hwnd, 0, True 'removes region and Refresh ctrl if visible

         End If
      End If

   Next

   Exit Sub
Error_Handling:
   If Error_Handler(err) Then Resume Next

End Sub

Public Sub DO_UpdateCtrlOutline(ctrl As Control, _
                                Optional lOutline As Long = -1, _
                                Optional lShadow As Long = -1)

   'updates the outlines for the specified control
   'use when you Move or Resize a control
  Dim a As Long
  Dim bClipDiff As Boolean

   OutlineCtrl ctrl, lOutline, lShadow, 1

   For a = ctrl.Parent.Controls.Count - 1 To 0 Step -1

      If TypeName(ctrl.Parent.Controls(a)) = "Frame" Then
         If ctrl.Parent.Controls(a).ClipControls <> ctrl.Parent.Controls(a).Parent.ClipControls _
            Then bClipDiff = True
      End If

   Next a

   If bClipDiff = True Then

      For a = ctrl.Parent.Controls.Count - 1 To 0 Step -1
         If TypeName(ctrl.Parent.Controls(a)) = "Frame" Then DO_HideShowCtrlOutline _
            ctrl.Parent.Controls(a), True
      Next a

      ctrl.ZOrder 0
   End If

End Sub

Private Function Dynamic_AddControl(sType As String, _
                                    sName As String, _
                                    oContainer As Object, _
                                    frm As Form, _
                                    Optional lTop As Long = 0, _
                                    Optional lLeft As Long = 0, _
                                    Optional lHeight As Long = 0, _
                                    Optional lWidth As Long = 0, _
                                    Optional bMove As Boolean = True, _
                                    Optional zOrd As Integer = 0) As Object

   'Returns a reference to a new control or an existing
   'one if it already exists. Totally dynamic :D

   On Error GoTo Error_Handling
   Set Dynamic_AddControl = Nothing

   Set Dynamic_AddControl = frm.Controls.Add(sType, sName, oContainer)

   If Dynamic_AddControl Is Nothing Then
      Set Dynamic_AddControl = frm.Controls.Item(sName)
      Set Dynamic_AddControl.Container = oContainer
      Dynamic_AddControl.Visible = False
      Dynamic_AddControl.ZOrder zOrd
     ''If gbDebugLogic Then Debug.Print mytime() & "Exists! " & Dynamic_AddControl.Name & " Form: " & frm.Name & " Container: " & _
         oContainer.Name & " Coords:(T" & lTop&; ", L" & lLeft & ")"

    Else
     ''If gbDebugLogic Then Debug.Print mytime() & "Created! " & Dynamic_AddControl.Name & " Form: " & frm.Name & " Container: " & _
         oContainer.Name & " Coords:(T" & lTop & ", L" & lLeft & ")"

   End If

   If Dynamic_AddControl Is Nothing = False Then
      If bMove = True Then Dynamic_AddControl.Move lLeft, lTop, lWidth, lHeight
      Dynamic_AddControl.ZOrder 0

    Else
      MsgBox "Add Control FAILED!"
   End If

   Exit Function
Error_Handling:
   If Error_Handler(err) Then Resume Next

End Function

Private Function Error_Handler(err As ErrObject) As Boolean

   Error_Handler = True 'the error is handled
   'Only if the error is
   'un-handled will the return be false

   Select Case err.Number
    Case 0: 'no error
    Case 343, 438, 727, 730, 735:
      '343 Control not in an array
      '438 Property or method doesn't exist
      '727 Control already exists
      '730 Control not found
      '735 Control doesnt exist

    Case Else

      MsgBox "Internal REH1, System error #" & err.Number & " Occured." & vbCrLf & err.Description & vbCrLf & "", 16, _
         "Un-Handled Error in: " & err.Source
      Error_Handler = False 'the error is not handled
   End Select

End Function

Private Sub OutlineCtrl(ctrl As Control, _
                        Optional lOutline As Long = vbButtonText, _
                        Optional lShadow As Long = vb3DDKShadow, _
                        Optional Show As Integer = -1, _
                        Optional ShadowWidth = 0)

  'Show:
  '1= same as control
  '0= False
  '-1= true
  
  '(ShadowWidth) 'negative numbers make the shadow smaller
  'Positive make it larger
  
  'Make control tag property -1 if you don't want shape
  'to be changed
  On Error GoTo Error_Handling

  Const VBShape = "VB.Shape"
  Dim Visible As Boolean

   If InStr(1, ctrl.Name, "cOutline", vbBinaryCompare) + InStr(1, ctrl.Name, "cShadow", _
      vbBinaryCompare) + InStr(1, ctrl.Name, "cLabel", vbBinaryCompare) Or ctrl.Tag = -1 Then Exit _
      Sub

  Dim cIndex As String

  cIndex = ctrl.Index
  
  If Show = 1 Then
    Visible = ctrl.Visible
  ElseIf Show = True Then Visible = True
  Else:
    Visible = False
  End If
  
  Dim ShpOutline As Shape
  Dim ShpShadow As Shape
  Dim FrmLbl As Label
  Dim lEllipse As Long
  
  Set ShpOutline = Nothing
  Set ShpShadow = Nothing
  
  'calculate the corner width and height so we can
  'match it to the rounded rectangle shape control
  
  If ctrl.Height > ctrl.Width Then
    lEllipse = (ctrl.Width / 61.5)
  Else
    lEllipse = (ctrl.Height / 61.5)
  End If
  
  Select Case TypeName(ctrl)
  
   Case "CommandButton"
     If ctrl.Appearance = 1 Then '3d
        SetCtrlRegion ctrl.hwnd, 0, ctrl.Visible, ctrl.Width, ctrl.Height
        Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
           ctrl.Container, ctrl.Parent, ctrl.Top + 15, ctrl.Left + 10, ctrl.Height + 10 + _
           ShadowWidth, ctrl.Width + 10 + ShadowWidth)
  
        Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
           ctrl.Container, ctrl.Parent, ctrl.Top + 5, ctrl.Left + 10, ctrl.Height, ctrl.Width _
           - 15)
        ShpOutline.FillStyle = 0 'solid
        ShpOutline.FillColor = ctrl.BackColor
  
     End If
  
   Case "TextBox"
     '1
     SetCtrlRegion ctrl.hwnd, 6, ctrl.Visible, ctrl.Width, ctrl.Height 'Do RectRegion
     Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top - lEllipse * 2, (ctrl.Left - lEllipse * 2) + 60, _
        (ctrl.Height + lEllipse * 4) + 25 + ShadowWidth, (ctrl.Width + lEllipse * 4) - 40 + _
        ShadowWidth)
  
     Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top - lEllipse * 2, ctrl.Left - lEllipse * 2, _
        ctrl.Height + lEllipse * 4, ctrl.Width + lEllipse * 4)
     '                                +5              -5
     ShpOutline.BackStyle = 1 'opaque
     ShpOutline.FillStyle = 0 'solid
     ShpOutline.FillColor = ctrl.BackColor
  
   Case "Label":
  
     If ctrl.BorderStyle = 1 Or ctrl.Tag = 1 Then
        Dim tmpClr As Long
  
        With ctrl
           .Visible = False
  
           If .Tag = 1 Then
              .Top = .Top - 40
              .Height = .Height + 40
              .Width = .Width + 70
              .Left = .Left - 40
           End If
  
        End With
        Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
           ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left, ctrl.Height + 40 + ShadowWidth, _
           ctrl.Width + 40 + ShadowWidth)
        Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
           ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left, ctrl.Height + 15, ctrl.Width + 15)
        ShpOutline.BackStyle = 0 'transparent
        ShpOutline.FillColor = ctrl.BackColor
        ShpOutline.FillStyle = 0 'Solid
  
        With ctrl
           .Top = .Top + 40
           .Height = .Height - 40
           .Width = .Width - 70
           .Left = .Left + 40
  
           If Val(.Tag) <> 1 Then
              tmpClr = .BackColor
              .AutoSize = False
              .BorderStyle = 0
              .Appearance = 0
              .BackColor = tmpClr
              .Tag = 1
           End If
  
        End With
  
      Else
        Set ShpOutline = Nothing
        Set ShpShadow = Nothing
     End If
  
     ctrl.Visible = True
  
   Case "PictureBox":
     SetCtrlRegion ctrl.hwnd, 3, ctrl.Visible, ctrl.Width, ctrl.Height 'Do RectRegion
     Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top + 5, ctrl.Left + 5, ctrl.Height + 10, ctrl.Width _
        + 10)
  
     Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left, ctrl.Height, ctrl.Width)
  
     ShpOutline.BackStyle = 1 'transparent
  
   Case "CheckBox", "OptionButton":
     SetCtrlRegion ctrl.hwnd, 3, ctrl.Visible, ctrl.Width, ctrl.Height 'Do RectRegion
     Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left - 15, ctrl.Height, ctrl.Width + 15)
     ShpOutline.BackStyle = 0 'transparent
     ShpOutline.FillColor = ctrl.BackColor
     ShpOutline.FillStyle = 0 'Solid
  
   Case "DirListBox", "FileListBox", "ListView", "ListBox", "TreeView", "Image"::
     SetCtrlRegion ctrl.hwnd, 6, ctrl.Visible, ctrl.Width, ctrl.Height 'Do RectRegion
     Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top - 10, ctrl.Left - 10, ctrl.Height + 90 + _
        ShadowWidth, ctrl.Width + 70 + ShadowWidth)
  
     Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top - 10, ctrl.Left - 20, ctrl.Height + 60, _
        ctrl.Width + 50)
  
     ShpOutline.BackStyle = 1 'opaque
     ShpOutline.BorderWidth = 2
     ShpOutline.FillStyle = 0 'solid
  
     If TypeName(ctrl) = "TreeView" Then
        ShpOutline.FillColor = &H80000005 'Window Back color 'Ctrl.BackColor
      Else
        ShpOutline.FillColor = ctrl.BackColor
     End If
  
   Case "DriveListBox", "ComboBox", "ImageCombo":
  
     SetCtrlRegion ctrl.hwnd, 2, ctrl.Visible, ctrl.Width, ctrl.Height + 35 'Do RectRegion
     Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left + 5, ctrl.Height + 25 + ShadowWidth, _
        ctrl.Width + ShadowWidth)
  
     Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left, ctrl.Height, ctrl.Width - 30)
     ShpOutline.BackStyle = 0 'Transparent
     If lOutline <> -1 Then ShpOutline.FillColor = lOutline
     ShpOutline.FillColor = ctrl.BackColor
     ShpOutline.FillStyle = 0 'SOLID
  
   Case "Frame":
  
     If ctrl.Appearance = 1 Then '3D
  
        If Val(ctrl.Tag) > 0 Then
           ctrl.BorderStyle = ctrl.Tag - 1
         Else
           ctrl.Tag = ctrl.BorderStyle + 1
        End If
  
        If ctrl.BorderStyle = 1 Then '3D
           Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
              ctrl.Container, ctrl.Parent, ctrl.Top + 100, ctrl.Left + 10, ctrl.Height - 120, _
              ctrl.Width - 40)
  
           With ShpShadow
              .BackStyle = 0 'transparent
              .FillStyle = 0 'solid
              .FillColor = ctrl.BackColor
              .BorderWidth = 1
              .BorderStyle = 1 'Fixed Single
              .Shape = 4 'Rounded Rectangle
              .BorderColor = ctrl.ForeColor
              '.ZOrder 1 'Back
              .Visible = Visible
           End With
  
           Set ShpShadow = Nothing 'No More Processing
  
           Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, ctrl, _
              ctrl.Parent, 100, 10, ctrl.Height - 120, ctrl.Width - 40)
  
           With ShpOutline
              .BackStyle = 0 'transparent
              .BorderWidth = 1
              .BorderStyle = 1 'Fixed Single
              .Shape = 4 'Rounded Rectangle
              .BorderColor = ctrl.ForeColor
              '.ZOrder 0 'Front
              .Visible = Visible
           End With
  
           Set ShpOutline = Nothing 'No More Processing
  
        End If
  
        If Len(ctrl.Caption) <> 0 And ctrl.BorderStyle = 1 Then
  
           Set FrmLbl = Dynamic_AddControl("VB.Label", ctrl.Name & "_cLabel" & cIndex, ctrl, _
              ctrl.Parent, 0, (ctrl.Width * 0.15), 10, ctrl.Width)
  
           With FrmLbl
              .Appearance = 0
              .Alignment = 0 '0 Left '2 center
              .AutoSize = True
              .BackColor = ctrl.BackColor
              .BackStyle = 1 'Opaque
              .BorderStyle = 0 'None
              .Caption = " " & ctrl.Caption & " "
              .Font = ctrl.Font
              .FontBold = ctrl.FontBold
              .FontItalic = ctrl.FontItalic
              .FontName = ctrl.FontName
              .FontSize = ctrl.FontSize
              .FontStrikethru = ctrl.FontStrikethru
              .FontUnderline = ctrl.FontUnderline
              .ForeColor = ctrl.ForeColor
              '.ZOrder 0 'Front
              .Visible = Visible
           End With
  
           Set FrmLbl = Nothing
        End If
  
        SetCtrlRegion ctrl.hwnd, 4, ctrl.Visible, ctrl.Width, ctrl.Height 'Do RectRegion
  
        'we are stopping processing
        'early because frames have different outline needs
        ctrl.ZOrder 0 'front
        ctrl.BorderStyle = 0 'get rid of existing border
     End If
  
   Case "HScrollBar", "VScrollBar"
     SetCtrlRegion ctrl.hwnd, 5, ctrl.Visible, ctrl.Width, ctrl.Height
     Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left, ctrl.Height + 10 + ShadowWidth, _
        ctrl.Width + 10 + ShadowWidth)
  
     Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top - 10, ctrl.Left - 10, ctrl.Height + 50, _
        ctrl.Width + 50)
  
     ShpOutline.BackStyle = 1 'transparent
  
   Case "Slider", "ProgressBar":
     If ctrl.BorderStyle <> 1 Then Exit Sub
     SetCtrlRegion ctrl.hwnd, 1, ctrl.Visible, ctrl.Width, ctrl.Height
     Set ShpShadow = Dynamic_AddControl(VBShape, ctrl.Name & "_cShadow" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left, ctrl.Height + 30 + ShadowWidth, _
        ctrl.Width + 30 + ShadowWidth)
  
     Set ShpOutline = Dynamic_AddControl(VBShape, ctrl.Name & "_cOutline" & cIndex, _
        ctrl.Container, ctrl.Parent, ctrl.Top, ctrl.Left, ctrl.Height, ctrl.Width)
  
     ShpOutline.BackStyle = 0 'transparent
     ShpOutline.FillStyle = 0 'solid
     ShpOutline.FillColor = &H8000000F 'Button Face
  
  End Select
  
  If ShpShadow Is Nothing = False Then
  
     With ShpShadow
  
        If lShadow <> -1 Then
           .BorderColor = lShadow
           .BackColor = lShadow
        End If
  
        .BorderWidth = 1
        .BorderStyle = 1 'change to 0 to remove drop shadow
        .Shape = 4 'Rounded Rectangle
        .BackStyle = 1 'opaque
        '.ZOrder 0
        .Visible = Visible
     End With
  
  End If
  
  If ShpOutline Is Nothing = False Then
  
     With ShpOutline
        .BorderWidth = 1
        .BorderStyle = 1
        .Shape = 4 'Rounded Rectangle
        If lShadow <> -1 Then .BackColor = lShadow
        If lOutline <> -1 Then .BorderColor = lOutline
        '.ZOrder 0
        .Visible = Visible
     End With
  
     ctrl.ZOrder 0 'Front
  End If
  
  Exit Sub

Error_Handling:
      If Error_Handler(err) Then Resume Next

   End Sub

Private Sub SetCtrlRegion(ByVal hwnd As Long, _
                          ByVal ctrl As Long, _
                          bIsVisible As Boolean, _
                          lWidth As Long, _
                          lHeight As Long, _
                          Optional lTop As Long = 0, _
                          Optional lLeft As Long = 0)

  On Error GoTo Error_Handling
  Dim hRgn As Long
  Dim lEllipse As Long

   If hwnd <> 0 Then
      ' calculate the corner width and height so we can
      'match it to the rounded rectangle shape control

      If lHeight > lWidth Then
         lEllipse = lWidth / 61.25 '4.1
       Else
         lEllipse = lHeight / 61.25 '4.1
      End If

      'put the measurements in the proper Units (Twips)

      'previous (61.5)
      lWidth = lWidth / Screen.TwipsPerPixelX
      lHeight = lHeight / Screen.TwipsPerPixelY

      'If gbDebugLogic Then Debug.Print mytime() & "Ellipse:" & lEllipse
      'If gbDebugLogic Then Debug.Print mytime() & "Height:" & lHeight & " Width:" & lWidth

      'if the width or height of the elipse is less than 2
      'CreateRoundRectRgn acts like CreateRectRgn thus no
      'rounding occurs so we will keep the width and height >= 2
      If lEllipse < 2 Then lEllipse = 2

      Select Case ctrl

       Case 0: ' Buttons
         lLeft = 2
         lTop = 1

       Case 1, 3: 'Pictureboxes
         lLeft = 1
         lTop = 1
         'lEllipse = 2

       Case 2: 'ComboBoxes
         lWidth = lWidth - 2
         lHeight = lHeight - 2
         lLeft = 2
         lTop = 2

       Case 4: 'Frame
         lWidth = lWidth - 1
         lHeight = lHeight - 1
         lLeft = 1
         lTop = 1

       Case 5: 'scrollbars
         lLeft = 0
         lTop = 0

       Case 6: 'boxes
         lWidth = lWidth - 2
         lHeight = lHeight - 2
         lLeft = 3
         lTop = 3
         lEllipse = 0 'dont round the edges

       Case Else
         lWidth = lWidth
         lHeight = lHeight
         lLeft = 0
         lTop = 0
         lEllipse = 0 'dont round the edges

      End Select

      'create a rounded rectangle region
      hRgn = CreateRoundRectRgn(lLeft, lTop, lWidth, lHeight, lEllipse, lEllipse)

      If hRgn <> 0 Then
         'Creates a handle to the Display / Screen

         SetWindowRgn hwnd, hRgn, bIsVisible 'Set new region and Refresh ctrl if visible
         '(Apparently not needed See MSDN)DeleteObject hRgn 'discard handle to region

      End If

   End If 'Hwnd <> 0
   Exit Sub
Error_Handling:
   If Error_Handler(err) Then Resume Next

End Sub

