VERSION 5.00
Begin VB.UserControl ScalableForm 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ScalableForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Declare Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetCursor Lib "USER32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "USER32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    

Private WithEvents mParentForm As Form
Attribute mParentForm.VB_VarHelpID = -1

Private Const WM_SYSCOMMAND = &H112&
Private Const SC_SIZE_W = &HF001&
Private Const SC_SIZE_E = &HF002&
Private Const SC_SIZE_N = &HF003&
Private Const SC_SIZE_S = &HF006&
Private Const SC_SIZE_NW = &HF004&
Private Const SC_SIZE_NE = &HF005&
Private Const SC_SIZE_SW = &HF007&
Private Const SC_SIZE_SE = &HF008&

Private Const IDC_SIZEWE = 32644
Private Const IDC_SIZENS = 32645
Private Const IDC_SIZENWSE = 32642
Private Const IDC_SIZENESW = 32643
Private Const IDC_ARROW = 32512

Private BaseLeft() As Single
Private BaseTop() As Single
Private BaseWidth() As Single
Private BaseHeight() As Single
Private BaseFontSize() As Single
Private ControlCount As Integer
Public BaseFormWidth As Single
Public BaseFormHeight As Single
Public SetScalePercentage As Double
Private Sub mParentForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const EDGE_SIZE = 300
    Dim hCur As Long

    If X < EDGE_SIZE And Y < EDGE_SIZE Then
        hCur = LoadCursor(0, IDC_SIZENWSE)
    ElseIf X > mParentForm.ScaleWidth - EDGE_SIZE And Y < EDGE_SIZE Then
        hCur = LoadCursor(0, IDC_SIZENESW)
    ElseIf X < EDGE_SIZE And Y > mParentForm.ScaleHeight - EDGE_SIZE Then
        hCur = LoadCursor(0, IDC_SIZENESW)
    ElseIf X > mParentForm.ScaleWidth - EDGE_SIZE And Y > mParentForm.ScaleHeight - EDGE_SIZE Then
        hCur = LoadCursor(0, IDC_SIZENWSE)
    ElseIf X < EDGE_SIZE Or X > mParentForm.ScaleWidth - EDGE_SIZE Then
        hCur = LoadCursor(0, IDC_SIZEWE)
    ElseIf Y < EDGE_SIZE Or Y > mParentForm.ScaleHeight - EDGE_SIZE Then
        hCur = LoadCursor(0, IDC_SIZENS)
    Else
        hCur = LoadCursor(0, IDC_ARROW)
    End If

    SetCursor hCur
End Sub
Private Sub mParentForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const EDGE_SIZE = 300
    Dim hWndForm As Long

    hWndForm = Extender.Container.Hwnd

    If Button = vbLeftButton Then
        ReleaseCapture
        If X < EDGE_SIZE And Y < EDGE_SIZE Then
            PostMessage hWndForm, WM_SYSCOMMAND, SC_SIZE_NW, 0&
        ElseIf X > mParentForm.ScaleWidth - EDGE_SIZE And Y < EDGE_SIZE Then
            PostMessage hWndForm, WM_SYSCOMMAND, SC_SIZE_NE, 0&
        ElseIf X < EDGE_SIZE And Y > mParentForm.ScaleHeight - EDGE_SIZE Then
            PostMessage hWndForm, WM_SYSCOMMAND, SC_SIZE_SW, 0&
        ElseIf X > mParentForm.ScaleWidth - EDGE_SIZE And Y > mParentForm.ScaleHeight - EDGE_SIZE Then
            PostMessage hWndForm, WM_SYSCOMMAND, SC_SIZE_SE, 0&
        ElseIf X < EDGE_SIZE Then
            PostMessage hWndForm, WM_SYSCOMMAND, SC_SIZE_W, 0&
        ElseIf X > mParentForm.ScaleWidth - EDGE_SIZE Then
            PostMessage hWndForm, WM_SYSCOMMAND, SC_SIZE_E, 0&
        ElseIf Y < EDGE_SIZE Then
            PostMessage hWndForm, WM_SYSCOMMAND, SC_SIZE_N, 0&
        ElseIf Y > mParentForm.ScaleHeight - EDGE_SIZE Then
            PostMessage hWndForm, WM_SYSCOMMAND, SC_SIZE_S, 0&
        End If
    End If
End Sub

Private Function ParentForm() As Form
    If Not Extender Is Nothing Then
        If TypeOf Extender.Container Is Form Then
            Set ParentForm = Extender.Container
            If BaseFormWidth = 0 Then BaseFormWidth = ParentForm.Width
            If BaseFormHeight = 0 Then BaseFormHeight = ParentForm.Height
            Exit Function
        Else
            MsgBox "This ScaleWindowControl must be placed directly on a Form.", vbExclamation
            Set ParentForm = Nothing
            End
        End If
    End If
    Set ParentForm = Nothing
End Function



Private Sub UserControl_Show()
On Error Resume Next
Set mParentForm = ParentForm()
If Not mParentForm Is Nothing Then
    If SetScalePercentage = 0 Then
        SetScalePercentage = 80
    End If
    InitResize mParentForm, SetScalePercentage
End If
On Error GoTo 0
End Sub

Private Sub mParentForm_Resize()
    If Not mParentForm Is Nothing Then
    ResizeControls mParentForm
    End If
End Sub

Public Sub InitResize(frm As Form, pct As Double)
    Dim i As Integer
    
    ControlCount = frm.Controls.Count
    ReDim BaseLeft(ControlCount - 1)
    ReDim BaseTop(ControlCount - 1)
    ReDim BaseWidth(ControlCount - 1)
    ReDim BaseHeight(ControlCount - 1)
    ReDim BaseFontSize(ControlCount - 1)

    For i = 0 To ControlCount - 1
        On Error Resume Next
        BaseLeft(i) = frm.Controls(i).Left
        BaseTop(i) = frm.Controls(i).Top
        BaseWidth(i) = frm.Controls(i).Width
        BaseHeight(i) = frm.Controls(i).Height
        If Not frm.Controls(i).Font Is Nothing Then
            BaseFontSize(i) = frm.Controls(i).Font.Size
        End If
        On Error GoTo 0
    Next i

    Dim screenW As Single, screenH As Single
    Dim targetW As Single, targetH As Single
    Dim scaleX As Double, scaleY As Double

    screenW = BaseFormWidth
    screenH = BaseFormHeight

    targetW = screenW * (pct / 100)
    targetH = screenH * (pct / 100)

    frm.Move (screenW - targetW) / 2, (screenH - targetH) / 2, targetW, targetH

    scaleX = targetW / BaseFormWidth
    scaleY = targetH / BaseFormHeight

    For i = 0 To ControlCount - 1
        On Error Resume Next
        With frm.Controls(i)
                .Left = BaseLeft(i) * scaleX
                .Top = BaseTop(i) * scaleY
                .Width = BaseWidth(i) * scaleX
                .Height = BaseHeight(i) * scaleY
            If BaseFontSize(i) > 0 Then
                .Font.Size = BaseFontSize(i) * scaleY
            End If

            If TypeOf frm.Controls(i) Is Image Then
                Dim img As Image
                Set img = frm.Controls(i)
                If Not img.Picture Is Nothing Then
                    frm.Cls
                    frm.PaintPicture img.Picture, img.Left, img.Top, img.Width, img.Height
                End If
            End If
        End With
        On Error GoTo 0
    Next i
End Sub
Public Sub ResizeControls(frm As Form)
    Dim scaleX As Double, scaleY As Double
    Dim i As Integer

    If BaseFormWidth = 0 Or BaseFormHeight = 0 Then Exit Sub

    scaleX = frm.Width / BaseFormWidth
    scaleY = frm.Height / BaseFormHeight
    
    For i = 0 To ControlCount - 1
        On Error Resume Next
        With frm.Controls(i)
        
            .Left = BaseLeft(i) * scaleX
            .Top = BaseTop(i) * scaleY
            .Width = BaseWidth(i) * scaleX
            .Height = BaseHeight(i) * scaleY

            If BaseFontSize(i) > 0 Then
                .Font.Size = BaseFontSize(i) * scaleY
            End If

            If TypeOf frm.Controls(i) Is Image Then
                Dim img As Image
                Set img = frm.Controls(i)
                If Not img.Picture Is Nothing Then
                    frm.Cls
                    frm.PaintPicture img.Picture, img.Left, img.Top, img.Width, img.Height
                End If
            End If
        End With
        On Error GoTo 0
    Next i
End Sub

