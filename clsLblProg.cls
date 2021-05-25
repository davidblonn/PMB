VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLblProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' By Adam Waller
' Last Modified:  12/16/05

'Private Const sngOffset As Single = 1.5    ' For Excel
Private Const sngOffset As Single = 15      ' For Access

Private mdblMax As Double   ' max value of progress bar
Private mdblVal As Double   ' current value of progress bar
Private mdblFullWidth As Double ' width of front label at 100%
Private mdblIncSize As Double
Private mblnHideCap As Boolean  ' display percent complete
Private mobjParent As Object    ' parent of back label
Private mlblBack As Access.label     ' existing label for back
Private mlblFront As Access.label   ' label created for front
Private mlblCaption As Access.label ' progress bar caption
Private mlblCurrent As Access.label
Private mstrCurrent As String
Private mdteLastUpdate As Date      ' Time last updated
Private mblnNotSmooth As Boolean    ' Display smooth bar by do events after every update.

' This class displays a progress bar created
' from 3 labels.
' to use, just add a label to your form,
' and use this back label to position the
' progress bar.

Public Sub Initialize(BackLabel As Access.label, FrontLabel As Access.label, _
                        CaptionLabel As Access.label, CurrentLabel As Access.label)
Dim objParent As Object ' could be a form or tab control
Dim frm As Form

    Set mobjParent = BackLabel.Parent
    ' set private variables
    Set mlblBack = BackLabel
    Set mlblFront = FrontLabel
    Set mlblCaption = CaptionLabel
    Set mlblCurrent = CurrentLabel
    mstrCurrent = "Process Started"

    ' set properties for back label
    With mlblBack
        .Visible = True
        .Caption = ""
        .SpecialEffect = 2  ' sunken. Seems to lose when not visible.
    End With

    ' set properties for front label
    With mlblFront
        mdblFullWidth = mlblBack.Width - (sngOffset * 2)
        .Left = mlblBack.Left + sngOffset
        .Top = mlblBack.Top + sngOffset
        .Width = 0
        .Height = mlblBack.Height - (sngOffset * 2)
        .Caption = ""
        .BackColor = 8388608
        .BackStyle = 1
        .Visible = True
    End With

    ' set properties for caption label
    With mlblCaption
        .Left = mlblBack.Left + 2
        .Top = mlblBack.Top + 2
        .Width = mlblBack.Width - 4
        .Height = mlblBack.Height - 4
        .TextAlign = 2 'fmTextAlignCenter
        .BackStyle = 0 'fmBackStyleTransparent
        .Caption = "0%"
        .Visible = Not Me.HideCaption
        .ForeColor = 16777215   ' white
    End With
    
    With mlblCurrent
        '.Left = mlblBack.Left + 2
        '.Top = mlblBack.Top + 2
        '.Width = mlblBack.Width - 4
        '.Height = mlblBack.Height - 4
        .TextAlign = 2 'fmTextAlignCenter
        .BackStyle = 0 'fmBackStyleTransparent
        .Caption = mstrCurrent
        .Visible = Not Me.HideCaption
        .ForeColor = 0   ' white
    End With

End Sub


Private Sub Class_Terminate()

    mlblFront.Visible = False
    mlblCaption.Visible = False
    mlblBack.Visible = False
    mlblCurrent.Visible = False
    
End Sub


Public Property Get Max() As Double
    Max = mdblMax
End Property
Public Property Let Max(ByVal dblMax As Double)
    mdblMax = dblMax
End Property


Public Property Get IncrementSize() As Double
    IncrementSize = mdblIncSize
End Property
Public Property Let IncrementSize(ByVal dblSize As Double)
    mdblIncSize = dblSize
End Property


Public Property Get HideCaption() As Boolean
    HideCaption = mblnHideCap
End Property
Public Property Let HideCaption(ByVal blnHide As Boolean)
    mblnHideCap = blnHide
End Property

Private Sub Update()

    Dim intPercent As Integer
    Dim dblWidth As Double
    'On Error Resume Next
    intPercent = mdblVal * (100 / mdblMax)
    dblWidth = mdblVal * (mdblFullWidth / mdblMax)
    mlblFront.Width = dblWidth
    mlblCaption.Caption = intPercent & "%"
    mlblCurrent.Caption = mstrCurrent
    'mlblFront.Parent.Repaint    ' may not be needed

    ' Use white or black, depending on progress
    If Me.Value > (Me.Max / 2) Then
        mlblCaption.ForeColor = 16777215   ' white
    Else
        mlblCaption.ForeColor = 0  ' black
    End If

    If mblnNotSmooth Then
        If mdteLastUpdate <> Now Then
            ' update every second.
            DoEvents
            mdteLastUpdate = Now
        End If
    Else
        DoEvents
    End If

End Sub

Public Property Get Value() As Double
    Value = mdblVal
End Property
Public Property Let Value(ByVal dblVal As Double)
    'update only if change is => 1%
    If (CInt(dblVal * (100 / mdblMax))) > (CInt(mdblVal * (100 / mdblMax))) Then
        mdblVal = dblVal
        Update
    Else
        mdblVal = dblVal
    End If
End Property

Public Sub Increment(Optional current As String)
    Dim dblVal As Double
    dblVal = Me.Value
    If dblVal < Me.Max Then
        Me.Value = dblVal + 1
        mstrCurrent = current
        'Call Update
    End If
End Sub


Public Sub Clear()
    Call Class_Terminate
End Sub


Private Function ParentForm(ctlControl As Control) As String
    ' returns the name of the parent form
    Dim objParent As Object
    Set objParent = ctlControl
    Do While Not TypeOf objParent Is Form
       Set objParent = objParent.Parent
    Loop
    ' Now we should have the parent form
    ParentForm = objParent.name
End Function

Public Property Get Smooth() As Boolean
    ' Display the progress bar smoothly.
    ' True by default, this property allows the call
    ' to doevents after every increment.
    ' If False, it will only update once per second.
    ' (This may increase speed for fast progresses.)
    '
    ' negative to set default to true
    Smooth = mblnNotSmooth
End Property

Public Property Let Smooth(ByVal IsSmooth As Boolean)
    mblnNotSmooth = Not IsSmooth
End Property



Private Sub LogErr(objErr, strMod, strProc, intLine)
    ' For future use.
End Sub

Public Property Get strCurrent() As String
    strCurrent = mstrCurrent
End Property

Public Property Let strCurrent(ByVal current As String)
    mstrCurrent = current
End Property
