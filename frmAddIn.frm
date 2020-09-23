VERSION 5.00
Begin VB.Form frmAddIn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TabIndex Manager"
   ClientHeight    =   2745
   ClientLeft      =   2175
   ClientTop       =   1890
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   405
      Left            =   3375
      TabIndex        =   3
      Top             =   1612
      Width           =   405
   End
   Begin VB.CommandButton cmdDiminuiTabIndex 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3375
      TabIndex        =   2
      Top             =   1125
      Width           =   405
   End
   Begin VB.CommandButton cmdAumentaTabIndex 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3375
      TabIndex        =   1
      Top             =   727
      Width           =   405
   End
   Begin VB.ListBox lbxItems 
      Height          =   2595
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   3210
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variável de Controle de Alterações
Private mAlterado As Boolean

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub cmdAumentaTabIndex_Click()
Dim strTxt As String, lngItemData As Long
Dim lngIndex As Long, lngAux As Long

    If lbxItems.ListCount > 1 And lbxItems.ListIndex < lbxItems.ListCount - 1 Then
        'Debug.Print "Aumentando TabIndex ..."
        lngIndex = lbxItems.ListIndex
        strTxt = lbxItems.List(lngIndex)
        lngItemData = lbxItems.ItemData(lngIndex)
        
        lbxItems.RemoveItem lngIndex
        lbxItems.ItemData(lngIndex) = lbxItems.ItemData(lngIndex) - 1
        
        lbxItems.AddItem strTxt, lngIndex + 1
        lbxItems.ItemData(lbxItems.NewIndex) = lngItemData + 1
        lbxItems.ListIndex = lbxItems.NewIndex
        
        mAlterado = True
    End If
End Sub

Private Sub cmdDiminuiTabIndex_Click()
Dim strTxt As String, lngItemData As Long
Dim lngIndex As Long, lngAux As Long

    'If lbxItems.ListCount > 1 And lbxItems.ListIndex < lbxItems.ListCount - 1 Then
    If lbxItems.ListCount > 1 And lbxItems.ListIndex > 0 Then
        'Debug.Print "Diminuindo TabIndex ..."
        
        lngIndex = lbxItems.ListIndex
        strTxt = lbxItems.List(lngIndex)
        lngItemData = lbxItems.ItemData(lngIndex)
        
        lbxItems.RemoveItem lngIndex
        lbxItems.ItemData(lngIndex - 1) = lbxItems.ItemData(lngIndex - 1) + 1
        
        lbxItems.AddItem strTxt, lngIndex - 1
        lbxItems.ItemData(lbxItems.NewIndex) = lngItemData - 1
        lbxItems.ListIndex = lbxItems.NewIndex
        
        mAlterado = True
    End If
End Sub

Private Sub cmdOK_Click()
Dim oFrm As VBForm
Dim oCtl As VBControl
Dim lngCount As Long
'Dim lngIndex As Long
Dim strAux As String
Dim lngAux As Long
Dim strNome As String

    If Len(lbxItems.Tag) = 0 And mAlterado Then
        Set oFrm = VBInstance.ActiveVBProject.VBComponents.Item(1).Designer
        For lngCount = 0 To lbxItems.ListCount - 1
            strAux = lbxItems.List(lngCount)
            strAux = Trim$(Mid$(strAux, InStr(1, strAux, ":") + 1, Len(strAux)))

            'Verifica se não é um array de objetos ...
            lngAux = InStr(1, strAux, "(")
            If lngAux > 0 Then
                strNome = Left$(strAux, lngAux - 1)
                lngAux = CLng(Mid$(strAux, lngAux + 1, (InStr(lngAux, strAux, ")") - (lngAux + 1))))
                For Each oCtl In oFrm.VBControls
                    If oCtl.Properties("Index") = lngAux And oCtl.ControlObject.Name = strNome Then
                        oCtl.ControlObject.TabIndex = lbxItems.ItemData(lngCount)
                        Exit For
                    End If
                Next
            Else
                oFrm.VBControls(strAux).ControlObject.TabIndex = lbxItems.ItemData(lngCount)
            End If
        Next
        
    End If

    Set oFrm = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
Dim oFrm As VBForm
Dim oCtl As VBControl
Dim lngAux As Long
Dim strAux As String

    mAlterado = False

    lbxItems.Visible = False
    lbxItems.Clear
    lbxItems.Tag = ""
    
    If Not VBInstance Is Nothing Then
        If Not VBInstance.ActiveVBProject Is Nothing Then
            'Set oFrm = VBInstance.ActiveVBProject.VBComponents.Item(1).Designer
            Set oFrm = VBInstance.SelectedVBComponent.Designer
            lngAux = Len(CStr(oFrm.VBControls.Count))
            If oFrm.VBControls.Count > 0 Then
                For Each oCtl In oFrm.VBControls
                    If fl_objHaveTabIndex(oCtl.ControlObject) Then
                        strAux = Format$(oCtl.ControlObject.TabIndex, String$(lngAux, "0")) & ": " & Trim$(oCtl.ControlObject.Name)
                        If oCtl.Properties("Index") > -1 Then
                            strAux = strAux & "(" & oCtl.Properties("Index") & ")"
                        End If
                        
                        lbxItems.AddItem strAux
                        'lbxItems.AddItem oCtl.ControlObject.Name
                        lbxItems.ItemData(lbxItems.NewIndex) = oCtl.ControlObject.TabIndex
                    End If
                Next
            Else
                lbxItems.AddItem "(Nenhum objeto no projeto.)"
                lbxItems.Tag = "VAZIO"
                cmdAumentaTabIndex.Enabled = False
                cmdDiminuiTabIndex.Enabled = False
            End If
        Else
            lbxItems.AddItem "(Nenhum projeto ativo.)"
            lbxItems.Tag = "VAZIO"
            cmdAumentaTabIndex.Enabled = False
            cmdDiminuiTabIndex.Enabled = False
        End If
    End If
    
    Set oFrm = Nothing
    Set oCtl = Nothing

    lbxItems.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set frmAddIn = Nothing
End Sub

Private Function fl_objHaveTabIndex(obj As Object) As Boolean
Dim lngAux As Long

    On Error GoTo Erro_Handle
    lngAux = obj.TabIndex
    fl_objHaveTabIndex = Not Err
    Exit Function
    
Erro_Handle:
    fl_objHaveTabIndex = False
End Function
