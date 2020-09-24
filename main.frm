VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "    Vertex Mesh Deformation [Linear Blend Skinning for Character Animation]"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   673
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLeaves 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000C0C0&
      Height          =   1065
      Left            =   6705
      Picture         =   "main.frx":0442
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   226
      TabIndex        =   17
      Top             =   4815
      Width           =   3390
      Begin VB.Label lblContact 
         BackStyle       =   0  'Transparent
         Caption         =   "vbinterface@gmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   765
         Width           =   2130
      End
      Begin VB.Label lblWish 
         BackStyle       =   0  'Transparent
         Caption         =   "I would like to work on 2D/3D graphics simulation programs as a freelancer."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   645
         Left            =   1080
         TabIndex        =   18
         Top             =   45
         Width           =   2130
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00444444&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   0
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   11
      Top             =   315
      Width           =   6675
   End
   Begin VB.Frame framTop 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6705
      Begin VB.HScrollBar hsbWorld 
         Height          =   105
         Index           =   0
         Left            =   2820
         Max             =   360
         Min             =   -360
         TabIndex        =   9
         Top             =   0
         Width           =   3930
      End
      Begin VB.HScrollBar hsbWorld 
         Height          =   105
         Index           =   1
         Left            =   2820
         Max             =   360
         Min             =   -360
         TabIndex        =   8
         Top             =   90
         Width           =   3930
      End
      Begin VB.HScrollBar hsbWorld 
         Height          =   105
         Index           =   2
         Left            =   2820
         Max             =   360
         Min             =   -360
         TabIndex        =   7
         Top             =   180
         Width           =   3930
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   " WorldView :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2850
      End
   End
   Begin VB.Frame framBottom 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5610
      Width           =   6690
      Begin VB.HScrollBar hsb 
         Height          =   105
         Index           =   0
         Left            =   2820
         Max             =   360
         Min             =   -360
         TabIndex        =   3
         Top             =   0
         Width           =   3885
      End
      Begin VB.HScrollBar hsb 
         Height          =   105
         Index           =   1
         Left            =   2820
         Max             =   360
         Min             =   -360
         TabIndex        =   2
         Top             =   90
         Width           =   3885
      End
      Begin VB.HScrollBar hsb 
         Height          =   105
         Index           =   2
         Left            =   2820
         Max             =   360
         Min             =   -360
         TabIndex        =   1
         Top             =   180
         Width           =   3885
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   " Bone Rotation :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1395
         TabIndex        =   5
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H0080C0FF&
         Caption         =   " status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1410
      End
   End
   Begin VB.Label lblInfo2 
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":5580
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1050
      Left            =   6840
      TabIndex        =   16
      Top             =   3465
      Width           =   3030
   End
   Begin VB.Label lblInfo1 
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":5616
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   915
      Left            =   6840
      TabIndex        =   15
      Top             =   2520
      Width           =   3030
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":56A3
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   915
      Left            =   6840
      TabIndex        =   14
      Top             =   1530
      Width           =   3030
   End
   Begin VB.Label lblMiddle 
      BackStyle       =   0  'Transparent
      Caption         =   "Use the scroll bars [Top] to view the vertex mesh in World Space."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   780
      Left            =   6885
      TabIndex        =   13
      Top             =   945
      Width           =   2985
   End
   Begin VB.Label lblTop 
      BackStyle       =   0  'Transparent
      Caption         =   "Use the scroll bars [Bottom] for rotating the bone in 3 Dimensions to deform the mesh at bone joints."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   780
      Left            =   6885
      TabIndex        =   12
      Top             =   135
      Width           =   2985
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'REQUIRES DIRECTX8 REFERENCE FOR MATH CALCULATIONS
    '--------------------------------------------------------------
    Dim Vertex(20, 14) As D3DVECTOR4, tVertex(20, 14) As D3DVECTOR4
    Dim Vec0 As D3DVECTOR4, Vec1 As D3DVECTOR4
    Dim trV0 As D3DVECTOR4, trV1 As D3DVECTOR4
    Dim LocVec0 As D3DVECTOR4, LocVec1 As D3DVECTOR4
    Dim tempVector As D3DVECTOR4, trVector As D3DVECTOR4
    Dim WorldMat As D3DMATRIX, Mat0 As D3DMATRIX, Mat1 As D3DMATRIX
    Dim BoneID As Integer
    Dim DiscWeight As Single, VertexWeight(20) As Single
    Dim Radius As Single
    Const PIBY180 = 3.14 / 180
    Dim CenX As Single, CenY As Single, CenZ As Single
    Dim Disc As Integer, V As Integer
Private Sub Form_Load()
    CenX = pic.ScaleWidth / 2: CenY = pic.ScaleHeight / 2
    Vec0.Y = -80: Vec1.Y = 80: trV0 = Vec0: trV1 = Vec1
    Radius = 20
    LoadMesh
    D3DXMatrixIdentity WorldMat
    D3DXMatrixIdentity Mat0: D3DXMatrixIdentity Mat1
    BoneID = 1
    hsb(2).Value = -45
End Sub
Sub LoadMesh()
    'create mesh and add weights
    Dim XL As Single, YL As Single, ZL As Single
    Dim Angle As Single, IncrAngle As Integer
    Dim vertWeight()
    YL = -88
    IncrAngle = 360 / 14
    For Disc = 0 To 20
        YL = YL + 8
        Angle = 0
        For V = 0 To 13
            Vertex(Disc, V).X = Sin(Angle * PIBY180) * Radius
            Vertex(Disc, V).Y = YL
            Vertex(Disc, V).z = Cos(Angle * PIBY180) * Radius
            Angle = Angle + IncrAngle
        Next V
    Next Disc
    vertWeight = Array(1, 0.995, 0.98, 0.96, 0.94, 0.92, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.1, 0.08, 0.06, 0.04, 0.02, 0.009, 0)
    For V = 0 To UBound(vertWeight)
        VertexWeight(V) = vertWeight(V)
    Next V
    Erase vertWeight
End Sub
Private Sub hsb_Change(Index As Integer)
    hsb_Scroll 0
End Sub
Private Sub hsb_Scroll(Index As Integer)
    pic.Cls
    TransformBoneAndDisplayMesh BoneID, hsb(0).Value, hsb(1).Value, hsb(2).Value
    DoEvents
End Sub
Private Sub hsbWorld_Scroll(Index As Integer)
    D3DXMatrixRotationYawPitchRoll WorldMat, hsbWorld(1).Value * PIBY180, hsbWorld(0).Value * PIBY180, hsbWorld(2).Value * PIBY180
    pic.Cls
    TransformBoneAndDisplayMesh BoneID, hsb(0).Value, hsb(1).Value, hsb(2).Value
    DoEvents
End Sub
Sub TransformBoneAndDisplayMesh(ByVal BoneID As Integer, ByVal RX As Single, RY As Single, RZ As Single)
    If BoneID = 0 Then
        D3DXMatrixRotationYawPitchRoll Mat0, RY * PIBY180, RX * PIBY180, RZ * PIBY180
    Else
        D3DXMatrixRotationYawPitchRoll Mat1, RY * PIBY180, RX * PIBY180, RZ * PIBY180
    End If
    For Disc = 0 To 20
        DiscWeight = VertexWeight(Disc)
        For V = 0 To 13
            D3DXVec4Transform LocVec0, Vertex(Disc, V), Mat0
            D3DXVec4Transform LocVec1, Vertex(Disc, V), Mat1
            tempVector.X = (DiscWeight * LocVec0.X) + ((1 - DiscWeight) * LocVec1.X)
            tempVector.Y = (DiscWeight * LocVec0.Y) + ((1 - DiscWeight) * LocVec1.Y)
            tempVector.z = (DiscWeight * LocVec0.z) + ((1 - DiscWeight) * LocVec1.z)
            D3DXVec4Transform trVector, tempVector, WorldMat
            tVertex(Disc, V) = trVector
            pic.PSet (trVector.X + CenX, trVector.Y + CenY), vbWhite
        Next V
    Next Disc
    D3DXVec4Transform trV0, Vec0, Mat0
    D3DXVec4Transform trV1, Vec1, Mat1
    D3DXVec4Transform trVector, trV0, WorldMat
    pic.Line (trVector.X + CenX, trVector.Y + CenY)-(CenX, CenY), vbWhite
    D3DXVec4Transform trVector, trV1, WorldMat
    pic.Line (trVector.X + CenX, trVector.Y + CenY)-(CenX, CenY), vbWhite
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Erase Vertex
    Erase tVertex
    Erase VertexWeight
End Sub
