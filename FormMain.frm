VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "һ����װ"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11190
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   746
   StartUpPosition =   2  '��Ļ����
   Begin TabDlg.SSTab SSTab 
      Height          =   6975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "��ҳ"
      TabPicture(0)   =   "FormMain.frx":E8CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelEquip"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabelTip"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelHwnd"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameBag"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TimerHotkey"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrameEquip(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FrameEquip(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FrameEquip(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "PictureEquip"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "OptionEquip1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "OptionEquip2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "OptionEquip3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TextEquip1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TextEquip2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TextEquip3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TextEquip4"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "FrameEquip(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "OptionEquip4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "FrameMode"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TimerHwnd"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "˵��"
      TabPicture(1)   =   "FormMain.frx":E8E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1"
      Tab(1).ControlCount=   1
      Begin VB.Timer TimerHwnd 
         Interval        =   3000
         Left            =   5760
         Top             =   960
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "FormMain.frx":E902
         Top             =   480
         Width           =   10815
      End
      Begin VB.Frame FrameMode 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   16
         Top             =   5760
         Width           =   5415
         Begin VB.CommandButton BtnSave 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   18
            Top             =   280
            Width           =   1455
         End
         Begin VB.CheckBox CbSlow 
            Caption         =   "����ģʽ"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   17
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.OptionButton OptionEquip4 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   15
         Top             =   5400
         Width           =   255
      End
      Begin VB.Frame FrameEquip 
         Caption         =   "������װ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   3
         Left            =   3120
         TabIndex        =   14
         Top             =   3480
         Width           =   2535
         Begin VB.Image ImageEquip4 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   0
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip4 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   1
            Left            =   960
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip4 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   2
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip4 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   3
            Left            =   240
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip4 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   4
            Left            =   960
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip4 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   5
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
      End
      Begin VB.TextBox TextEquip4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   13
         Top             =   4320
         Width           =   6615
      End
      Begin VB.TextBox TextEquip3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   12
         Top             =   3720
         Width           =   6615
      End
      Begin VB.TextBox TextEquip2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   11
         Top             =   3120
         Width           =   6615
      End
      Begin VB.TextBox TextEquip1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   10
         Top             =   2520
         Width           =   6615
      End
      Begin VB.OptionButton OptionEquip3 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   5400
         Width           =   255
      End
      Begin VB.OptionButton OptionEquip2 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   8
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton OptionEquip1 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   7
         Top             =   3000
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.PictureBox PictureEquip 
         AutoRedraw      =   -1  'True
         Height          =   1035
         Left            =   12000
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   6
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Frame FrameEquip 
         Caption         =   "������װ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   3480
         Width           =   2535
         Begin VB.Image ImageEquip3 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   5
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip3 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   4
            Left            =   960
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip3 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   3
            Left            =   240
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip3 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   2
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip3 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   1
            Left            =   960
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip3 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   0
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame FrameEquip 
         Caption         =   "�ڶ���װ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   1
         Left            =   3120
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
         Begin VB.Image ImageEquip2 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   0
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip2 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   1
            Left            =   960
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip2 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   2
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip2 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   3
            Left            =   240
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip2 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   4
            Left            =   960
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip2 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   5
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
      End
      Begin VB.Frame FrameEquip 
         Caption         =   "��һ��װ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   1080
         Width           =   2535
         Begin VB.Image ImageEquip1 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   5
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip1 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   4
            Left            =   960
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip1 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   3
            Left            =   240
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   615
         End
         Begin VB.Image ImageEquip1 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   2
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip1 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   1
            Left            =   960
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
         Begin VB.Image ImageEquip1 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   0
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Timer TimerHotkey 
         Interval        =   50
         Left            =   10440
         Top             =   600
      End
      Begin VB.Frame FrameBag 
         Caption         =   "����װ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   6240
         TabIndex        =   3
         Top             =   1080
         Width           =   4695
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   41
            Left            =   3840
            Stretch         =   -1  'True
            Top             =   4800
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   40
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   4800
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   39
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   4800
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   38
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   4800
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   37
            Left            =   960
            Stretch         =   -1  'True
            Top             =   4800
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   36
            Left            =   240
            Stretch         =   -1  'True
            Top             =   4800
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   35
            Left            =   3840
            Stretch         =   -1  'True
            Top             =   4080
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   34
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   4080
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   33
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   4080
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   32
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   4080
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   31
            Left            =   960
            Stretch         =   -1  'True
            Top             =   4080
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   30
            Left            =   240
            Stretch         =   -1  'True
            Top             =   4080
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   29
            Left            =   3840
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   28
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   27
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   26
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   25
            Left            =   960
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   24
            Left            =   240
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   23
            Left            =   3840
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   22
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   21
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   20
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   19
            Left            =   960
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   18
            Left            =   240
            Stretch         =   -1  'True
            Top             =   2640
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   17
            Left            =   3840
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   16
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   15
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   14
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   13
            Left            =   960
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   12
            Left            =   240
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   11
            Left            =   3840
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   10
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   9
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   8
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   7
            Left            =   960
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   6
            Left            =   240
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   5
            Left            =   3840
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   4
            Left            =   3120
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   3
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   2
            Left            =   1680
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   1
            Left            =   960
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image ImageBag 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Index           =   0
            Left            =   240
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Label LabelHwnd 
         Caption         =   "��Ϸδ����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5480
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LabelTip 
         Caption         =   "��Ȫˮ����Ctrl+1��2��3��4�л���װ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label LabelEquip 
         Caption         =   "��ս���򿪱���������Home����ȡװ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ������
Option Explicit

' ��Ϸ�Ĵ��ھ��
Dim hwnd300 As Long

' ���壬����
Private Sub Form_Load()
    ' ��ʼ����Į���
    Dim dm_ret As Integer
    dm_ret = InitDm()
    If dm_ret = 0 Then
        MsgBox "��ʼ��ʧ�ܣ���ʹ�ù���ԱȨ�����У����ر�ɱ�������" & vbCrLf & "������ǲ��У���ȷ��û�ж�ʧ�ļ��������������ء�"
        End
    End If
End Sub

' ���壬��ʾ
Private Sub Form_Activate()
    ' �������Ӱ汾��
    FormMain.Caption = FormMain.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
    ' ��ȡ�����ļ�
    Call LoadConfig
    ' ����װ��
    Call ScreenEquipDecode
    ' ��ȡ��Ϸ�Ĵ��ھ��
    Delay 50
    Call RefreshHwnd
End Sub

' װ���������װ��
Private Sub ImageBag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim equip As String
    equip = Trim(Index \ 6 + 1) + "-" + Trim(Index Mod 6 + 1)
    ' �ж�ѡ�������һ����װ
    If OptionEquip1.Value = True Then
        If CountChar(TextEquip1.Text, "/") < 5 And CountChar(TextEquip1.Text, equip) < 1 Then
            If TextEquip1.Text = "" Then
                TextEquip1.Text = equip
            Else
                TextEquip1.Text = Trim(TextEquip1.Text) + "/" + equip
            End If
        End If
    ElseIf OptionEquip2.Value = True Then
        If CountChar(TextEquip2.Text, "/") < 5 And CountChar(TextEquip2.Text, equip) < 1 Then
            If TextEquip2.Text = "" Then
                TextEquip2.Text = equip
            Else
                TextEquip2.Text = Trim(TextEquip2.Text) + "/" + equip
            End If
        End If
    ElseIf OptionEquip3.Value = True Then
        If CountChar(TextEquip3.Text, "/") < 5 And CountChar(TextEquip3.Text, equip) < 1 Then
            If TextEquip3.Text = "" Then
                TextEquip3.Text = equip
            Else
                TextEquip3.Text = Trim(TextEquip3.Text) + "/" + equip
            End If
        End If
    Else
        If CountChar(TextEquip4.Text, "/") < 5 And CountChar(TextEquip4.Text, equip) < 1 Then
            If TextEquip4.Text = "" Then
                TextEquip4.Text = equip
            Else
                TextEquip4.Text = Trim(TextEquip4.Text) + "/" + equip
            End If
        End If
    End If
    ' װ��֮��ˢ��
    Call ScreenEquipDecode
End Sub

' �����һ��װ��
Private Sub ImageEquip1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim arr() As String
    arr() = Split(TextEquip1.Text, "/")
    arr() = RemoveArrayItem(arr, Index)
    Dim equip As String
    Dim i As Integer
    
    If IsNotEmpty(arr) Then
        For i = LBound(arr) To UBound(arr)
            If i < UBound(arr) Then
                equip = equip & arr(i) & "/"
            Else
                equip = equip & arr(i)
            End If
        Next i
    End If
    
    TextEquip1.Text = equip
    ' װ��֮��ˢ��
    Call ScreenEquipDecode
End Sub

' ����ڶ���װ��
Private Sub ImageEquip2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim arr() As String
    arr() = Split(TextEquip2.Text, "/")
    arr() = RemoveArrayItem(arr, Index)
    Dim equip As String
    Dim i As Integer
    
    If IsNotEmpty(arr) Then
        For i = LBound(arr) To UBound(arr)
            If i < UBound(arr) Then
                equip = equip & arr(i) & "/"
            Else
                equip = equip & arr(i)
            End If
        Next i
    End If
    
    TextEquip2.Text = equip
    ' װ��֮��ˢ��
    Call ScreenEquipDecode
End Sub

' ���������װ��
Private Sub ImageEquip3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim arr() As String
    arr() = Split(TextEquip3.Text, "/")
    arr() = RemoveArrayItem(arr, Index)
    Dim equip As String
    Dim i As Integer
    
    If IsNotEmpty(arr) Then
        For i = LBound(arr) To UBound(arr)
            If i < UBound(arr) Then
                equip = equip & arr(i) & "/"
            Else
                equip = equip & arr(i)
            End If
        Next i
    End If
    
    TextEquip3.Text = equip
    ' װ��֮��ˢ��
    Call ScreenEquipDecode
End Sub

' ���������װ��
Private Sub ImageEquip4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim arr() As String
    arr() = Split(TextEquip4.Text, "/")
    arr() = RemoveArrayItem(arr, Index)
    Dim equip As String
    Dim i As Integer
    
    If IsNotEmpty(arr) Then
        For i = LBound(arr) To UBound(arr)
            If i < UBound(arr) Then
                equip = equip & arr(i) & "/"
            Else
                equip = equip & arr(i)
            End If
        Next i
    End If
    
    TextEquip4.Text = equip
    ' װ��֮��ˢ��
    Call ScreenEquipDecode
End Sub

' ��ʱ�����жϿ�ݼ�
Private Sub TimerHotkey_Timer()
    If hwnd300 <> 0 Then
        Call CheckHotKey(hwnd300)
    End If
End Sub

' ��ʱ������ȡ���ھ��
Private Sub TimerHwnd_Timer()
    Call RefreshHwnd
End Sub

' ��ť�����������ļ�
Private Sub BtnSave_Click()
    Call SaveConfig
End Sub

' ��ȡ���ھ�����ж���Ϸ�Ƿ�������
Private Sub RefreshHwnd()
    Dim hwnd As Long
    hwnd = GetHwnd
    
    If hwnd <> 0 Then
        ' ��Ϸ��ǰ������
        If hwnd300 <> 0 Then
            ' ��Ϸ֮ǰҲ�����У����ô���
        Else
            ' ��Ϸ֮ǰ�������У�˵���ո�������Ϸ
            hwnd300 = hwnd
            LabelHwnd.Caption = "��Ϸ������"
            LabelHwnd.ForeColor = vbBlue
        End If
    Else
        ' ��Ϸ��ǰû������
        If hwnd300 <> 0 Then
            ' ��Ϸ֮ǰ�����У�˵���ոչر���Ϸ
            hwnd300 = 0
            LabelHwnd.Caption = "��Ϸδ����"
            LabelHwnd.ForeColor = vbRed
        Else
            ' ��Ϸ֮ǰҲ�������У����ô���
        End If
    End If
End Sub

' ��ȡ�����ļ�
Private Sub LoadConfig()
    ' ����ģʽ
    CbSlow.Value = (GetSlowMode)
    ' һ����װ
    TextEquip1.Text = (GetEquip(1))
    TextEquip2.Text = (GetEquip(2))
    TextEquip3.Text = (GetEquip(3))
    TextEquip4.Text = (GetEquip(4))
End Sub

' ���������ļ�
Private Sub SaveConfig()
    ' ����ģʽ
    Call SetSlowMode(CbSlow.Value)
    ' һ����װ
    Call SetEquip(1, TextEquip1.Text)
    Call SetEquip(2, TextEquip2.Text)
    Call SetEquip(3, TextEquip3.Text)
    Call SetEquip(4, TextEquip4.Text)
End Sub

