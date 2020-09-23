VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Enhanced Text Control Demonstrations ::."
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   15055
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "General Properties"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Input Enhancing"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Normal View"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(2)=   "Frame10"
      Tab(2).Control(3)=   "Frame11"
      Tab(2).Control(4)=   "Frame12"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Focus View"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame13"
      Tab(3).Control(1)=   "Frame14"
      Tab(3).Control(2)=   "Frame15"
      Tab(3).Control(3)=   "Frame16"
      Tab(3).Control(4)=   "Frame17"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Misc"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame18"
      Tab(4).Control(1)=   "Frame19"
      Tab(4).Control(2)=   "Frame20"
      Tab(4).Control(3)=   "Frame23"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Example"
      TabPicture(5)   =   "Form1.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label79"
      Tab(5).Control(1)=   "Label80"
      Tab(5).Control(2)=   "lblStatus"
      Tab(5).Control(3)=   "Frame21"
      Tab(5).Control(4)=   "Frame22"
      Tab(5).Control(5)=   "Frame24"
      Tab(5).ControlCount=   6
      Begin VB.Frame Frame24 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   162
         Top             =   4920
         Width           =   7335
         Begin VB.CommandButton cmdAbout 
            Caption         =   "About"
            Height          =   375
            Left            =   6240
            TabIndex        =   170
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label Label85 
            Alignment       =   2  'Center
            Caption         =   "Credits"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   2760
            TabIndex        =   171
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label84 
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.geocities.com/priyank_modi/"
            DragIcon        =   "Form1.frx":00A8
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   840
            MouseIcon       =   "Form1.frx":01FA
            MousePointer    =   99  'Custom
            TabIndex        =   169
            Top             =   3120
            Width           =   6375
         End
         Begin VB.Label Label83 
            BackStyle       =   0  'Transparent
            Caption         =   "Visite :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label77 
            Caption         =   "I hope you enjoy it...Pls pls Have your valuable vote!!!"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   167
            Top             =   2760
            Width           =   7095
         End
         Begin VB.Label Label75 
            Caption         =   $"Form1.frx":034C
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   855
            Left            =   120
            TabIndex        =   166
            Top             =   1800
            Width           =   7095
         End
         Begin VB.Label Label74 
            Caption         =   $"Form1.frx":0482
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   495
            Left            =   120
            TabIndex        =   165
            Top             =   1320
            Width           =   7095
         End
         Begin VB.Label Label72 
            Caption         =   $"Form1.frx":0528
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   615
            Left            =   120
            TabIndex        =   164
            Top             =   600
            Width           =   7095
         End
         Begin VB.Label Label71 
            Caption         =   "Hi friends,"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   240
            Width           =   7095
         End
      End
      Begin VB.Frame Frame23 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   157
         Top             =   7200
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText24 
            Height          =   330
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   582
            NormalBackColor =   -2147483643
            NormalFontColor =   -2147483640
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Press ENTER Key After Focusing This TextBox"
            EnterExitKey    =   -1  'True
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText EnhancedText26 
            Height          =   330
            Left            =   3720
            TabIndex        =   38
            Top             =   600
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   582
            NormalBackColor =   -2147483643
            NormalFontColor =   -2147483640
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Property is false here"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label82 
            BackStyle       =   0  'Transparent
            Caption         =   "When this property is true Allows Tabing with Enter key.Try out"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   960
            Width           =   7095
         End
         Begin VB.Line Line21 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   1560
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label81 
            BackStyle       =   0  'Transparent
            Caption         =   "# EnterExit Key :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame22 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   150
         Top             =   3360
         Width           =   7335
         Begin VB.CommandButton cmdLast 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6840
            MouseIcon       =   "Form1.frx":0613
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":0765
            Style           =   1  'Graphical
            TabIndex        =   154
            ToolTipText     =   "Move Last"
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdNext 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6480
            MouseIcon       =   "Form1.frx":09B7
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":0B09
            Style           =   1  'Graphical
            TabIndex        =   153
            ToolTipText     =   "Move Next"
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdPrevious 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   480
            MouseIcon       =   "Form1.frx":0D15
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":0E67
            Style           =   1  'Graphical
            TabIndex        =   152
            ToolTipText     =   "Move Previous"
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdFirst 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            MouseIcon       =   "Form1.frx":1076
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":11C8
            Style           =   1  'Graphical
            TabIndex        =   151
            ToolTipText     =   "Move First"
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Cl&ose"
            Height          =   375
            Left            =   6120
            MouseIcon       =   "Form1.frx":1417
            MousePointer    =   99  'Custom
            TabIndex        =   49
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   4920
            MouseIcon       =   "Form1.frx":1569
            MousePointer    =   99  'Custom
            TabIndex        =   48
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   375
            Left            =   3720
            MouseIcon       =   "Form1.frx":16BB
            MousePointer    =   99  'Custom
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   2520
            MouseIcon       =   "Form1.frx":180D
            MousePointer    =   99  'Custom
            TabIndex        =   46
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   1320
            MouseIcon       =   "Form1.frx":195F
            MousePointer    =   99  'Custom
            TabIndex        =   45
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "Add &New"
            Height          =   375
            Left            =   120
            MouseIcon       =   "Form1.frx":1AB1
            MousePointer    =   99  'Custom
            TabIndex        =   39
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRecordCount 
            Alignment       =   2  'Center
            Height          =   210
            Left            =   840
            TabIndex        =   160
            Top             =   780
            Width           =   5655
         End
      End
      Begin VB.Frame Frame21 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   139
         Top             =   1200
         Width           =   7335
         Begin Project1.EnhancedText enhPinCode 
            Height          =   330
            Left            =   960
            TabIndex        =   43
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            InputType       =   2
            CharCase        =   3
            Alignment       =   1
            NormalBackColor =   -2147483643
            NormalFontColor =   -2147483640
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            EnterExitKey    =   -1  'True
            FocusBorderPattern=   3
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   6
            AutoTab         =   -1  'True
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText enhSalary 
            Height          =   330
            Left            =   960
            TabIndex        =   44
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            InputType       =   2
            CharCase        =   3
            Alignment       =   1
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            EnterExitKey    =   -1  'True
            FocusBorderPattern=   3
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   "0.00"
            MaxLength       =   8
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText enhPhone 
            Height          =   330
            Left            =   960
            TabIndex        =   42
            Top             =   960
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   582
            InputType       =   4
            CharCase        =   3
            NormalBackColor =   -2147483643
            NormalFontColor =   -2147483640
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            EnterExitKey    =   -1  'True
            FocusBorderPattern=   3
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   50
            SpecialCharacter=   "0123456789, ()"
         End
         Begin Project1.EnhancedText enhAddress 
            Height          =   330
            Left            =   960
            TabIndex        =   41
            Top             =   600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   582
            InputType       =   3
            CharCase        =   3
            NormalBackColor =   -2147483643
            NormalFontColor =   -2147483640
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            EnterExitKey    =   -1  'True
            FocusBorderPattern=   3
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   200
            SpecialCharacter=   "(){}[],:"
         End
         Begin Project1.EnhancedText enhName 
            Height          =   330
            Left            =   960
            TabIndex        =   40
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   582
            InputType       =   1
            CharCase        =   3
            NormalBackColor =   -2147483643
            NormalFontColor =   -2147483640
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            EnterExitKey    =   -1  'True
            FocusBorderPattern=   3
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   50
            SpecialCharacter=   ""
         End
         Begin VB.Label Label78 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PIN Code :"
            Height          =   255
            Left            =   0
            TabIndex        =   149
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblPincode 
            Caption         =   "6"
            Height          =   255
            Left            =   3240
            TabIndex        =   148
            Top             =   1320
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label76 
            Alignment       =   1  'Right Justify
            Caption         =   "Salary :"
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lblSalary 
            Caption         =   "8"
            Height          =   255
            Left            =   3240
            TabIndex        =   146
            Top             =   1680
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblPhone 
            Caption         =   "50"
            Height          =   255
            Left            =   6600
            TabIndex        =   145
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            Caption         =   "Phone :"
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblAddress 
            Caption         =   "200"
            Height          =   255
            Left            =   6600
            TabIndex        =   143
            Top             =   600
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblName 
            Caption         =   "50"
            Height          =   255
            Left            =   6600
            TabIndex        =   142
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            Caption         =   "Address :"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Caption         =   "Name :"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame20 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   134
         Top             =   5280
         Width           =   7335
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   960
            Width           =   3375
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   960
            Width           =   3255
         End
         Begin Project1.EnhancedText TLock 
            Height          =   330
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            Alignment       =   2
            NormalBackColor =   -2147483643
            NormalFontColor =   -2147483640
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Unlock - Enabled"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label68 
            Alignment       =   2  'Center
            Caption         =   "Locked"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   3840
            TabIndex        =   138
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label Label67 
            Alignment       =   2  'Center
            Caption         =   "Enabled"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "Select the Key gaurd constrains and see the effect."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   1560
            Width           =   7095
         End
         Begin VB.Line Line20 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2880
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label65 
            BackStyle       =   0  'Transparent
            Caption         =   "# Locked and Enabled Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame19 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   126
         Top             =   2280
         Width           =   7335
         Begin Project1.EnhancedText txtEnhFormat 
            Height          =   330
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            InputType       =   2
            NormalBackColor =   -2147483643
            NormalFontColor =   -2147483640
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "$ 0.00"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   "$ 0.00"
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label64 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "For more you can refer to Normal Textcontrol format rules from MSDN or directly from control."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   2520
            Width           =   6735
         End
         Begin VB.Label Label63 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "For formating with 4 decimal placing Plus $ symbol at first fill '$ 0.0000' In the TextFormat Property"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   2280
            Width           =   6735
         End
         Begin VB.Label Label62 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "For formating with 4 decimal placing fill 0.0000 In the TextFormat Property"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   2040
            Width           =   5895
         End
         Begin VB.Label Label61 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "For formating with 3 decimal placing fill 0.000 In the TextFormat Property"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   1800
            Width           =   5895
         End
         Begin VB.Label Label60 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "For formating with 2 decimal placing fill 0.00 In theTextFormat Property"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   1560
            Width           =   5895
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "This property when Set to some equations automatically adjust enhanced Text value according to format. for ex."
            ForeColor       =   &H00404040&
            Height          =   495
            Left            =   120
            TabIndex        =   128
            Top             =   1080
            Width           =   7095
         End
         Begin VB.Line Line19 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2280
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "# Built In Format options :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame18 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   121
         Top             =   420
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText21 
            Height          =   330
            Left            =   1800
            TabIndex        =   31
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This Text Get selected."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            AutoTab         =   -1  'True
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText EnhancedText22 
            Height          =   330
            Left            =   1800
            TabIndex        =   32
            Top             =   960
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   582
            OnFocusSelect   =   0   'False
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This Text Remain Unselected."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "# OnFocus Selection Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   3255
         End
         Begin VB.Line Line18 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2760
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form1.frx":1C03
            ForeColor       =   &H00404040&
            Height          =   495
            Left            =   120
            TabIndex        =   124
            Top             =   1320
            Width           =   7095
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto Selction True :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto Selection False :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Frame Frame17 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   113
         Top             =   420
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText20 
            Height          =   330
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            OnFocusSelect   =   0   'False
            FocusFontColor  =   0
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Focus background color."
            NormalBorderColor=   8421504
            FocusBorderColor=   8421504
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label49 
            Caption         =   "Try out By focusing"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   5760
            TabIndex        =   116
            Top             =   240
            Width           =   1455
         End
         Begin VB.Line Line17 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2520
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "# Focus Background Color :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label47 
            Caption         =   "Select Your custom Background color from color selection dialog."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   960
            Width           =   5895
         End
      End
      Begin VB.Frame Frame16 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   110
         Top             =   1920
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText19 
            Height          =   330
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            OnFocusSelect   =   0   'False
            FocusBackColor  =   16777215
            FocusFontColor  =   0
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Focus Border color."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label50 
            Caption         =   "Try out By focusing"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   5760
            TabIndex        =   117
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label46 
            Caption         =   "Select Your custom Bordercolor from color selection dialog."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "# Focus Border Color :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   2895
         End
         Begin VB.Line Line16 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame15 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   107
         Top             =   3480
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText18 
            Height          =   330
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            OnFocusSelect   =   0   'False
            FocusBackColor  =   16777215
            FocusFontColor  =   0
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Focus Border Pattern."
            FocusBorderPattern=   3
            NormalBorderColor=   8421504
            FocusBorderColor=   8421504
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label51 
            Caption         =   "Try out By focusing"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   5760
            TabIndex        =   118
            Top             =   240
            Width           =   1455
         End
         Begin VB.Line Line15 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2280
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "# Focus Border Pattern :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label43 
            Caption         =   "Select Your custom Borderpattern from Nine diffrent Border style list."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   960
            Width           =   5895
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   104
         Top             =   5040
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText17 
            Height          =   330
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            OnFocusSelect   =   0   'False
            FocusBackColor  =   16777215
            FocusFontColor  =   0
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Focus Fonts."
            NormalBorderColor=   8421504
            FocusBorderColor=   8421504
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label52 
            Caption         =   "Try out By focusing"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   5760
            TabIndex        =   119
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label42 
            Caption         =   "Select Your custom font from font selection dialog."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "# Focus Font :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   2895
         End
         Begin VB.Line Line14 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   1320
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   101
         Top             =   6600
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText16 
            Height          =   330
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            OnFocusSelect   =   0   'False
            FocusBackColor  =   16777215
            FocusFontColor  =   192
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Focus Font Color."
            NormalBorderColor=   8421504
            FocusBorderColor=   8421504
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label53 
            Caption         =   "Try out By focusing"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   5760
            TabIndex        =   120
            Top             =   240
            Width           =   1455
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   1920
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "# Focus Font Color :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label39 
            Caption         =   "Select Your custom font color from selection dialog."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   960
            Width           =   5895
         End
      End
      Begin VB.Frame Frame12 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   98
         Top             =   6600
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText15 
            Height          =   330
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            NormalFontColor =   12582912
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Normal Font Color."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label38 
            Caption         =   "Select Your custom font color from selection dialog."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "# Normal Font Color :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   240
            Width           =   2895
         End
         Begin VB.Line Line12 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   1920
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   95
         Top             =   5040
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText14 
            Height          =   330
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Normal Fonts."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   1440
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "# Normal Font :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label35 
            Caption         =   "Select Your custom font from font selection dialog."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   960
            Width           =   5895
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   92
         Top             =   3480
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText13 
            Height          =   330
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Normal Border Pattern."
            NormalBorderPattern=   3
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label34 
            Caption         =   "Select Your custom Borderpattern from Nine diffrent Border style list."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "# Normal Border Pattern :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   240
            Width           =   2895
         End
         Begin VB.Line Line10 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2280
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   89
         Top             =   1920
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText12 
            Height          =   330
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Normal Border color."
            NormalBorderColor=   12583104
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "# Normal Border Color :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label31 
            Caption         =   "Select Your custom Bordercolor from color selection dialog."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   960
            Width           =   5895
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   86
         Top             =   420
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText11 
            Height          =   330
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            NormalBackColor =   16777152
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "This is sample Normal background color."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label30 
            Caption         =   "Select Your custom Background color from color selection dialog."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "# Normal Background Color :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   2895
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2520
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2055
         Left            =   120
         TabIndex        =   83
         Top             =   5940
         Width           =   7335
         Begin VB.CommandButton Command2 
            Caption         =   "Set Length"
            Height          =   330
            Left            =   2040
            TabIndex        =   9
            Top             =   960
            Width           =   1335
         End
         Begin Project1.EnhancedText txtLen 
            Height          =   330
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            InputType       =   2
            Alignment       =   1
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText enh 
            Height          =   330
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label28 
            Caption         =   "# Maximum Length Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   2775
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2880
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form1.frx":1C8D
            ForeColor       =   &H00404040&
            Height          =   495
            Left            =   120
            TabIndex        =   84
            Top             =   1440
            Width           =   6735
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   120
         TabIndex        =   80
         Top             =   4500
         Width           =   7335
         Begin Project1.EnhancedText tpwdchar 
            Height          =   330
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   "&"
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "This is same Property as the passowrd character for Simple Textbox."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   960
            Width           =   7215
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2880
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label25 
            Caption         =   "# Password Character Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   65
         Top             =   3180
         Width           =   7335
         Begin VB.CommandButton Command1 
            Caption         =   "Set"
            Height          =   330
            Left            =   6360
            TabIndex        =   20
            Top             =   4440
            Width           =   615
         End
         Begin Project1.EnhancedText tsp 
            Height          =   330
            Left            =   2400
            TabIndex        =   19
            Top             =   4440
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText t1 
            Height          =   330
            Left            =   1320
            TabIndex        =   14
            Top             =   600
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText t2 
            Height          =   330
            Left            =   1320
            TabIndex        =   15
            Top             =   1200
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   582
            InputType       =   1
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText t3 
            Height          =   330
            Left            =   1320
            TabIndex        =   16
            Top             =   1800
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   582
            InputType       =   2
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText t4 
            Height          =   330
            Left            =   1320
            TabIndex        =   17
            Top             =   2400
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   582
            InputType       =   3
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText t5 
            Height          =   330
            Left            =   1320
            TabIndex        =   18
            Top             =   3000
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   582
            InputType       =   4
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label24 
            Caption         =   "Enter Special charaters whichever you want and Press Set to set the char for above 5 Textboxs."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   4800
            Width           =   7095
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Customize Special Characters:"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   4440
            Width           =   2175
         End
         Begin VB.Label Label22 
            Caption         =   $"Form1.frx":1D17
            ForeColor       =   &H00404040&
            Height          =   615
            Left            =   120
            TabIndex        =   77
            Top             =   3720
            Width           =   7095
         End
         Begin VB.Label Label21 
            Caption         =   "Customize:Special Character field Char's."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1320
            TabIndex        =   76
            Top             =   3360
            Width           =   5895
         End
         Begin VB.Label Label11 
            Caption         =   "Numeric:A-Z, a-z, 0-9, DOT, BKSPACE  And Special Character field Char's."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1320
            TabIndex        =   75
            Top             =   2760
            Width           =   5895
         End
         Begin VB.Label Label20 
            Caption         =   "Numeric:0-9, DOT, BKSPACE  And Special Character field Char's."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1320
            TabIndex        =   74
            Top             =   2160
            Width           =   5895
         End
         Begin VB.Label Label19 
            Caption         =   "Alphabetic:A-Z,a-z,SPACE,BKSPACE And SpecialCharacter Property Char's."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1320
            TabIndex        =   73
            Top             =   1560
            Width           =   5895
         End
         Begin VB.Label Label18 
            Caption         =   "None:All the Characters."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1320
            TabIndex        =   72
            Top             =   960
            Width           =   5895
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Customize :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "AlphaNumeric :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Numeric :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Alphabetic :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "None :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "# Character Filter Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   2895
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2520
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   62
         Top             =   420
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText7 
            Height          =   330
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "None - Format as your choice"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText EnhancedText8 
            Height          =   330
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            CharCase        =   1
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "UPPERCASE - GENERALLY WRITTEN CHARACTERS ARE IN UPPERCASE."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText EnhancedText9 
            Height          =   330
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            CharCase        =   2
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "lower case - generally writtern characters are in lower case."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText EnhancedText10 
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            CharCase        =   2
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Proper Case - Written Charactyers Are In Proper Combinations Of Small And Capital."
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label10 
            Caption         =   $"Form1.frx":1E16
            ForeColor       =   &H00404040&
            Height          =   615
            Left            =   120
            TabIndex        =   64
            Top             =   2040
            Width           =   7095
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2520
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "# Character Case Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Left            =   120
         TabIndex        =   53
         Top             =   420
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText1 
            Height          =   330
            Left            =   120
            TabIndex        =   1
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Left Alignment"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText EnhancedText2 
            Height          =   330
            Left            =   120
            TabIndex        =   2
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            Alignment       =   2
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Center Alignment"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText EnhancedText3 
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            Alignment       =   1
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "Right Alignment"
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "# Alignment Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   2295
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label2 
            Caption         =   "These are alignment Properties same as normal text control."
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1680
            Width           =   7095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   120
         TabIndex        =   51
         Top             =   2580
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText4 
            Height          =   330
            Left            =   1800
            TabIndex        =   4
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            EnterExitKey    =   -1  'True
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   10
            AutoTab         =   -1  'True
            SpecialCharacter=   ""
         End
         Begin Project1.EnhancedText EnhancedText5 
            Height          =   330
            Left            =   1800
            TabIndex        =   5
            Top             =   960
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   10
            SpecialCharacter=   ""
         End
         Begin VB.Label Label6 
            Caption         =   "Automatic TAB False :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Automatic TAB True :"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form1.frx":1EF2
            ForeColor       =   &H00404040&
            Height          =   495
            Left            =   120
            TabIndex        =   56
            Top             =   1320
            Width           =   7095
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   1920
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label4 
            Caption         =   "# AutoTab Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   120
         TabIndex        =   59
         Top             =   2580
         Width           =   7335
         Begin Project1.EnhancedText EnhancedText6 
            Height          =   330
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            BeginProperty NormalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FocusFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NormalBorderColor=   8421504
            FocusBorderColor=   33023
            PasswordChar    =   ""
            Object.Tag             =   ""
            TextFormat      =   ""
            MaxLength       =   0
            SpecialCharacter=   ""
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "This Property is the Same property as the Simple TextBox validations Property.Try out with your validations code."
            ForeColor       =   &H00404040&
            Height          =   495
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Width           =   7215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00404040&
            X1              =   120
            X2              =   2640
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label7 
            Caption         =   "# Cause Validation Property :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Label lblStatus 
         ForeColor       =   &H000B33AA&
         Height          =   255
         Left            =   -74880
         TabIndex        =   161
         Top             =   4680
         Width           =   7335
      End
      Begin VB.Label Label80 
         Caption         =   $"Form1.frx":1FA4
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   -74880
         TabIndex        =   156
         Top             =   720
         Width           =   7335
      End
      Begin VB.Label Label79 
         Caption         =   "User Registration Form ::."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   -74880
         TabIndex        =   155
         Top             =   480
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim SaveFlag As Boolean
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmdAbout_Click()
    Load frmSplash
    frmSplash.Show modal
End Sub

Private Sub Combo1_Click()
    TLock.Enabled = Combo1.ItemData(Combo1.ListIndex)
    Call ComboStatus
End Sub

Private Sub Combo2_Click()
    TLock.Locked = Combo2.ItemData(Combo2.ListIndex)
    Call ComboStatus
End Sub

Public Sub ComboStatus()
    If (TLock.Enabled = True And TLock.Locked = True) Then
        TLock.Text = "Enable - Locked"
    ElseIf (TLock.Enabled = True And TLock.Locked = False) Then
        TLock.Text = "Enable - UnLocked"
    ElseIf (TLock.Enabled = False And TLock.Locked = True) Then
        TLock.Text = "Disabled - Locked"
    Else
        TLock.Text = "Disabled - UnLocked"
    End If
End Sub

Private Sub Command1_Click()
    t1.SpecialCharacter = tsp.Text
    t2.SpecialCharacter = tsp.Text
    t3.SpecialCharacter = tsp.Text
    t4.SpecialCharacter = tsp.Text
    t5.SpecialCharacter = tsp.Text
End Sub

Private Sub Command2_Click()
    enh.MaxLength = CDbl(txtLen.Text)
End Sub

Private Sub Command3_Click()
    txtEnhFormat.TextFormat = txtequ.Text
End Sub

Private Sub enhAddress_GotFocus()
    lblAddress.Visible = True
    lblAddress.Caption = enhAddress.MaxLength - Len(enhAddress.Text)
    lblStatus.Caption = "Enter User Address.Best example AlphaNumeric text."
End Sub
Private Sub enhAddress_KeyPress(KeyAscii As Integer)
    lblAddress.Caption = enhAddress.MaxLength - Len(enhAddress.Text)
End Sub
Private Sub enhAddress_LostFocus()
    lblAddress.Visible = False
End Sub

Private Sub enhName_GotFocus()
    lblName.Visible = True
    lblName.Caption = enhName.MaxLength - Len(enhName.Text)
    lblStatus.Caption = "Enter User FullName.Best example for Alphabetic text."
End Sub
Private Sub enhName_KeyPress(KeyAscii As Integer)
    lblName.Caption = enhName.MaxLength - Len(enhName.Text)
End Sub
Private Sub enhName_LostFocus()
    lblName.Visible = False
End Sub

Private Sub enhPhone_GotFocus()
    lblPhone.Visible = True
    lblPhone.Caption = enhPhone.MaxLength - Len(enhPhone.Text)
    lblStatus.Caption = "Enter User Phone No.Best example for custom text."
End Sub
Private Sub enhPhone_KeyPress(KeyAscii As Integer)
    lblPhone.Caption = enhPhone.MaxLength - Len(enhPhone.Text)
End Sub
Private Sub enhPhone_LostFocus()
    lblPhone.Visible = False
End Sub

Private Sub enhPinCode_GotFocus()
    lblPincode.Visible = True
    lblPincode.Caption = enhPinCode.MaxLength - Len(enhPinCode.Text)
    lblStatus.Caption = "Enter User PinCode.Best example for Numeric 6 digite autoTabing text."
End Sub
Private Sub enhPinCode_KeyPress(KeyAscii As Integer)
    lblPincode.Caption = enhPinCode.MaxLength - Len(enhPinCode.Text)
End Sub
Private Sub enhPinCode_LostFocus()
    lblPincode.Visible = False
End Sub

Private Sub enhSalary_GotFocus()
    lblSalary.Visible = True
    lblSalary.Caption = enhSalary.MaxLength - Len(enhSalary.Text)
    lblStatus.Caption = "Enter User Salary.Best example for Numeric with formating text."
End Sub
Private Sub enhSalary_KeyPress(KeyAscii As Integer)
    lblSalary.Caption = enhSalary.MaxLength - Len(enhSalary.Text)
End Sub
Private Sub enhSalary_LostFocus()
    lblSalary.Visible = False
End Sub

Private Sub Form_Load()
On Error GoTo errLable
   
    Dim str As String
    
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\db.mdb;Persist Security Info=False;"
    
    Set rs = New ADODB.Recordset
    str = "Select * from masUser;"
    rs.Open str, db, adOpenStatic, adLockOptimistic

    showdata
    SetTextLock (True)
    SetButton (True)
    
    Combo1.AddItem "True"
    Combo1.ItemData(Combo1.NewIndex) = True
    Combo1.AddItem "False"
    Combo1.ItemData(Combo1.NewIndex) = False
    Combo1.ListIndex = 0
    
    Combo2.AddItem "True"
    Combo2.ItemData(Combo2.NewIndex) = True
    Combo2.AddItem "False"
    Combo2.ItemData(Combo2.NewIndex) = False
    Combo2.ListIndex = 1

'    Call cmdAbout_Click
Exit Sub
errLable:
MsgBox (Err.Number & "  " & Err.Description)
End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Private Function cheak() As Boolean
   Dim status As Boolean
   status = False
    If enhName.Text = "" Then
        MsgBox ("Please enter the name."), vbInformation, "Information required"
    ElseIf enhAddress.Text = "" Then
        MsgBox ("Please enter the Address."), vbInformation, "Information required"
    ElseIf enhPhone.Text = "" Then
        MsgBox ("Please enter the Phone."), vbInformation, "Information required"
    ElseIf enhPinCode.Text = "" Then
        MsgBox ("Please enter the PinCode."), vbInformation, "Information required"
'   ElseIf enhSalary.Text = "" Then
'       MsgBox ("Please enter the Salary."), vbInformation, "Information required"
    Else
        status = True
    End If
cheak = status
End Function
Private Sub SetTextLock(val As Boolean)
    enhName.Locked = val
    enhAddress.Locked = val
    enhPhone.Locked = val
    enhPinCode.Locked = val
    enhSalary.Locked = val
End Sub
Private Sub ClearText()
    enhName.Text = ""
    enhAddress.Text = ""
    enhPhone.Text = ""
    enhPinCode.Text = ""
    enhSalary.Text = ""
End Sub
Private Sub showdata()
    If rs.EOF = False And rs.BOF = False Then
        enhName.Text = rs("fldName")
        enhAddress.Text = rs("fldAddress")
        enhPhone.Text = rs("fldPhone")
        enhPinCode.Text = rs("fldPinCode")
        enhSalary.Text = rs("fldSalary")
        lblRecordCount.Caption = rs.AbsolutePosition & " OF " & rs.RecordCount
    End If
End Sub
Private Sub SetButton(val As Boolean)
    cmdFirst.Enabled = val
    cmdPrevious.Enabled = val
    cmdNext.Enabled = val
    cmdLast.Enabled = val
    cmdDelete.Enabled = val
    cmdEdit.Enabled = val
    cmdAddNew.Enabled = val
    cmdSave.Enabled = Not val
    cmdCancel.Enabled = Not val
  '  cmdClose.Enabled = Not val
    If Not (rs.EOF = False And rs.BOF = False) Then
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
    End If
End Sub
Private Sub cmdCancel_Click()
    SetTextLock (True)
    ClearText
    SetButton (True)
    cmdFirst_Click
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
On Error GoTo lable
Dim str As String
Beep
If MsgBox("Execution of command will delete current Datarecord, Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
    str = "DELETE FROM masUser WHERE "
    str = str & "ID = "
    str = str & CDbl(rs("ID"))
    db.Execute str
    rs.Requery
    MsgBox "Record deleted sucessfully.", vbinformayion, "Delete"

    If rs.BOF And rs.EOF Then
        Call ClearText
        MsgBox ("The previous record was last record, Now no record left In Database."), vbInformation, "Last record"
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
    Else
        rs.MoveNext
            If rs.EOF Then
                rs.MoveLast
            End If
    showdata
    End If
End If
Exit Sub
lable:
If Err.Number = -2147217873 Then
MsgBox "This record refrences Daily Transaction.This record cannot be deleted.", vbInformation, "Cannot Delete"
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub
Private Sub cmdEdit_Click()
On Error GoTo errLable
    SetTextLock (False)
    SetButton (False)
    SaveFlag = False
    enhName.SetFocus
    Exit Sub
errLable:
MsgBox Err.Number & "  " & Err.Description
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo errLable
    SetTextLock (False)
    ClearText
    SetButton (False)
    SaveFlag = True
    enhName.SetFocus
    Exit Sub
errLable:
MsgBox Err.Number & "  " & Err.Description
End Sub
Private Sub cmdSave_Click()
On Error GoTo errLable
    If (cheak = True) Then
        With rs
            If SaveFlag = True Then
                .AddNew
            End If
                .Fields("fldName") = Trim(enhName.Text)
                .Fields("fldAddress") = Trim(enhAddress.Text)
                .Fields("fldPhone") = Trim(enhPhone.Text)
                .Fields("fldPinCode") = Trim(enhPinCode.Text)
                .Fields("fldSalary") = CDbl(enhSalary.Text)
                .Update
                .Requery
                .MoveLast
        End With
        SetTextLock (True)
        SetButton (True)
        showdata
        cmdAddNew.SetFocus
    End If
Exit Sub
errLable:
MsgBox Err.Number & "  " & Err.Description
End Sub
Private Sub cmdFirst_Click()
On Error GoTo GoFirstError
    If rs.BOF = False And rs.EOF = False Then
        rs.MoveFirst
        Call showdata
    End If
Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub
Private Sub cmdLast_Click()
On Error GoTo GoLastError
    If rs.BOF = False And rs.EOF = False Then
        rs.MoveLast
        Call showdata
    End If
Exit Sub
GoLastError:
MsgBox Err.Description
End Sub
Private Sub cmdNext_Click()
On Error GoTo GoNextError
    If Not rs.EOF Then
        rs.MoveNext
    End If
    
    If rs.EOF And rs.RecordCount > 0 Then
        rs.MoveLast
    End If
    
    Call showdata
Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo GoPrevError
    If Not rs.BOF Then
        rs.MovePrevious
    End If
    
    If rs.BOF And rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    Call showdata
Exit Sub
GoPrevError:
MsgBox Err.Number & Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If MsgBox("Are You Sure you want to Quit ?", vbExclamation + vbOKCancel, "Enhanced Text Control") = vbOK Then
    If MsgBox("Visite priyank modis personal website ?", vbYesNo, "Enhanced Text Control") = vbYes Then
    Call Label84_Click
    End If
    Unload frmSample
    Unload frmSplash
Else
    Cancel = True
End If
End Sub

Private Sub Label84_Click()
ShellExecute Me.hWnd, vbNullString, "http://www.geocities.com/priyank_modi/", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub
