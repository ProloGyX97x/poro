VERSION 5.00
Begin VB.Form FrmIntervalos 
   BackColor       =   &H00000010&
   BorderStyle     =   0  'None
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7080
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00151515&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5865
      ScaleWidth      =   6855
      TabIndex        =   0
      Top             =   600
      Width           =   6885
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   5175
         Left            =   120
         ScaleHeight     =   5175
         ScaleWidth      =   3375
         TabIndex        =   8
         Top             =   120
         Width           =   3375
         Begin VB.CheckBox cAimBot 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Modo Aimbot"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CheckBox cModoMacros 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Modo Macros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   960
            TabIndex        =   14
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtData 
            BackColor       =   &H00151515&
            BorderStyle     =   0  'None
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
            Height          =   615
            Left            =   30
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   13
            Text            =   "FrmIntervalos.frx":0000
            Top             =   720
            Width           =   3210
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00151515&
            BorderStyle     =   0  'None
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
            Height          =   855
            Left            =   30
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   12
            Text            =   "FrmIntervalos.frx":0061
            Top             =   1800
            Width           =   3210
         End
         Begin VB.CheckBox cAutoRemo 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Auto remover inteligente"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   3120
            Width           =   2775
         End
         Begin VB.TextBox iAutoRemo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   2400
            TabIndex        =   10
            Text            =   "500"
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox iAutoLanzar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "500"
            Top             =   4500
            Width           =   495
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "[General]"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   120
            Width           =   3345
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            Index           =   0
            X1              =   0
            X2              =   3360
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Auto remover despues de"
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
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   3480
            Width           =   2235
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MS"
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
            Height          =   225
            Index           =   2
            Left            =   2880
            TabIndex        =   19
            Top             =   3480
            Width           =   315
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            Index           =   1
            X1              =   0
            X2              =   3360
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   1200
            TabIndex        =   18
            Top             =   4500
            Width           =   285
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   1950
            TabIndex        =   17
            Top             =   4500
            Width           =   270
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Intervalo de auto lanzar(MACRO)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   4
            Left            =   15
            TabIndex        =   16
            Top             =   4200
            Width           =   3315
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            Index           =   2
            X1              =   0
            X2              =   3360
            Y1              =   5040
            Y2              =   5040
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   5055
         Left            =   3600
         ScaleHeight     =   5055
         ScaleWidth      =   3195
         TabIndex        =   5
         Top             =   120
         Width           =   3200
         Begin VB.TextBox ipRojas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            TabIndex        =   42
            Text            =   "70"
            Top             =   3840
            Width           =   375
         End
         Begin VB.TextBox ipAzules 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   1920
            TabIndex        =   39
            Text            =   "50"
            Top             =   4560
            Width           =   375
         End
         Begin VB.TextBox iAzules 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   1680
            TabIndex        =   34
            Text            =   "250"
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox idAzules 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   1650
            TabIndex        =   30
            Text            =   "250"
            Top             =   3000
            Width           =   495
         End
         Begin VB.TextBox idRojas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   1650
            TabIndex        =   26
            Text            =   "250"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox iRojas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            IMEMode         =   3  'DISABLE
            Left            =   1680
            TabIndex        =   6
            Text            =   "250"
            Top             =   840
            Width           =   495
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            Index           =   4
            X1              =   0
            X2              =   3360
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "%"
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
            Height          =   225
            Index           =   20
            Left            =   2310
            TabIndex        =   44
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Porcentaje de Vida:"
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
            Height          =   225
            Index           =   19
            Left            =   180
            TabIndex        =   43
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "%"
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
            Height          =   225
            Index           =   22
            Left            =   2310
            TabIndex        =   41
            Top             =   4560
            Width           =   255
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Porcentaje de Mana:"
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
            Height          =   225
            Index           =   21
            Left            =   135
            TabIndex        =   40
            Top             =   4560
            Width           =   1785
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Priorizar azul si mi mana es menor a :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   0
            TabIndex        =   38
            Top             =   4200
            Width           =   3105
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MS"
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
            Height          =   225
            Index           =   17
            Left            =   2160
            TabIndex        =   37
            Top             =   1440
            Width           =   315
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "INTERVALO:"
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
            Height          =   225
            Index           =   16
            Left            =   600
            TabIndex        =   36
            Top             =   1440
            Width           =   1035
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Intervalo de Poción AZÚL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   15
            Left            =   375
            TabIndex        =   35
            Top             =   1080
            Width           =   2475
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Priorizar rojas si mi vida es menor a:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   14
            Left            =   0
            TabIndex        =   33
            Top             =   3480
            Width           =   3105
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            Index           =   5
            X1              =   0
            X2              =   3360
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MS"
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
            Height          =   225
            Index           =   13
            Left            =   2130
            TabIndex        =   32
            Top             =   3000
            Width           =   315
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "INTERVALO:"
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
            Height          =   225
            Index           =   12
            Left            =   600
            TabIndex        =   31
            Top             =   3000
            Width           =   1035
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Empezar a potear Azul después de:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   11
            Left            =   45
            TabIndex        =   29
            Top             =   2640
            Width           =   3105
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MS"
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
            Height          =   225
            Index           =   10
            Left            =   2130
            TabIndex        =   28
            Top             =   2280
            Width           =   315
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "INTERVALO:"
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
            Height          =   225
            Index           =   9
            Left            =   600
            TabIndex        =   27
            Top             =   2280
            Width           =   1035
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Empezar a potear Rojas después de:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   8
            Left            =   15
            TabIndex        =   25
            Top             =   1920
            Width           =   3135
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            Index           =   3
            X1              =   0
            X2              =   3360
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MS"
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
            Height          =   225
            Index           =   7
            Left            =   2160
            TabIndex        =   24
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "INTERVALO:"
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
            Height          =   225
            Index           =   6
            Left            =   600
            TabIndex        =   23
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H000000C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Intervalo de Poción ROJA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   5
            Left            =   360
            TabIndex        =   22
            Top             =   480
            Width           =   2505
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "[Auto Pociones]"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   3225
         End
      End
      Begin VB.Label Intervalos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Guardar config"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1680
         TabIndex        =   1
         Top             =   5280
         Width           =   3615
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración general"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   1920
      TabIndex        =   4
      Top             =   0
      Width           =   3150
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label cmdDat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "FrmIntervalos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Begin Code

Option Explicit
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2 '
Private Const APPLICATION As String = "CONFIG"
Private Declare Function SetWindowPos _
    Lib "User32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
'-  -   -   -   -   -   -   -   -   -   -   -   -       -   -   -   -   -   -   -

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
  
'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
    Private Function Leer_Ini(Path_INI As String, key As String, Default As Variant) As String
  
Dim bufer As String * 256
Dim Len_Value As Long
  
        Len_Value = GetPrivateProfileString(APPLICATION, _
                                         key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
          
        Leer_Ini = Left$(bufer, Len_Value)
  
End Function
  
'Escribe un dato en el INI _
-----------------------------
'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave
  
Private Function Grabar_Ini(Path_INI As String, key As String, Valor As Variant) As String
  
    WritePrivateProfileString APPLICATION, _
                                         key, _
                                         Valor, _
                                         Path_INI
  
End Function


Private Sub cmdDat_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub cModoMacros_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
cAimBot.Value = "0"
cModoMacros.Enabled = False
cAimBot.Enabled = True

End Sub

Private Sub cAimBot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
cModoMacros.Value = "0"
cAimBot.Enabled = False
cModoMacros.Enabled = True

End Sub

Private Sub Form_Load()

Call LeerYNYS
Me.Visible = False
End Sub

Private Sub intervalAutoRemo_Change()

End Sub

Private Sub intervalAutoRemo_KeyPress(KeyAscii As Integer)
End Sub

Private Sub iAutoLanzar_KeyPress(KeyAscii As Integer)
If Not IsNumeric(iAutoLanzar.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0

End Sub

Private Sub iAutoRemo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(iAutoRemo.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0

End Sub

Private Sub iAzules_KeyPress(KeyAscii As Integer)
If Not IsNumeric(iAzules.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0

End Sub

Private Sub idAzules_KeyPress(KeyAscii As Integer)
If Not IsNumeric(idAzules.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0

End Sub

Private Sub idRojas_KeyPress(KeyAscii As Integer)
If Not IsNumeric(idRojas.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0

End Sub

Private Sub Intervalos_Click()
'guardar datos
Dim Path_Archivo_Ini As String
    Path_Archivo_Ini = App.Path & "\cfg.ini"
    Call Grabar_Ini(Path_Archivo_Ini, "cModoMacros", cModoMacros.Value)
                    
    Call Grabar_Ini(Path_Archivo_Ini, "cAimBot", cAimBot.Value)
    
    Call Grabar_Ini(Path_Archivo_Ini, "cAutoRemo", cAutoRemo.Value)
                            
    Call Grabar_Ini(Path_Archivo_Ini, "iAutoRemo", iAutoRemo.Text)
    
    Call Grabar_Ini(Path_Archivo_Ini, "iAutoLanzar", iAutoLanzar.Text)
                                    
    Call Grabar_Ini(Path_Archivo_Ini, "iRojas", iRojas.Text)
    
    Call Grabar_Ini(Path_Archivo_Ini, "iAzules", iAzules.Text)
                                            
    Call Grabar_Ini(Path_Archivo_Ini, "idRojas", idRojas.Text)
    
    Call Grabar_Ini(Path_Archivo_Ini, "idAzules", idAzules.Text)
                                                    
    Call Grabar_Ini(Path_Archivo_Ini, "ipRojas", ipRojas.Text)
    
    Call Grabar_Ini(Path_Archivo_Ini, "ipAzules", ipAzules.Text)

Call LeerYNYS
Call FormNoPrimerPlano(FrmIntervalos)
Unload Me
FrmMain.Visible = True
Inicial.Visible = True
Me.Visible = False


actual = "[Info: Intervalos OK]"
End Sub
Sub LeerYNYS()
Dim Path_Archivo_Ini As String
    Path_Archivo_Ini = App.Path & "\cfg.ini"


    vModoMacros = Leer_Ini(Path_Archivo_Ini, "cModoMacros", 1)
    vAimbot = Leer_Ini(Path_Archivo_Ini, "cAimBot", 0)
    vAutoRemo = Leer_Ini(Path_Archivo_Ini, "cAutoRemo", 0)
    
    IntervaloAutoRemo = Leer_Ini(Path_Archivo_Ini, "iAutoRemo", 300)
    
    IntervaloAutoLanzar = Leer_Ini(Path_Archivo_Ini, "iAutoLanzar", 750)
    
    IntervaloAutoRojas = Leer_Ini(Path_Archivo_Ini, "iRojas", 250)
    
    IntervaloAutoAzules = Leer_Ini(Path_Archivo_Ini, "iAzules", 200)
    
    DelayRojas = Leer_Ini(Path_Archivo_Ini, "idRojas", 100)

    DelayAzules = Leer_Ini(Path_Archivo_Ini, "idAzules", 90)
    
    PorcentajeRojas = Leer_Ini(Path_Archivo_Ini, "ipRojas", 70)
    
    PorcentajeAzules = Leer_Ini(Path_Archivo_Ini, "ipAzules", 30)

    
    cModoMacros.Value = vModoMacros
    
    cAimBot.Value = vAimbot
    
    cAutoRemo.Value = vAutoRemo
    
    iAutoRemo.Text = IntervaloAutoRemo
    
    iAutoLanzar.Text = IntervaloAutoLanzar
    
    iRojas.Text = IntervaloAutoRojas
    
    iAzules.Text = IntervaloAutoAzules
    
    idRojas.Text = DelayRojas
    
    idAzules.Text = DelayAzules
    
    ipRojas.Text = PorcentajeRojas
    
    ipAzules.Text = PorcentajeAzules

    If vModoMacros = "1" Then
    cAimBot.Value = "0"
    cModoMacros.Value = vbChecked
    cModoMacros.Enabled = False
    cAimBot.Enabled = True
    
    End If
    
    If vAimbot = "1" Then
    cModoMacros.Value = "0"
    cAimBot.Value = vbChecked
    cAimBot.Enabled = False
    cModoMacros.Enabled = True
    
    End If

End Sub


Private Sub ipAzules_KeyPress(KeyAscii As Integer)
If Not IsNumeric(ipAzules.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0

End Sub

Private Sub ipRojas_KeyPress(KeyAscii As Integer)
If Not IsNumeric(ipRojas.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0

End Sub

Private Sub iRojas_KeyPress(KeyAscii As Integer)
If Not IsNumeric(iRojas.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0

End Sub

Private Sub Label4_Click()
iAutoLanzar.Text = iAutoLanzar.Text + 50
End Sub

Private Sub Label6_Click()
Me.Visible = False
vIntervalos = "A"
frmInfo.Show
frmInfo.txtData.Text = "Esta sección es para modificar el intervalo de lanzar y tomar pociones" & vbNewLine & _
"Mientras mas alto el valor, mas lento se lanza y se potea" & vbNewLine & _
"Usa la opcion que mas te guste!"
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Label7_Click()
iAutoLanzar.Text = iAutoLanzar.Text - 50
If iAutoLanzar.Text < 350 Then iAutoLanzar.Text = 350
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub txtAutoLanzar_Change()

End Sub


