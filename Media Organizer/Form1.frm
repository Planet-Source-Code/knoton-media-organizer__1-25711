VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "   Knoton´s Media Organizer"
   ClientHeight    =   6015
   ClientLeft      =   2940
   ClientTop       =   1305
   ClientWidth     =   6630
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6630
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483638
      ForeColor       =   -2147483637
      TabCaption(0)   =   "List and Search Media"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label17"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label18(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblFileInfo(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblFileInfo(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblFileInfo(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblFileInfo(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "noOfFile"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label19"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblFileInfo(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblFileInfo(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label18(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lstMedia"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lstFiles"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtSearch"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdSearch(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdSearch(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdRunFile"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Scan and Save Media"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstReScanMediaId"
      Tab(1).Control(1)=   "Option1(2)"
      Tab(1).Control(2)=   "cmdClear"
      Tab(1).Control(3)=   "List1(1)"
      Tab(1).Control(4)=   "cmdSelAll"
      Tab(1).Control(5)=   "cmdAddDefExt"
      Tab(1).Control(6)=   "txtExt"
      Tab(1).Control(7)=   "lstselExt"
      Tab(1).Control(8)=   "lstExt(5)"
      Tab(1).Control(9)=   "lstExt(4)"
      Tab(1).Control(10)=   "lstExt(3)"
      Tab(1).Control(11)=   "lstExt(2)"
      Tab(1).Control(12)=   "lstExt(1)"
      Tab(1).Control(13)=   "lstExt(0)"
      Tab(1).Control(14)=   "Option1(1)"
      Tab(1).Control(15)=   "Option1(0)"
      Tab(1).Control(16)=   "txtInfo"
      Tab(1).Control(17)=   "cmdSaveToDB"
      Tab(1).Control(18)=   "List1(0)"
      Tab(1).Control(19)=   "Label14"
      Tab(1).Control(20)=   "Label22"
      Tab(1).Control(21)=   "Label13"
      Tab(1).Control(22)=   "Label12"
      Tab(1).Control(23)=   "Label11(5)"
      Tab(1).Control(24)=   "Label11(4)"
      Tab(1).Control(25)=   "Label11(3)"
      Tab(1).Control(26)=   "Label11(2)"
      Tab(1).Control(27)=   "Label11(1)"
      Tab(1).Control(28)=   "Label11(0)"
      Tab(1).Control(29)=   "Label8"
      Tab(1).Control(30)=   "Label7"
      Tab(1).Control(31)=   "lblNoFiles"
      Tab(1).Control(32)=   "Label5"
      Tab(1).ControlCount=   33
      TabCaption(2)   =   "Backup/Restore/Delete"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CD"
      Tab(2).Control(1)=   "cmdRestoreDB"
      Tab(2).Control(2)=   "cmdBackupDB"
      Tab(2).Control(3)=   "cmdDelDB"
      Tab(2).Control(4)=   "lstDelMedia"
      Tab(2).Control(5)=   "Label21"
      Tab(2).Control(6)=   "Label20"
      Tab(2).Control(7)=   "Label15"
      Tab(2).Control(8)=   "Label10"
      Tab(2).ControlCount=   9
      Begin VB.ListBox lstReScanMediaId 
         Height          =   1035
         Left            =   -70680
         TabIndex        =   60
         ToolTipText     =   "Rescan selected media"
         Top             =   720
         Width           =   2175
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   -71040
         Top             =   5400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Database (*.mdb)|*.mdb"
      End
      Begin VB.CommandButton cmdRestoreDB 
         Caption         =   "Restore Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70200
         TabIndex        =   59
         ToolTipText     =   "Restore the database with the backed up one"
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CommandButton cmdBackupDB 
         Caption         =   "Backup Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70200
         TabIndex        =   58
         ToolTipText     =   "Backup the database to a place or your choice"
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelDB 
         Caption         =   "Delete all media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70200
         TabIndex        =   55
         ToolTipText     =   "Delete all media and get fresh database"
         Top             =   4440
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Floppy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -72480
         TabIndex        =   52
         ToolTipText     =   "Mediatype floppy"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdRunFile 
         Caption         =   "Run selected file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear all"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   41
         ToolTipText     =   "Clear the selected formats to scan for"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ListBox List1 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   1
         Left            =   -74880
         TabIndex        =   40
         Top             =   4800
         Width           =   6375
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "Select all"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   39
         ToolTipText     =   "Scan for all files"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddDefExt 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74280
         TabIndex        =   38
         ToolTipText     =   "Add format that are not listed in the various listboxes"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtExt 
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   -74880
         TabIndex        =   36
         Top             =   1920
         Width           =   495
      End
      Begin VB.ListBox lstselExt 
         ForeColor       =   &H80000007&
         Height          =   840
         Left            =   -74880
         TabIndex        =   35
         ToolTipText     =   "All formats to scan for, click to remove unwanted ones"
         Top             =   720
         Width           =   975
      End
      Begin VB.ListBox lstExt 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   5
         Left            =   -71760
         TabIndex        =   27
         ToolTipText     =   "Various executable formats, including compressed formats"
         Top             =   2160
         Width           =   855
      End
      Begin VB.ListBox lstExt 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   4
         Left            =   -72720
         TabIndex        =   26
         ToolTipText     =   "Various web formats"
         Top             =   2160
         Width           =   855
      End
      Begin VB.ListBox lstExt 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   3
         Left            =   -73680
         TabIndex        =   25
         ToolTipText     =   "Various text formats"
         Top             =   2160
         Width           =   855
      End
      Begin VB.ListBox lstExt 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   2
         Left            =   -71760
         TabIndex        =   24
         ToolTipText     =   "Various picture formats"
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox lstExt 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   1
         Left            =   -72720
         TabIndex        =   23
         ToolTipText     =   "Various audio formats"
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox lstExt 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   0
         Left            =   -73680
         TabIndex        =   22
         ToolTipText     =   "Various Video formats"
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox lstDelMedia 
         Height          =   3570
         Left            =   -74880
         TabIndex        =   20
         ToolTipText     =   "Delete the selected media"
         Top             =   720
         Width           =   6375
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search all"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   960
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Search the entire database"
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search Media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Search the listed media"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtSearch 
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "You can use % as wildcard"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -73080
         TabIndex        =   13
         ToolTipText     =   "Mediatype CD-ROM"
         Top             =   3600
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -73680
         TabIndex        =   12
         ToolTipText     =   "Mediatype hard drive partition"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox txtInfo 
         ForeColor       =   &H80000007&
         Height          =   525
         Left            =   -73680
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Insert any info you want to put in the database about the media"
         Top             =   4080
         Width           =   2415
      End
      Begin VB.ListBox lstFiles 
         ForeColor       =   &H80000007&
         Height          =   2205
         ItemData        =   "Form1.frx":0496
         Left            =   1920
         List            =   "Form1.frx":049D
         TabIndex        =   4
         Top             =   2040
         Width           =   4575
      End
      Begin VB.ListBox lstMedia 
         ForeColor       =   &H80000007&
         Height          =   1035
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   6375
      End
      Begin VB.CommandButton cmdSaveToDB 
         Caption         =   "Save media"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71160
         TabIndex        =   2
         ToolTipText     =   "Save the result of the scan to database"
         Top             =   4080
         Width           =   855
      End
      Begin VB.ListBox List1 
         ForeColor       =   &H80000007&
         Height          =   1035
         Index           =   0
         Left            =   -74880
         TabIndex        =   1
         ToolTipText     =   "Select the drive to scan and start the scan"
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Restore the database from your backup."
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Label Label20 
         Caption         =   "Backup the database to any place and name to be kept safely."
         Height          =   495
         Left            =   -74880
         TabIndex        =   63
         Top             =   4920
         Width           =   3975
      End
      Begin VB.Label Label14 
         Caption         =   $"Form1.frx":04AB
         Height          =   1815
         Left            =   -70680
         TabIndex        =   62
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label22 
         Caption         =   "ReScan Media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70680
         TabIndex        =   61
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "Media Info:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label lblFileInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1200
         TabIndex        =   56
         Top             =   5640
         UseMnemonic     =   0   'False
         Width           =   5325
      End
      Begin VB.Label lblFileInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   54
         Top             =   4440
         Width           =   5295
      End
      Begin VB.Label Label19 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label noOfFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   51
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblFileInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1200
         TabIndex        =   50
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblFileInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   49
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label lblFileInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   48
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label lblFileInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   47
         Top             =   4680
         Width           =   5325
      End
      Begin VB.Label Label18 
         Caption         =   "MediaType:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "MediaId:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Select the media you want to delete. Or delete all media. This operation can´t be undone."
         Height          =   375
         Left            =   -74880
         TabIndex        =   42
         Top             =   4440
         Width           =   3975
      End
      Begin VB.Label Label13 
         Caption         =   "Add format"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Scan list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Executable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -71760
         TabIndex        =   33
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Web"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -72720
         TabIndex        =   32
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -73680
         TabIndex        =   31
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -71760
         TabIndex        =   30
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Audio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -72720
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Video"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -73680
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Select the media you want to delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "Search for file/s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Insert info about Media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73680
         TabIndex        =   15
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Select mediatype"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73680
         TabIndex        =   14
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label lblNoFiles 
         Caption         =   "No of files: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70200
         TabIndex        =   10
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Select Drive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Select a file for more info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Files found:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Select a media to list from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************COPYRIGHT INFO****************************'
'***    This application is developed by Kenneth Hedman         ***'
'***    Rights are given to anyone who wish to use parts        ***'
'***    of the code, as long as there are no commercial         ***'
'***    interest of any kind in the end product.                ***'
'******************************************************************'

'**************************CONTACT INFO****************************'
'***    Name:       Kenneth Hedman                              ***'
'***    Adress:     Jungfruv 165 b                              ***'
'***    City:       Falun                                       ***'
'***    Country:    Sweden                                      ***'
'***    Email:      knoton@hotmail.com                          ***'
'***    web:        http://www.knoton.dns2go.com                ***'
'******************************************************************'

'**************************CREDITS*********************************'
'***    http://www.planet-source-code.com and the people on it  ***'
'***    who makes it the best programming community there is    ***'
'***    http://www.allapi.net                                   ***'
'***    For the help/information about the API calls            ***'
'******************************************************************'

Option Explicit

Dim objRsFiles As ADODB.Recordset       'Used to get/alter filenames,path,size...
Dim objRsMedia As ADODB.Recordset       'Used to get/alter MediaId,Mediatype,Info...
Dim fso As Scripting.FileSystemObject   'Used to Check if file is existing
Dim conString As String                 'Holds the ConnectionString
Dim VarMediaType As String              'Holds the Mediatype to be stored in db
Dim intMediaId As Integer               'Holds what media to select and search, and also what files to delete during rescan
Dim strMediaType As String              'Holds Mediatype to show in info about file, translate CD to CD-rom
Dim strRunFile As String                'Holds the path to the file to run
Dim bolReScan As Boolean                'Holds if a rescan is in progress
Dim bolClear As Boolean                 'Holds if clear settings is in progress
Dim strExtensions As String             'Holds the formats/extensions to be saved, used during rescan

'*** API to get a list of all drives on the system ***'
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Sub Form_Load()
Set objRsMedia = New ADODB.Recordset
Set objRsFiles = New ADODB.Recordset
Set fso = New Scripting.FileSystemObject
VarMediaType = "CD"
conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & _
            "\FileDb.mdb;Persist Security Info=False"
ShowAllMedia
GetAllDrives
addExtensions
End Sub

Private Sub GetAllDrives()
Dim i As Integer
Dim listDrive As Long
Dim strSave As String
strSave = String(255, Chr(0))
listDrive = GetLogicalDriveStrings(255, strSave) 'get drives

For i = 1 To 100 ' split the string with drives and list it
    If Left(strSave, InStr(1, strSave, Chr(0))) = Chr(0) Then Exit For
    List1(0).AddItem Left(strSave, InStr(1, strSave, Chr(0)) - 1)
    strSave = Right(strSave, Len(strSave) - InStr(1, strSave, Chr(0)))
Next

End Sub

Private Sub ShowAllMedia()
If objRsFiles.State <> adStateClosed Then objRsFiles.Close

With objRsFiles
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = "Select * from tblMedia"
    .Open
    
'clear various listboxes and labels
lstReScanMediaId.Clear
lstMedia.Clear
lstDelMedia.Clear
lstFiles.Clear
noOfFile.Caption = ""
ClearFileInfo

If .RecordCount <> 0 Then 'List medias in listboxes
    .MoveFirst
    While Not .EOF
        lstMedia.AddItem .Fields("MediaId") & _
        ", " & .Fields("MediaType") & _
        ", " & .Fields("Info")
        lstDelMedia.AddItem .Fields("MediaId") & _
        ", " & .Fields("MediaType") & _
        ", " & .Fields("Info")
        lstReScanMediaId.AddItem .Fields("MediaId") & _
        ", " & .Fields("Info")
        .MoveNext
    Wend
Else 'No medias found
    MsgBox "No Medias are saved to Database"
    SSTab1.Tab = 1
End If
End With
End Sub

Private Sub ShowFilesInSelectedUnit(varMediaId As Integer)
Dim i As Integer
Screen.MousePointer = vbHourglass
If objRsFiles.State <> adStateClosed Then objRsFiles.Close

'Get records that matches the MediaId asked for
With objRsFiles
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = "Select * from tblFileName where MediaId =" & varMediaId
    .Open

    .MoveFirst
    
'List the result
lstFiles.Clear
If .RecordCount > 32767 Then
MsgBox "The records are to many to list !" & vbCrLf & _
        "I will show the first 32767 files of: " & .RecordCount & vbCrLf & _
        "Please limit the listing by searching" & vbCrLf & _
        "for what you want."
        For i = 0 To 32766
            lstFiles.AddItem .Fields("Filename")
            .MoveNext
        Next
Else
    While Not .EOF
        lstFiles.AddItem .Fields("FileName")
        .MoveNext
    Wend
End If
End With
Screen.MousePointer = vbDefault
noOfFile.Caption = objRsFiles.RecordCount
End Sub

Private Sub AddFilesToDB()
Dim i As Long
Screen.MousePointer = vbHourglass
If objRsMedia.State <> adStateClosed Then objRsMedia.Close
'Open db and add/alter media
With objRsMedia
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
Select Case bolReScan
    Case False
    .Source = "Select * from tblMedia"
    Case True
    .Source = "Select * from tblMedia where MediaId =" & intMediaId
End Select
    .Open
'Add the data from the scan to the database
Select Case bolReScan
    Case False
    .AddNew
End Select
    .Fields("MediaType") = VarMediaType
    .Fields("Info") = txtInfo.Text
    .Fields("Extensions") = strExtensions
    .Update
    .Requery
    .MoveLast
End With
strExtensions = ""
If objRsFiles.State <> adStateClosed Then objRsFiles.Close
'Delete files that match the mediaId you have rescan
With objRsFiles
Select Case bolReScan
    Case True
        .ActiveConnection = conString
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Source = "Delete * from tblFileName where MediaId =" & intMediaId
        .Open
End Select

If objRsFiles.State <> adStateClosed Then objRsFiles.Close
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = "Select * from tblFileName"
    .Open
    
'Add the data from the scan to the database
For i = 1 To UBound(ReturnFileName)
    .AddNew
Select Case bolReScan
    Case False
    .Fields("MediaID") = objRsMedia.Fields("mediaId")
    Case True
    .Fields("MediaID") = intMediaId
End Select
    .Fields("FileName") = ReturnFileName(i)
    .Fields("FileSize") = ReturnFileSize(i)
    .Fields("Path") = ReturnPath(i)
    .Update
Next
End With
If bolReScan = True Then CompressDB

Screen.MousePointer = vbDefault
If bolReScan = True Then bolReScan = False
End Sub

Private Sub clearAll()
'Clear various settings
bolClear = True
ClearSelFormatlst
lstselExt.Clear
List1(0).ListIndex = -1
Option1(1).Value = True
bolReScan = False
cmdSaveToDB.Enabled = False
txtInfo.Text = ""
txtExt.Text = ""
lblNoFiles.Caption = "No of files: "
bolClear = False
End Sub

Private Sub cmdClear_Click()
clearAll
End Sub

Private Sub cmdAddDefExt_Click()
If txtExt.Text <> "" Then lstselExt.AddItem txtExt.Text
End Sub

Private Sub cmdRunFile_Click()
Dim i As Integer
Dim strTemp As String
'Run the file
If fso.FileExists(strRunFile) = True Then
    RunFile strRunFile, Me
Else
'If file is not found on saved path search for it
For i = 2 To List1(0).ListCount - 1 'Skip the floppy drive
    strTemp = List1(0).List(i) & Right(strRunFile, Len(strRunFile) - 3)
    If fso.FileExists(strTemp) = True Then
        RunFile strTemp, Me
        Exit Sub
    End If
Next
    'File might not be able to run or media is not in any cd-rom
    MsgBox "File is not aviable !" & vbCrLf & _
            "Try to insert the media"
End If

End Sub

Private Sub cmdSaveToDB_Click()
If txtInfo.Text = "" Then
    If MsgBox("Do you want to add this without inserting any info " & _
            "about the media ? That info vill be visible in " & _
            "the media info, could help determine what media it is", vbYesNo) = vbNo Then
        txtInfo.SetFocus
        Exit Sub
    End If
End If
AddFilesToDB
ShowAllMedia
clearAll
End Sub

Private Sub cmdSearch_Click(Index As Integer)
Dim whereString As String
Dim whereMediaId As String
Dim i As Integer

If bolReScan = True Then
    bolReScan = False
    clearAll
End If
lstFiles.Clear
ClearFileInfo

If intMediaId = 0 And Index = 0 Then
    MsgBox "You must choose a media !"
    Exit Sub
End If

'Get all filenames that contains the combination of letters in txtsearch
whereString = " Where Filename  Like " & "'" & txtSearch.Text & "%'"

whereMediaId = " and mediaid = " & intMediaId
If objRsFiles.State <> adStateClosed Then objRsFiles.Close
Screen.MousePointer = vbHourglass

With objRsFiles
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
Select Case Index
    Case 0
    .Source = "Select * from tblFilename" & whereString & whereMediaId & _
              " Order by Filename"
    Case 1
    .Source = "Select * from tblFilename" & whereString & " Order by Mediaid"
End Select
    .Open

noOfFile.Caption = .RecordCount

If .RecordCount = 0 Then
    MsgBox "No records found !"
Else
    .MoveFirst

If .RecordCount > 32767 Then
MsgBox "The records are to many to list !" & vbCrLf & _
        "I will show the first 32767 files of: " & .RecordCount & vbCrLf & _
        "Please limit the listing by searching" & vbCrLf & _
        "for what you want."
    For i = 0 To 32766
        Select Case Index
            Case 0
                lstFiles.AddItem .Fields("Filename")
            Case 1
                lstFiles.AddItem "Id = " & .Fields("MediaId") & _
                                       ", " & .Fields("FileName")
        End Select
                .MoveNext
    Next
Else
    While Not .EOF
        Select Case Index
            Case 0
                lstFiles.AddItem .Fields("FileName")
            Case 1
                lstFiles.AddItem "Id = " & .Fields("MediaId") & _
                                       ", " & .Fields("FileName")
        End Select
            .MoveNext
    Wend
End If
End If
End With
txtSearch.Text = ""
txtSearch.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelAll_Click()
'All formats
lstselExt.Clear
lstselExt.AddItem "*"
End Sub

Private Sub cmdDelDB_Click()
If MsgBox("Are you sure you want to delete all media in Database ?" & _
            "This operation cant be undone.", vbYesNo) = vbYes Then
Set objRsFiles = Nothing
Set objRsMedia = Nothing
Kill App.Path & "\filedb.mdb"
FileCopy App.Path & "\filedb.bak", App.Path & "\filedb.mdb"
Set objRsFiles = New ADODB.Recordset
Set objRsMedia = New ADODB.Recordset
ShowAllMedia
clearAll
End If
End Sub

Private Sub cmdBackupDB_Click()
Set objRsFiles = Nothing
Set objRsMedia = Nothing
CD.DialogTitle = "Backup Knoton´s Media Organizer Database as"
CD.ShowSave
If CD.FileName <> "" Then FileCopy App.Path & "/Filedb.mdb", CD.FileName
Set objRsFiles = New ADODB.Recordset
Set objRsMedia = New ADODB.Recordset
ShowAllMedia
End Sub

Private Sub cmdRestoreDB_Click()
Set objRsFiles = Nothing
Set objRsMedia = Nothing
CD.DialogTitle = "Restore Knoton´s Media Organizer Database as"
CD.ShowOpen
If CD.FileName <> "" Then
    Kill App.Path & "\filedb.mdb"
    FileCopy CD.FileName, App.Path & "/Filedb.mdb"
End If
Set objRsFiles = New ADODB.Recordset
Set objRsMedia = New ADODB.Recordset
ShowAllMedia
End Sub

Private Sub lblFileInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
lblFileInfo(Index).ToolTipText = lblFileInfo(Index).Caption
End Sub

Private Sub List1_Click(Index As Integer)
Select Case Index
    Case 0
Dim i As Integer
Dim x As Long
If bolClear = False Then
    List1(1).Clear
    noOfFiles = 0
    cmdSaveToDB.Enabled = False
    If lstselExt.ListCount > 0 Then
        For i = 0 To lstselExt.ListCount - 1
        ReDim Preserve strExt(i)
        strExt(i) = lstselExt.List(i)
        strExtensions = strExtensions & lstselExt.List(i) & " "
        Next
        
        listAll List1(0).List(List1(0).ListIndex)
        
        strExtensions = Mid(strExtensions, 1, Len(strExtensions) - 1)
        
        If noOfFiles = 0 Then
            Screen.MousePointer = vbDefault
            clearAll
            MsgBox "No files found"
            Exit Sub
        End If
        
        If MsgBox("Do you want to list " & noOfFiles & _
        " Files before saving, it might take a while ?", vbYesNo) = vbYes Then
            Screen.MousePointer = vbHourglass
            
            If noOfFiles > 32767 Then
            MsgBox "The files are to many to list !" & vbCrLf & _
                    "I will show the first 32767 files of: " & noOfFiles
                For x = 1 To 32767
                    List1(1).AddItem ReturnPath(x) & ReturnFileName(x)
                Next
            Else
                For x = 1 To noOfFiles
                    List1(1).AddItem ReturnPath(x) & ReturnFileName(x)
                Next
            End If
            
        Screen.MousePointer = vbDefault
        End If
        cmdSaveToDB.Enabled = True
    Else
        MsgBox "You must select what formats to scan for first!"
    End If
End If
End Select
End Sub
Private Sub listAll(SearchPath As String)
'Get all files matching the formats to scan for
Screen.MousePointer = vbHourglass
Call FindFiles(SearchPath, "*.*")
lblNoFiles.Caption = "No of files: " & noOfFiles
Screen.MousePointer = vbDefault
End Sub

Private Sub lstDelMedia_Click()
Dim i As Integer
Dim strTemp As String
bolReScan = False
i = 1
While Mid(lstDelMedia.List(lstDelMedia.ListIndex), i, 1) <> ","
    strTemp = Mid(lstDelMedia.List(lstDelMedia.ListIndex), 1, i)
    i = i + 1
Wend
If MsgBox("Are you sure you want to delete MediaId " & strTemp, vbYesNo, "Delete Media") = vbYes Then
Screen.MousePointer = vbHourglass

If objRsFiles.State <> adStateClosed Then objRsFiles.Close

With objRsFiles
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = "Delete * from tblFileName where MediaId =" & CInt(strTemp)
    .Open
End With

If objRsMedia.State <> adStateClosed Then objRsMedia.Close

With objRsMedia
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = "Delete * from tblMedia where mediaId =" & CInt(strTemp)
    .Open
End With

If objRsFiles.State <> adStateClosed Then objRsFiles.Close
If objRsMedia.State <> adStateClosed Then objRsMedia.Close
CompressDB
ShowAllMedia
clearAll
Screen.MousePointer = vbDefault
End If
End Sub

Private Sub lstExt_Click(Index As Integer)
Dim i As Integer
For i = 0 To lstselExt.ListCount - 1
If lstselExt.List(i) = lstExt(Index).List(lstExt(Index).ListIndex) Then
    Exit Sub
End If
Next
    lstselExt.AddItem lstExt(Index).List(lstExt(Index).ListIndex)
End Sub

Private Sub lstFiles_Click()
objRsFiles.MoveFirst
objRsFiles.Move lstFiles.ListIndex
strRunFile = objRsFiles.Fields("path") & objRsFiles.Fields("FileName")
SetFileInfo
End Sub

Private Sub SetFileInfo()
GetInfoMedia
If strMediaType = "CD" Then strMediaType = "CD-Rom"
If strMediaType = "HD" Then strMediaType = "Hard Drive"
lblFileInfo(0).Caption = objRsFiles.Fields("Path")
lblFileInfo(4).Caption = objRsFiles.Fields("FileName")
lblFileInfo(1).Caption = objRsFiles.Fields("FileSize")
lblFileInfo(2).Caption = objRsFiles.Fields("MediaId")
lblFileInfo(5).Caption = objRsMedia.Fields("Info")
lblFileInfo(3).Caption = strMediaType
End Sub

Private Sub ClearFileInfo()
Dim i As Integer
For i = 0 To 5
    lblFileInfo(i).Caption = ""
Next
End Sub

Private Sub GetInfoMedia()
Dim i As Long
If objRsMedia.State <> adStateClosed Then objRsMedia.Close

With objRsMedia
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = "Select * from tblMedia where MediaId = " & objRsFiles.Fields("MediaId")
    .Open
End With
strMediaType = objRsMedia.Fields("MediaType")
End Sub

Private Sub lstFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
lstFiles.ToolTipText = lstFiles.List(lstFiles.ListIndex)
End Sub

Private Sub lstMedia_Click()
Dim i As Integer
Dim strTemp As String

If bolReScan = True Then
    bolReScan = False
    clearAll
End If

i = 1
While Mid(lstMedia.List(lstMedia.ListIndex), i, 1) <> ","
    strTemp = Mid(lstMedia.List(lstMedia.ListIndex), 1, i)
    i = i + 1
Wend

intMediaId = CInt(strTemp)
ClearFileInfo
ShowFilesInSelectedUnit (intMediaId)
End Sub

Private Sub lstReScanMediaId_Click()
Dim i, x As Integer
Dim strTemp As String
Dim tmpExtensions() As String

bolReScan = True

i = 1
While Mid(lstReScanMediaId.List(lstReScanMediaId.ListIndex), i, 1) <> ","
    strTemp = Mid(lstReScanMediaId.List(lstReScanMediaId.ListIndex), 1, i)
    i = i + 1
Wend
intMediaId = CInt(strTemp)


If objRsMedia.State <> adStateClosed Then objRsMedia.Close

With objRsMedia
    .ActiveConnection = conString
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = "Select * from tblMedia where mediaId =" & intMediaId
    .Open
    lstselExt.Clear
    tmpExtensions = Split(.Fields("Extensions"), " ", -1, 1)
    
    For i = 0 To UBound(tmpExtensions)
        lstselExt.AddItem tmpExtensions(i)
    Next
    
    txtInfo.Text = .Fields("Info")
    
    For i = 0 To 2
        If Option1(i).Caption = .Fields("MediaType") Then Option1(i).Value = True
    Next
    
End With
End Sub

Private Sub lstselExt_Click()
lstselExt.RemoveItem (lstselExt.ListIndex)
End Sub

Private Sub Option1_Click(Index As Integer)
VarMediaType = Option1(Index).Caption
End Sub

Private Sub CompressDB()
Dim objCompress As JetEngine
Set objCompress = New JetEngine
Set objRsFiles = Nothing
Set objRsMedia = Nothing

If Dir(App.Path & "\filedb2.mdb") <> "" Then Kill _
        App.Path & "\filedb2.mdb"

objCompress.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & App.Path & "\filedb.mdb;" & _
        "Jet OLEDB:Engine Type=3", _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & App.Path & "\filedb2.mdb;" & _
        "Jet OLEDB:Engine Type=3"
        
Kill App.Path & "\filedb.mdb"
FileCopy App.Path & "\filedb2.mdb", App.Path & "\filedb.mdb"
Kill App.Path & "\filedb2.mdb"
Kill App.Path & "\filedb2.ldb"
Set objRsFiles = New ADODB.Recordset
Set objRsMedia = New ADODB.Recordset

End Sub

Private Sub ClearSelFormatlst()
Dim i As Integer
For i = 0 To 5
    lstExt(i).ListIndex = -1
Next
End Sub

Private Sub addExtensions()
With lstExt(0)
    .AddItem "asf"
    .AddItem "avi"
    .AddItem "mpe"
    .AddItem "mpeg"
    .AddItem "mpg"
    .AddItem "mov"
    .AddItem "wmv"
End With

With lstExt(1)
    .AddItem "cda"
    .AddItem "mid"
    .AddItem "midi"
    .AddItem "mp3"
    .AddItem "wav"
    .AddItem "wma"
End With

With lstExt(2)
    .AddItem "bmp"
    .AddItem "cdr"
    .AddItem "cur"
    .AddItem "drw"
    .AddItem "gif"
    .AddItem "jpeg"
    .AddItem "jpg"
    .AddItem "ico"
    .AddItem "mix"
    .AddItem "pcd"
    .AddItem "pcx"
    .AddItem "png"
    .AddItem "psd"
    .AddItem "tif"
    .AddItem "tiff"
    .AddItem "wmf"
    
End With

With lstExt(3)
    .AddItem "doc"
    .AddItem "log"
    .AddItem "nfo"
    .AddItem "pdf"
    .AddItem "rtf"
    .AddItem "txt"
End With

With lstExt(4)
    .AddItem "asa"
    .AddItem "asp"
    .AddItem "chm"
    .AddItem "css"
    .AddItem "htm"
    .AddItem "html"
End With

With lstExt(5)
    .AddItem "bat"
    .AddItem "dll"
    .AddItem "exe"
    .AddItem "js"
    .AddItem "ocx"
    .AddItem "tlb"
    .AddItem "vbs"
    .AddItem "rar"
    .AddItem "zip"
End With
End Sub
