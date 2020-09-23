VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4185
   ClientLeft      =   225
   ClientTop       =   4905
   ClientWidth     =   13095
   Icon            =   "GeraPlaylist.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   6000
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancella m3u e txt"
      Height          =   285
      Left            =   8880
      TabIndex        =   5
      Top             =   3410
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Genera liste"
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   7005
   End
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   0
      TabIndex        =   7
      Top             =   -60
      Width           =   3495
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Scan"
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   3720
         Width           =   1065
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   2790
         Left            =   90
         TabIndex        =   0
         Top             =   540
         Width           =   3315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cd/Dvd N°"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   795
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Direct deletion!!!   Check before deleting"
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   8880
      TabIndex        =   13
      Top             =   3675
      Width           =   1740
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "www.ganzaborn.altervista.org"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10785
      TabIndex        =   6
      ToolTipText     =   "Close INFO and Link"
      Top             =   3435
      Width           =   2130
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   10800
      Picture         =   "GeraPlaylist.frx":1CCA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6000
      TabIndex        =   10
      Top             =   3435
      Width           =   60
   End
   Begin VB.Menu LINGUA 
      Caption         =   "LINGUA"
      Begin VB.Menu Bahasa_Indonesia 
         Caption         =   "Bahasa_Indonesia"
      End
      Begin VB.Menu Dansk 
         Caption         =   "Dansk"
      End
      Begin VB.Menu Deutsch 
         Caption         =   "Deutsch"
      End
      Begin VB.Menu English 
         Caption         =   "English"
      End
      Begin VB.Menu Español 
         Caption         =   "Español"
      End
      Begin VB.Menu Français 
         Caption         =   "Français"
      End
      Begin VB.Menu Italiano 
         Caption         =   "Italiano"
      End
      Begin VB.Menu Nederlands 
         Caption         =   "Nederlands"
      End
      Begin VB.Menu Polski 
         Caption         =   "Polski"
      End
      Begin VB.Menu Português 
         Caption         =   "Português"
      End
      Begin VB.Menu Suomi 
         Caption         =   "Suomi"
      End
   End
   Begin VB.Menu INFO 
      Caption         =   "INFO"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const DRIVE_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_NO_ROOT_DIR = 1
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_UNKNOWN = 0
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Dim XDir(2) As New Collection
Dim fso As New FileSystemObject
Dim BASEPATH As Folder

Sub ScanFolders(path As String)
Dim FLD As Folder
Dim FIL As File
SetAttr (path), vbNormal
Set BASEPATH = fso.GetFolder(path)
For Each FLD In BASEPATH.SubFolders
    ScanFolders FLD.path
    'For Each FIL In FLD.Files
    'List1.AddItem FIL.Path
        List1.AddItem FLD.path
        DoEvents
    'Next FIL
    DoEvents
Next FLD
End Sub

Sub ScanFolders3(path As String)
Dim FLD As Folder
Dim FIL As File
SetAttr (path), vbNormal
Set BASEPATH = fso.GetFolder(path)
For Each FLD In BASEPATH.SubFolders
    'ScanFolders FLD.path
    'For Each FIL In FLD.Files
    'List1.AddItem FIL.Path
        List1.AddItem FLD.path
        DoEvents
    'Next FIL
    DoEvents
Next FLD
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Audio" Then
Text1.Enabled = True
Command4.Enabled = True
Label4.Visible = True
File1.Pattern = "*.mp3;*.wma;*mid;*.mp1;*.mp2;*.mod;*.ogg;*.m4a;*.flac"
End If
If Combo1.Text = "Video" Then
Text1.Text = ""
Text1.Enabled = False
Command4.Enabled = False
Label4.Visible = False
File1.Pattern = "*.avi;*.asf;*.mpe;*.mpeg;*.mpg;*.mov;*.wmv;*.mp4"
End If
If Combo1.Text = "Nds - Gba" Then
Text1.Text = ""
Text1.Enabled = False
Command4.Enabled = False
Label4.Visible = False
File1.Pattern = "*.nds;*.gba;*.zip;*.7z;*.rar"
End If
If Combo1.Text = "Only Folders" Then
Text1.Text = ""
Text1.Enabled = False
Command4.Enabled = False
Label4.Visible = False
File1.Pattern = "*.*"
End If
End Sub

Sub ScanFolders2(path As Variant)
On Error Resume Next
    Dim fso As New FileSystemObject
    Dim BASEPATH2 As Folder
    Dim FLD2 As Folder
    Dim FIL2 As File
    Dim strSx2 As String
    Dim MySubDir2 As String
    Dim FileWayCopy2 As String
    Dim MyFile2 As Long
    Set BASEPATH2 = fso.GetFolder(path)
    For Each FLD2 In BASEPATH2.SubFolders
        ScanFolders2 FLD2.path
        'Acquisisce il Nome delle Cartelle:
        strSx2 = FLD2.Name
        ' Percorso completo della SubCartella:
        MySubDir2 = FLD2.path
        For Each FIL2 In FLD2.Files
            DoEvents
            'Percorso completo del file da copiare:
            FileWayCopy2 = FIL2
            'Acquisisce il valore dell'attributo dei files:
            MyFile2 = GetAttr(FileWayCopy2)
            'verifica il valore dell'attributo:
                If (MyFile2 > 32 And MyFile2 <> 38) Or MyFile2 = 1 Then
                    'Toglie l'attributo di sola lettura:
                    SetAttr FileWayCopy2, GetAttr(FileWayCopy2) And &HFFFE
                    'Conta il n° delle Cartelle:
                    intRisc2 = intRisc2 + 1
                    'Espone il Nome delle Cartelle:
                    strSx4 = strSx4 & "   " & FIL2
                    strSx4 = strSx4 & vbCrLf
                End If
        Next FIL2
            DoEvents
    Next FLD2
End Sub

Public Sub VerifyFile(Filename As String)
On Error Resume Next
'Apre lo specifico file
Open Filename For Input As #1
'Gestione dell'eventuale errore
If Err Then
MsgBox ("Il file " & Filename & " non è stato trovato.")
Exit Sub
End If
Close #1
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Dir1.path = Combo2.Text
Dir1.SetFocus
End Sub

Private Sub Dir1_Click()
List1.Clear
End Sub

Private Sub Command2_Click()
If Combo1.Text = "Audio" Then
If Text1.Text = "" Then
If List1.ListCount = 0 Then
SELEZIONARE
Else
Screen.MousePointer = 11
List1.Visible = False
 For X = 0 To List1.ListCount - 1
    List1.ListIndex = X
    vdest = Mid(List1.Text, InStrRev(List1.Text, "\") + 1, Len(List1.Text))
    Open List1.Text & "\" & vdest & ".m3u" For Output As #1
    Print #1, "#EXTM3U"
        For i = 0 To File1.ListCount - 1
            Print #1, File1.List(i)
        Next
    Close #1
Next
   e = Mid(Dir1.path, InStrRev(Dir1.path, "\") + 1, Len(Dir1.path))
    Open Dir1.path & "\" & ".txt" For Output As #3
            Print #3, e
    Close #3
For h = 0 To List1.ListCount - 1
    List1.ListIndex = h
    vdest2 = Mid(List1.Text, InStrRev(List1.Text, "\") + 1, Len(List1.Text))
    Open Dir1.path & "\" & vdest2 & ".txt" For Output As #4
            Print #4, vdest2
            Print #4, ""
        For K = 0 To File1.ListCount - 1
            Print #4, ClipExt(File1.List(K))
        Next
    Close #4
Next
TextFiles
End If
Else
If List1.ListCount = 0 Then
SELEZIONARE
Else
Screen.MousePointer = 11
List1.Visible = False
 For X = 0 To List1.ListCount - 1
    List1.ListIndex = X
    vdest = Mid(List1.Text, InStrRev(List1.Text, "\") + 1, Len(List1.Text))
    Open List1.Text & "\" & vdest & ".m3u" For Output As #1
    Print #1, "#EXTM3U"
        For i = 0 To File1.ListCount - 1
            Print #1, File1.List(i)
        Next
    Close #1
Next
   e = Mid(Dir1.path, InStrRev(Dir1.path, "\") + 1, Len(Dir1.path))
    Open Dir1.path & "\" & ".txt" For Output As #3
            Print #3, e; " "; "(" & Text1.Text & ")"
    Close #3
For h = 0 To List1.ListCount - 1
    List1.ListIndex = h
    vdest2 = Mid(List1.Text, InStrRev(List1.Text, "\") + 1, Len(List1.Text))
    Open Dir1.path & "\" & vdest2 & ".txt" For Output As #4
            Print #4, vdest2; " "; "("; e; ")"; "(" & Text1.Text & ")"
            Print #4, ""
        For K = 0 To File1.ListCount - 1
            Print #4, ClipExt(File1.List(K))
        Next
    Close #4
Next
TextFiles
End If
End If
List1.Visible = True
Screen.MousePointer = 0
End If
If Combo1.Text = "Video" Then
If Text1.Text = "" Then
If List1.ListCount = 0 Then
SELEZIONARE
Else
Screen.MousePointer = 11
List1.Visible = False
   e = Mid(Dir1.path, InStrRev(Dir1.path, "\") + 1, Len(Dir1.path))
    Open Dir1.path & "\" & ".txt" For Output As #3
            Print #3, e
    Close #3
For h = 0 To List1.ListCount - 1
    List1.ListIndex = h
    vdest2 = Mid(List1.Text, InStrRev(List1.Text, "\") + 1, Len(List1.Text))
    Open Dir1.path & "\" & vdest2 & ".txt" For Output As #4
            Print #4, vdest2
            Print #4, ""
        For K = 0 To File1.ListCount - 1
            Print #4, ClipExt(File1.List(K))
        Next
    Close #4
Next
TextFiles
End If
End If
List1.Visible = True
Screen.MousePointer = 0
End If
If Combo1.Text = "Nds - Gba" Then
If Text1.Text = "" Then
If List1.ListCount = 0 Then
SELEZIONARE
Else
Screen.MousePointer = 11
List1.Visible = False
   e = Mid(Dir1.path, InStrRev(Dir1.path, "\") + 1, Len(Dir1.path))
    Open Dir1.path & "\" & ".txt" For Output As #3
            Print #3, e
    Close #3
For h = 0 To List1.ListCount - 1
    List1.ListIndex = h
    vdest2 = Mid(List1.Text, InStrRev(List1.Text, "\") + 1, Len(List1.Text))
    Open Dir1.path & "\" & vdest2 & ".txt" For Output As #4
            Print #4, vdest2
            Print #4, ""
        For K = 0 To File1.ListCount - 1
            Print #4, ClipExt(File1.List(K))
        Next
    Close #4
Next
TextFiles
End If
End If
List1.Visible = True
Screen.MousePointer = 0
End If
If Combo1.Text = "Only Folders" Then
If Text1.Text = "" Then
If List1.ListCount = 0 Then
SELEZIONARE
Else
Screen.MousePointer = 11
List1.Visible = False
   e = Mid(Dir1.path, InStrRev(Dir1.path, "\") + 1, Len(Dir1.path))
    Open Dir1.path & "\" & ".txt" For Output As #3
            Print #3, e
    Close #3
For h = 0 To List1.ListCount - 1
    List1.ListIndex = h
    vdest2 = Mid(List1.Text, InStrRev(List1.Text, "\") + 1, Len(List1.Text))
    Open Dir1.path & "\" & vdest2 & ".txt" For Output As #4
            Print #4, vdest2
    Close #4
Next
TextFiles
End If
End If
List1.Visible = True
Screen.MousePointer = 0
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &H80000005
Label2.ForeColor = &H0&
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Buff As String
Dim DR() As String
Dim i As Long
'File1.Visible = False
Buff = Space$(50)
GetLogicalDriveStrings 50&, Buff
Buff = RTrim$(Buff)
Buff = Left$(Buff, Len(Buff) - 2)
DR = Split(Buff, vbNullChar)
For i = 0 To UBound(DR)
  If Drivetype(DR(i)) = "Removable Drive" Or _
     Drivetype(DR(i)) = "Fixed Drive" Or _
     Drivetype(DR(i)) = "Remote" Then
    Combo2.AddItem UCase(Left$(DR(i), 1) & ":")
  End If
Next
Form1.Width = 10785
Form1.Height = 4730
Me.Caption = "Lista mp3" & " v" & App.Major & "." & App.Minor & "." & App.Revision & " - " & "27/02/2009" & " - " & "16/02/2012"
Combo2.Text = UCase(Left$(Dir1.path, 1) & ":")
Combo1.AddItem "Audio"
Combo1.AddItem "Video"
Combo1.AddItem "Nds - Gba"
Combo1.AddItem "Only Folders"
Combo1.ListIndex = 0
If Combo1.Text = "Audio" Then
Text1.Enabled = True
Else
Text1.Enabled = False
End If
End Sub

Private Function Drivetype(ByVal DriveSpec As String) As String
Select Case GetDriveType(DriveSpec)
  Case 2
   Drivetype = "Removable Drive"
  Case 3
   Drivetype = "Fixed Drive"
  Case Is = 4
   Drivetype = "Remote"
  Case Is = 5
   Drivetype = "CD-ROM"
  Case Is = 6
   Drivetype = "Ram Disk"
  Case Else
   Drivetype = "Unrecognized"
End Select
End Function

Private Sub Command1_Click()
On Error Resume Next
List1.Clear
If Combo1.Text = "Only Folders" Then
ScanFolders3 (Dir1.path)
CARTROVATE
Else
ScanFolders2 (Dir1.path)
ScanFolders (Dir1.path)
CARTROVATE
End If
End Sub

Sub CARTROVATE()
If LINGUA.Caption = "LINGUA" Then
Label1.Caption = "Cartelle trovate: " & List1.ListCount
End If
If LINGUA.Caption = "LANGUAGE" Then
Label1.Caption = "Folders found: " & List1.ListCount
End If
If LINGUA.Caption = "BAHASA" Then
Label1.Caption = "Folder ditemukan: " & List1.ListCount
End If
If LINGUA.Caption = "SPROG" Then
Label1.Caption = "Mapper fundet: " & List1.ListCount
End If
If LINGUA.Caption = "SPRACHE" Then
Label1.Caption = "Ordner gefunden: " & List1.ListCount
End If
If LINGUA.Caption = "IDIOMA" Then
Label1.Caption = "Carpetas encontradas: " & List1.ListCount
End If
If LINGUA.Caption = "LANGUE" Then
Label1.Caption = "Dossiers trouvés: " & List1.ListCount
End If
If LINGUA.Caption = "TAAL" Then
Label1.Caption = "Mappen gevonden: " & List1.ListCount
End If
If LINGUA.Caption = "JEZYK" Then
Label1.Caption = "Znaleziono foldery: " & List1.ListCount
End If
If LINGUA.Caption = "LINGUAGEM" Then
Label1.Caption = "Pastas encontradas: " & List1.ListCount
End If
If LINGUA.Caption = "KIELI" Then
Label1.Caption = "Löytyneet kansiot: " & List1.ListCount
End If
End Sub

Sub SELEZIONARE()
If LINGUA.Caption = "BAHASA" Then
MsgBox "Pilih folder yang berisi MP3 dan tekan scan", vbExclamation, "DAFTAR KOSONG!!!"
End If
If LINGUA.Caption = "SPROG" Then
MsgBox "Vælg den mappe, der indeholder MP3 og tryk scan", vbExclamation, "LISTE TOM!!!"
End If
If LINGUA.Caption = "SPRACHE" Then
MsgBox "Wählen Sie den Ordner mit den MP3-und Presse-Scan", vbExclamation, " LISTE LEER!!!"
End If
If LINGUA.Caption = "LANGUAGE" Then
MsgBox "Select the folder containing the MP3 and press scan", vbExclamation, "LIST EMPTY!!!"
End If
If LINGUA.Caption = "IDIOMA" Then
MsgBox "Seleccione la carpeta que contiene los MP3 y pulse scan", vbExclamation, "LISTA VACIA!!!"
End If
If LINGUA.Caption = "LANGUE" Then
MsgBox "Sélectionnez le dossier contenant les MP3 et appuyez sur scan", vbExclamation, "LISTE VIDE!!!"
End If
If LINGUA.Caption = "LINGUA" Then
MsgBox "Selezionare la cartella contenente gli MP3 e premere scan", vbExclamation, "LISTA VUOTA!!!"
End If
If LINGUA.Caption = "TAAL" Then
MsgBox "Selecteer de map met de MP3-en druk-scan", vbExclamation, "LIJST LEEG!!!"
End If
If LINGUA.Caption = "JEZYK" Then
MsgBox "Wybierz folder zawierajacy MP3 i nacisnij scan", vbExclamation, "LISTA PUSTA!!!"
End If
If LINGUA.Caption = "LINGUAGEM" Then
MsgBox "Seleccione a pasta contendo os MP3 e pressione scan", vbExclamation, "LISTA VAZIA!!!"
End If
If LINGUA.Caption = "KIELI" Then
MsgBox "Valitse kansio, joka sisältää MP3 ja paina scan", vbExclamation, "LUETTELO TYHJÄ!!!"
End If
End Sub

Private Sub Command4_Click()
On Error GoTo No_File
If List1.ListCount = 0 Then
SELEZIONARE
Else
ScanFolders2 (Dir1.path)
DeleteDir (Dir1.path)
'Open Dir1.Path & "\Elimina.bat" For Output As #1
'Print #1, "Erase *.m3u /S /Q"
'Print #1, "Erase *.txt /S /Q"
 '  Close #1
'CommonDialog1.CancelError = True
'CommonDialog1.InitDir = Dir1.Path
'CommonDialog1.Filter = "All Files|*.bat"
'CommonDialog1.FileName = "Elimina.bat"
'EventPause 1.5
'CommonDialog1.ShowOpen
'Shell (Dir1.Path & "\Elimina.bat")
'EventPause 0.6
'Kill Dir1.Path & "\Elimina.bat"
Kill Dir1.path & "*.txt"
Exit Sub
No_File:
If Err.Number = cdlCancel Then
'Kill Dir1.Path & "\Elimina.bat"
MsgBox "Error", vbCritical
    Err.Clear
End If
End If
End Sub

Public Sub DeleteDir(path$)
    On Error Resume Next
    Dim vDirName As String, LastDir As String
    Dim z As New Collection, i As Integer
    'Adjust so No Deletion of Drive
    If Len(path$) < 4 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If Right(path$, 1) <> "\" Then path$ = path$ & "\"
    vDirName = Dir(path, vbDirectory) ' Retrieve the first entry.
    z.Add Mid(path$, 1, Len(path$) - 1)
    Do While vDirName <> ""
        If vDirName <> "." And vDirName <> ".." Then
            If (GetAttr(path & vDirName)) = vbDirectory Then
                LastDir = vDirName
                'Finds Directory Name then Repeats
                DeleteDir (path$ & vDirName)
                vDirName = Dir(path$, vbDirectory)
                Do Until vDirName = LastDir Or vDirName = ""
                    vDirName = Dir
                Loop
                If vDirName = "" Then Exit Do
            End If
        End If
        vDirName = Dir
    Loop
    Screen.MousePointer = vbDefault
    For i = 1 To z.Count
        'Deletes Files In Directories
         On Error Resume Next
        SetAttr z.Item(i) & "\*.*", vbNormal
        Kill z.Item(i) & "\*.m3u"
        Kill z.Item(i) & "\*.txt"
        'Deletes the Directories
        'RmDir (z.Item(i))
    Next
End Sub

Private Sub INFO_Click()
If INFO.Caption = "INFO" Then
Form1.Width = 13185
Form1.WindowState = 0
INFO.Caption = "Close INFO"
    Else
Form1.Width = 10785
Form1.WindowState = 0
INFO.Caption = "INFO"
    End If
End Sub

Private Sub Bahasa_Indonesia_Click()
LINGUA.Caption = "BAHASA"
Command2.Caption = "Buat daftar"
Command4.Caption = "M3u jelas dan txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Dansk_Click()
LINGUA.Caption = "SPROG"
Command2.Caption = "Generere lister"
Command4.Caption = "Ryd m3u og txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Deutsch_Click()
LINGUA.Caption = "SPRACHE"
Command2.Caption = "Listen erzeugen"
Command4.Caption = "Frei m3u und txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub English_Click()
LINGUA.Caption = "LANGUAGE"
Command2.Caption = "Generate lists"
Command4.Caption = "Clear m3u and txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Español_Click()
LINGUA.Caption = "IDIOMA"
Command2.Caption = "Generar listas"
Command4.Caption = "Cancelar m3u y txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Français_Click()
LINGUA.Caption = "LANGUE"
Command2.Caption = "Générer des listes"
Command4.Caption = "Clair et m3u txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Italiano_Click()
LINGUA.Caption = "LINGUA"
Command2.Caption = "Genera liste"
Command4.Caption = "Cancella m3u e txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Nederlands_Click()
LINGUA.Caption = "TAAL"
Command2.Caption = "Genereer lijsten"
Command4.Caption = "Duidelijke m3u en txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Polski_Click()
LINGUA.Caption = "JEZYK"
Command2.Caption = "Generowanie list"
Command4.Caption = "Wyczysc m3u i txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Português_Click()
LINGUA.Caption = "LINGUAGEM"
Command2.Caption = "Gerar listas"
Command4.Caption = "Limpar m3u e txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Suomi_Click()
LINGUA.Caption = "KIELI"
Command2.Caption = "Luo luettelot"
Command4.Caption = "Vapaa m3u ja txt"
If List1.ListCount = 0 Then
Label1.Caption = ""
Else
CARTROVATE
End If
End Sub

Private Sub Label2_Click()
Dim ApriPaginaWeb As Long
Form1.Width = 10785
Form1.WindowState = 0
INFO.Caption = "INFO"
    ApriPaginaWeb = ShellExecute(Me.hWnd, vbNullString, Label2.Caption, vbNullString, "c:\", 1)

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HFF8080
Label2.ForeColor = &HFFFFFF
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HFF8080
End Sub

Private Sub List1_Click()
File1.path = List1.Text
If Combo1.Text = "Audio" Then
File1.Pattern = "*.mp3;*.wma;*mid;*.mp1;*.mp2;*.mod;*.ogg;*.m4a;*.mp4;*.flac"
End If
If Combo1.Text = "Video" Then
File1.Pattern = "*.avi;*.asf;*.mpe;*.mpeg;*.mpg;*.mov;*.wmv"
End If
If Combo1.Text = "Nds - Gba" Then
File1.Pattern = "*.nds;*.gba;*.zip;*.7z;*.rar"
End If
If Combo1.Text = "Only Folders" Then
File1.Pattern = "*.*"
End If
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.ToolTipText = "mp3, wma, mid, mp1, mp2, mod, ogg, m4a, avi, asf, mpe, mpeg ,mpg ,mov ,wmv, mp4, flac, nds, gba, zip, 7z"
End Sub

Private Function ClipExt(ByVal fname As String) As String
Dim nPos As String
Dim nPos0 As String
Dim nPos1 As String
Dim nPos2 As String
Dim nPos3 As String
Dim nPos4 As String
Dim nPos5 As String
Dim nPos6 As String
Dim nPos7 As String
Dim nPos8 As String
Dim nPos9 As String
Dim nPos10 As String
Dim nPos11 As String
Dim nPos12 As String
Dim nPos13 As String
Dim nPos14 As String
Dim nPos15 As String
Dim nPos16 As String
Dim nPos17 As String
Dim nPos18 As String
Dim nPos19 As String
Dim nPos20 As String
Dim nPos21 As String
Dim nPos22 As String
Dim nPos23 As String
Dim nPos24 As String
Dim nPos25 As String
Dim nPos26 As String
Dim nPos27 As String
Dim nPos28 As String
Dim nPos29 As String
Dim nPos30 As String
Dim nPos31 As String
Dim nPos32 As String
Dim nPos33 As String
Dim nPos34 As String
Dim nPos35 As String
Dim nPos36 As String
Dim nPos37 As String
Dim nPos38 As String
Dim nPos39 As String
Dim nPos40 As String
Dim nPos41 As String
Dim nPos42 As String
Dim nPos43 As String
nPos = InStrRev(fname, ".mp3")
nPos0 = InStrRev(fname, ".Mp3")
nPos1 = InStrRev(fname, ".MP3")
nPos2 = InStrRev(fname, ".wma")
nPos3 = InStrRev(fname, ".WMA")
nPos4 = InStrRev(fname, ".mid")
nPos5 = InStrRev(fname, ".MID")
nPos6 = InStrRev(fname, ".mp1")
nPos7 = InStrRev(fname, ".MP1")
nPos8 = InStrRev(fname, ".mp2")
nPos9 = InStrRev(fname, ".MP2")
nPos10 = InStrRev(fname, ".mod")
nPos11 = InStrRev(fname, ".MOD")
nPos12 = InStrRev(fname, ".ogg")
nPos13 = InStrRev(fname, ".OGG")
nPos14 = InStrRev(fname, ".m4a")
nPos15 = InStrRev(fname, ".M4A")
nPos16 = InStrRev(fname, ".avi")
nPos17 = InStrRev(fname, ".AVI")
nPos18 = InStrRev(fname, ".asf")
nPos19 = InStrRev(fname, ".ASF")
nPos20 = InStrRev(fname, ".mpe")
nPos21 = InStrRev(fname, ".MPE")
nPos22 = InStrRev(fname, ".mpeg")
nPos23 = InStrRev(fname, ".MPEG")
nPos24 = InStrRev(fname, ".mpg")
nPos25 = InStrRev(fname, ".MPG")
nPos26 = InStrRev(fname, ".mov")
nPos27 = InStrRev(fname, ".MOV")
nPos28 = InStrRev(fname, ".wmv")
nPos29 = InStrRev(fname, ".WMV")
nPos30 = InStrRev(fname, ".mp4")
nPos31 = InStrRev(fname, ".MP4")
nPos32 = InStrRev(fname, ".nds")
nPos33 = InStrRev(fname, ".NDS")
nPos34 = InStrRev(fname, ".gba")
nPos35 = InStrRev(fname, ".GBA")
nPos36 = InStrRev(fname, ".zip")
nPos37 = InStrRev(fname, ".ZIP")
nPos38 = InStrRev(fname, ".7z")
nPos39 = InStrRev(fname, ".7Z")
nPos40 = InStrRev(fname, ".rar")
nPos41 = InStrRev(fname, ".RAR")
nPos42 = InStrRev(fname, ".flac")
nPos43 = InStrRev(fname, ".FLAC")
If nPos > 0 Then
fname = Mid(fname, 1, nPos - 1)
End If
If nPos0 > 0 Then
fname = Mid(fname, 1, nPos0 - 1)
End If
If nPos1 > 0 Then
fname = Mid(fname, 1, nPos1 - 1)
End If
If nPos2 > 0 Then
fname = Mid(fname, 1, nPos2 - 1)
End If
If nPos3 > 0 Then
fname = Mid(fname, 1, nPos3 - 1)
End If
If nPos4 > 0 Then
fname = Mid(fname, 1, nPos4 - 1)
End If
If nPos5 > 0 Then
fname = Mid(fname, 1, nPos5 - 1)
End If
If nPos6 > 0 Then
fname = Mid(fname, 1, nPos6 - 1)
End If
If nPos7 > 0 Then
fname = Mid(fname, 1, nPos7 - 1)
End If
If nPos8 > 0 Then
fname = Mid(fname, 1, nPos8 - 1)
End If
If nPos9 > 0 Then
fname = Mid(fname, 1, nPos9 - 1)
End If
If nPos10 > 0 Then
fname = Mid(fname, 1, nPos10 - 1)
End If
If nPos11 > 0 Then
fname = Mid(fname, 1, nPos11 - 1)
End If
If nPos12 > 0 Then
fname = Mid(fname, 1, nPos12 - 1)
End If
If nPos13 > 0 Then
fname = Mid(fname, 1, nPos13 - 1)
End If
If nPos14 > 0 Then
fname = Mid(fname, 1, nPos14 - 1)
End If
If nPos15 > 0 Then
fname = Mid(fname, 1, nPos15 - 1)
End If
If nPos16 > 0 Then
fname = Mid(fname, 1, nPos16 - 1)
End If
If nPos17 > 0 Then
fname = Mid(fname, 1, nPos17 - 1)
End If
If nPos18 > 0 Then
fname = Mid(fname, 1, nPos18 - 1)
End If
If nPos19 > 0 Then
fname = Mid(fname, 1, nPos19 - 1)
End If
If nPos20 > 0 Then
fname = Mid(fname, 1, nPos20 - 1)
End If
If nPos21 > 0 Then
fname = Mid(fname, 1, nPos21 - 1)
End If
If nPos22 > 0 Then
fname = Mid(fname, 1, nPos22 - 1)
End If
If nPos23 > 0 Then
fname = Mid(fname, 1, nPos23 - 1)
End If
If nPos24 > 0 Then
fname = Mid(fname, 1, nPos24 - 1)
End If
If nPos25 > 0 Then
fname = Mid(fname, 1, nPos25 - 1)
End If
If nPos26 > 0 Then
fname = Mid(fname, 1, nPos26 - 1)
End If
If nPos27 > 0 Then
fname = Mid(fname, 1, nPos27 - 1)
End If
If nPos28 > 0 Then
fname = Mid(fname, 1, nPos28 - 1)
End If
If nPos29 > 0 Then
fname = Mid(fname, 1, nPos29 - 1)
End If
If nPos30 > 0 Then
fname = Mid(fname, 1, nPos30 - 1)
End If
If nPos31 > 0 Then
fname = Mid(fname, 1, nPos31 - 1)
End If
If nPos32 > 0 Then
fname = Mid(fname, 1, nPos32 - 1)
End If
If nPos33 > 0 Then
fname = Mid(fname, 1, nPos33 - 1)
End If
If nPos34 > 0 Then
fname = Mid(fname, 1, nPos34 - 1)
End If
If nPos35 > 0 Then
fname = Mid(fname, 1, nPos35 - 1)
End If
If nPos36 > 0 Then
fname = Mid(fname, 1, nPos36 - 1)
End If
If nPos37 > 0 Then
fname = Mid(fname, 1, nPos37 - 1)
End If
If nPos38 > 0 Then
fname = Mid(fname, 1, nPos38 - 1)
End If
If nPos39 > 0 Then
fname = Mid(fname, 1, nPos39 - 1)
End If
If nPos40 > 0 Then
fname = Mid(fname, 1, nPos40 - 1)
End If
If nPos41 > 0 Then
fname = Mid(fname, 1, nPos41 - 1)
End If
If nPos42 > 0 Then
fname = Mid(fname, 1, nPos42 - 1)
End If
If nPos43 > 0 Then
fname = Mid(fname, 1, nPos43 - 1)
End If
ClipExt = fname
End Function

Public Sub TextFiles()
    Dim intSourceNum As Integer
    Dim intDestNum As Integer
    Dim strfile As String
    Dim aaa As String
    On Error Resume Next
    Kill Dir1.path & ".txt"
    strfile = Dir(Dir1.path & "\*.txt") 'adapt
    intDestNum = FreeFile()
    Open Dir1.path & ".txt" For Append As intDestNum 'adapt
    Do Until strfile = ""
        intSourceNum = FreeFile()
        Open Dir1.path & "\" & strfile For Input As intSourceNum
        Do While Not EOF(intSourceNum)
            Line Input #intSourceNum, strTMP
            If Trim(strTMP) <> " " Then
                Print #intDestNum, strTMP
            End If
        Loop
        Print #intDestNum,
        Close #intSourceNum
        strfile = Dir
    Loop
    Close #intDestNum
    EventPause 0.1
    Kill Dir1.path & "\*.txt"
End Sub

Public Sub EventPause(sngSeconds As Single)
    Dim dblTotal As Double, dblDateCounter As Double, sngStart As Single
    Dim dblReset As Double, sngTotalSecs As Single, intTemp As Integer
        dblDateCounter = ((Year(Date) + Month(Date) + Day(Date)) _
          & 0 & 0 & 0 & 0 & 0)
        sngStart = Timer
        sngTotalSecs = (sngStart + sngSeconds)
        intTemp = (sngTotalSecs \ 86400)
        dblReset = (intTemp * 100000) + (sngTotalSecs - (intTemp * 86400))
        dblTotal = dblDateCounter + dblReset
    Do
        DoEvents
    Loop While (dblDateCounter + Timer) < dblTotal
End Sub
