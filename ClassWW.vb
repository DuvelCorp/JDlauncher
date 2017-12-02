Imports Microsoft.VisualBasic
Imports System
Imports System.Runtime.InteropServices


Public Class WoW_Session

    Private wSessionName As String

    Private wWoWversionID As Integer

    Private wServerID As Integer
    Private wServerName As String
    Private wServerAddress As String

    Private wRealmID As Integer
    Private wRealmName As String

    Private wAccountID As Integer
    Private wAccountName As String
    Private wAccountPassword As String

    Private wCharID As Integer
    Private wCharIndex As Integer
    Private wCharName As Integer
    Private wCharClass As String
    Private wCharRace As String
    Private wCharFaction As String
    Private wCharColor As String
    Private wCharLevel As Integer

    Private wScreen As Integer
    Private wIsWindowed As Boolean

    Private wProcID As Integer

    Private wHandle As IntPtr
    Private wX As Integer
    Private wY As Integer
    Private wW As Integer
    Private wH As Integer




    Public Property AccountID() As Integer
        Get
            Return wAccountID
        End Get

        Set(ByVal Value As Integer)


            Dim objConn = New SQLiteConnection(CONSTRING)
            Dim objCommand As SQLiteCommand
            Dim objReader As SQLiteDataReader

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)

                objConn.Open()

                sql = " SELECT A.AccountName, A.AccountPassword, " &
                  "        S.ServerID, S.ServerName, S.ServerRealmList,   " &
                  "        R.RealmName, R.RealmID,  " &
                  "        v.WoWversionID, v.WoWversionName     " &
                  " FROM account a INNER JOIN realm r ON A.RealmID=r.RealmID " &
                  " INNER JOIN server s on S.ServerID=r.ServerID " &
                  " INNER JOIN wowversion v on S.wowversionID=v.wowversionID " &
                  " WHERE a.AccountID = " & Value

            End Using

            objCommand = objConn.CreateCommand()
            objCommand.CommandText = sql
            objReader = objCommand.ExecuteReader()

            While (objReader.Read())
                wAccountID = Value
                wAccountName = objReader("AccountName")
                wAccountPassword = objReader("AccountPassword")
                wRealmID = CInt(objReader("RealmID"))
                wRealmName = objReader("RealmName")
                wServerID = CInt(objReader("ServerID"))
                wServerName = objReader("ServerName")
                wServerAddress = objReader("ServerRealmList")
                wWoWversionID = CInt(objReader("WowversionID"))

            End While

        End Set

    End Property

    Public Property CharID() As Integer
        Get
            Return wCharID
        End Get

        Set(ByVal Value As Integer)

            Dim objConn = New SQLiteConnection(CONSTRING)
            Dim objCommand As SQLiteCommand
            Dim objReader As SQLiteDataReader

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)

                objConn.Open()

                sql = " SELECT .ServerID, s.ServerName, s.ServerRealmList, s.WoWversionID,  " &
                      "        r.RealmID, r.RealmName, r.RealmType, r.RealmLang,  " &
                      "        a.AccountID, a.AccountName, a.AccountPassword, a.AccountNote, " &
                      "        ra.RaceName, ra.RaceFaction, cl.ClassName, cl.ClassColor,  " &
                      "        c.CharID, c.CharName, c.CharIndex,  c.CharNote, C.CharLevel, C.CharGender, C.CharFaction  " &
                      " FROM character c " &
                      " INNER JOIN account a ON a.AccountID=c.AccountID " &
                      " LEFT JOIN class cl ON c.CLassID=cl.ClassID " &
                      " LEFT JOIN race ra ON c.RaceID=ra.RaceID " &
                      "  INNER JOIN realm r ON r.RealmID=a.RealmID " &
                      "  INNER JOIN server s ON s.ServerID=r.ServerID " &
                      " WHERE c.CharID=" & Value

                objCommand = objConn.CreateCommand()
                objCommand.CommandText = sql
                objReader = objCommand.ExecuteReader()

                While (objReader.Read())
                    wCharID = Value
                    wCharName = objReader("CharName")
                    wCharIndex = objReader("CharIndex")
                    wCharLevel = objReader("CharLevel")
                    wCharColor = objReader("CharColor")
                    wCharRace = objReader("CharRace")
                    wCharClass = objReader("CharClass")
                    wCharFaction = objReader("CharFaction")

                    wAccountID = objReader("AccountID")
                    wAccountName = objReader("AccountName")
                    wAccountPassword = objReader("AccountPassword")
                    wRealmID = CInt(objReader("RealmID"))
                    wRealmName = objReader("RealmName")
                    wServerID = CInt(objReader("ServerID"))
                    wServerName = objReader("ServerName")
                    wServerAddress = objReader("ServerRealmList")
                    wWoWversionID = CInt(objReader("WowversionID"))

                End While

            End Using




        End Set

    End Property


    Public Property Screen() As Integer
        Get
            Return wScreen
        End Get

        Set(ByVal Value As Integer)
            wScreen = Value
        End Set

    End Property



    Public Property IsWindowed() As Boolean
        Get
            Return wIsWindowed
        End Get

        Set(ByVal Value As Boolean)
            wIsWindowed = Value
        End Set

    End Property



    Public Property WindowX() As Integer
        Get
            Return wX
        End Get

        Set(ByVal Value As Integer)
            wX = Value
        End Set

    End Property


    Public Property WindowY() As Integer
        Get
            Return wY
        End Get

        Set(ByVal Value As Integer)
            wY = Value
        End Set

    End Property

    Public Property WindowW() As Integer
        Get
            Return wW
        End Get

        Set(ByVal Value As Integer)
            wW = Value
        End Set

    End Property

    Public Property WindowH() As Integer
        Get
            Return wH
        End Get

        Set(ByVal Value As Integer)
            wH = Value
        End Set

    End Property

    Public Property SessionName() As String
        Get
            Return wSessionName
        End Get

        ' read-only
        'Set(ByVal Value As String)
        '    wSessionName = Value
        'End Set

    End Property

    Public Property WoWversion() As Integer
        Get
            Return wWoWversion
        End Get

        ' read-only
        'Set(ByVal Value As Integer)
        '    wWoWversion = Value
        'End Set

    End Property

    Public Property ProcID() As Integer
        Get
            Return wprocid
        End Get

        ' read-only
        'Set(ByVal Value As Integer)
        '    wprocid = Value
        'End Set

    End Property





    Public Sub Start_WoW()

        Dim sql As String
        Dim Proc As process
        Dim i As Integer
        Dim strEXE As String

        Try

            Cursor.Current = Cursors.WaitCursor

            If Me.wAccountID = 0 Then Exit Sub

            If MainForm.WindowState = FormWindowState.Maximized Or MainForm.WindowState = FormWindowState.Normal Then
                MainForm.Timer1.Enabled = False
                MainForm.Overview_Init()

                If LAST_STARTED_CHAR = 0 Then
                    MainForm.Overview_Account(Me.wAccountID)
                Else
                    MainForm.Overview_Char(LAST_STARTED_CHAR)
                End If

                MainForm.LabelOverviewProcessStarted.Text = "Starting..."
                MainForm.LabelOverviewProcessStarted.ForeColor = Color.Orange
                MainForm.Refresh()
            End If


            If Me.wWoWversionID = 2 Then
                strEXE = BC_EXE
            Else
                strEXE = VANILLA_EXE
            End If


            If strEXE = "" Then
                MessageBox.Show("You have not located your WoW.exe ! Do that it the settings.")
                Exit Sub
            End If


            Me.Set_WoW_Config()
            Me.Set_WoW_RealmList()


            Me.wProcID = OpenApp(strEXE, Me.wIsWindowed, WAIT_LOAD, Me.wX, Me.wY, Me.wW.W, Me.wH)


            Me.wSessionName = Me.wWoWversionID & Me.wRealmName & " -> " & Me.wAccountName

            LAST_STARTED_ACCOUNT = Me.wAccountID

            COLL_SESSIONS.add(Me)



            If MainForm.WindowState = FormWindowState.Maximized Or MainForm.WindowState = FormWindowState.Normal Then
                MainForm.Check_Processes()
            End If



            Threading.Thread.Sleep(WAIT_LOAD)
            My.Computer.Keyboard.SendKeys(Me.wAccountName, True)
            My.Computer.Keyboard.SendKeys("{TAB}", True)

            If ISCRYPT Then
                My.Computer.Keyboard.SendKeys(PWD_Decrypt(Me.wAccountPassword), True)
            Else
                My.Computer.Keyboard.SendKeys(Me.wAccountPassword, True)
            End If

            My.Computer.Keyboard.SendKeys("{ENTER}", True)

            If Me.wCharIndex > 0 Then

                Threading.Thread.Sleep(WAIT_LOGIN)

                For i = 1 To Me.wCharIndex - 1
                    My.Computer.Keyboard.SendKeys("{DOWN}", True)
                    Threading.Thread.Sleep(WAIT_DOWN)
                Next i
                My.Computer.Keyboard.SendKeys("{ENTER}", True)

            End If


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        Finally
            Cursor.Current = Cursors.Arrow
            If MainForm.WindowState = FormWindowState.Maximized Or MainForm.WindowState = FormWindowState.Normal Then
                MainForm.Timer1.Enabled = True
            End If

        End Try

    End Sub



    Public Sub Set_WoW_Window(Optional Template As String = "")

        Dim res1 As Point
        Dim res2 As Point

        Try

            res1 = GetPrimaryScreenWorkingArea()
            res2 = GetSecondaryScreenWorkingArea()

            If Me.wScreen = 2 Then
                Me.wX = res1.X + 1
                Me.wY = 1
                Me.wH = res1.Y
            Else
                Me.wX = 1
                Me.wY = 1
                Me.wH = res1.Y
            End If


            If Me.wIsWindowed And Template <> "" Then

                Select Case Template

                    Case "TopLeft"

                        If Me.wScreen = 2 Then
                            Me.wX = res1.X + 1
                            Me.wY = 1
                            Me.wH = Math.Floor(res2.Y / 2)
                        Else
                            Me.wX = 1
                            Me.wY = 1
                            Me.wH = Math.Floor(res1.Y / 2)
                        End If

                    Case "TopRight"
                        Me.wH = res1.Y / 2

                        If Me.wScreen = 2 Then
                            Me.wX = res1.X + 1 + Math.Floor(res2.X / 2)
                            Me.wY = 1
                            Me.wH = Math.Floor(res2.Y / 2)
                        Else
                            Me.wX = 1 + Math.Floor(res1.X / 2)
                            Me.wY = 1
                            Me.wH = Math.Floor(res1.Y / 2)
                        End If

                    Case "BottomLeft"

                        If Me.wScreen = 2 Then
                            Me.wX = res1.X + 1
                            Me.wY = 1 + Math.Floor(res2.Y / 2)
                            Me.wH = Math.Floor(res2.Y / 2)
                        Else
                            Me.wX = 1
                            Me.wY = 1 + Math.Floor(res1.Y / 2)
                            Me.wH = Math.Floor(res1.Y / 2)
                        End If


                    Case "BottomRight"
                        If Me.wScreen = 2 Then
                            Me.wX = res1.X + 1 + Math.Floor(res2.X / 2)
                            Me.wY = 1 + Math.Floor(res2.Y / 2)
                            Me.wH = Math.Floor(res2.Y / 2)
                        Else
                            Me.wX = 1 + Math.Floor(res1.X / 2)
                            Me.wY = 1 + Math.Floor(res1.Y / 2)
                            Me.wH = Math.Floor(res1.Y / 2)
                        End If

                    Case Else


                End Select

                If Me.wScreen = 2 Then
                    Me.wW = Me.wH * SCREEN2_RATIO
                Else
                    Me.wW = Me.wH * SCREEN1_RATIO
                End If


            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        End Try


    End Sub


    Private Sub Set_WoW_Config()

        Dim LineIN As String
        Dim LineOUT As String
        Dim Reader As StreamReader
        Dim Writer As StreamWriter

        Dim nbLines As Integer
        Dim WrittenFile As System.IO.FileInfo

        Dim InitialFilePath As String
        Dim NewFilePath As String

        Dim IsMaximizeFOund As Boolean
        Dim IsWindowFound As Boolean
        Dim IsCharIndexFound As Boolean
        Dim IsRealmNameFound As Boolean
        Dim IsRealmListFound As Boolean

        Dim MaximizeVal As Integer

        Dim strGamePath As String

        Const DQ As String = Chr(34)

        Try

            If Me.wWoWversionID = 2 Then
                strGamePath = BC_DIR
            Else
                strGamePath = VANILLA_DIR
            End If

            LineOUT = ""

            InitialFilePath = strGamePath & "\WTF\Config.wtf"
            NewFilePath = "WoW_COnfig_Temp.wtf"

            WrittenFile = New System.IO.FileInfo(NewFilePath)

            If File.Exists(NewFilePath) Then
                WrittenFile.Attributes = IO.FileAttributes.Normal
            End If

            Reader = New StreamReader(InitialFilePath)
            Writer = New StreamWriter(NewFilePath, False)


            Do
                LineIN = Reader.ReadLine()

                LineOUT = LineIN

                If Strings.Left(LineIN, 13) = "SET realmList" Then
                    IsRealmListFound = True
                    LineOUT = "SET realmList " & DQ & Me.wServerAddress & DQ
                End If

                If Strings.Left(LineIN, 13) = "SET realmName" Then
                    IsRealmNameFound = True
                    LineOUT = "SET realmName " & DQ & Me.wRealmName & DQ
                End If

                If Strings.Left(LineIN, 14) = "SET gxMaximize" Then

                    IsMaximizeFOund = True
                    MaximizeVal = CInt(Strings.Mid(17, 1))
                    If Me.wIsWindowed Then
                        LineOUT = "SET gxMaximize ""0"""
                    Else
                        LineOUT = "SET gxMaximize ""1"""
                    End If
                End If

                If Strings.Left(LineIN, 12) = "SET gxWindow" Then
                    IsWindowFound = True
                    MaximizeVal = CInt(Strings.Mid(15, 1))
                    LineOUT = "SET gxWindow ""1"""
                End If


                If Me.wCharID Then
                    If Strings.Left(LineIN, 22) = "SET lastCharacterIndex" Then
                        IsCharIndexFound = True
                        LineOUT = "SET lastCharacterIndex ""0"""
                    End If
                End If


                Writer.Write(LineOUT & vbCrLf)

                nbLines = nbLines + 1

            Loop Until LineIN Is Nothing

            If Not IsRealmListFound Then
                LineOUT = "SET realmList " & DQ & Me.wServerAddress & DQ
                Writer.Write(LineOUT & vbCrLf)
            End If

            If Not IsRealmNameFound Then
                LineOUT = "SET realmName " & DQ & Me.wServerName & DQ
                Writer.Write(LineOUT & vbCrLf)
            End If


            If Not IsMaximizeFOund Then
                If Me.wIsWindowed Then
                    LineOUT = "SET gxMaximize ""0"""
                Else
                    LineOUT = "SET gxMaximize ""1"""
                End If
                Writer.Write(LineOUT & vbCrLf)
            End If

            If Not IsWindowFound Then
                '   If IsWindowed Then
                ' LineOUT = "SET gxWindow ""0"""
                '  Else
                LineOUT = "SET gxWindow ""1"""
                'End If
                Writer.Write(LineOUT & vbCrLf)
            End If

            If Me.wCharID And Not IsCharIndexFound Then
                LineOUT = "SET lastCharacterIndex ""0"""
                Writer.Write(LineOUT & vbCrLf)
            End If

            Reader.Close()

            Writer.Close()

            FileCopy(NewFilePath, InitialFilePath)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

    End Sub




    Private Sub Set_WoW_RealmList()

        Dim Writer As StreamWriter
        Dim WrittenFile As System.IO.FileInfo

        Dim InitialFilePath As String
        Dim NewFilePath As String

        Dim strGamePath As String

        Try

            If Me.wWoWversionID = 2 Then
                strGamePath = BC_DIR
            Else
                strGamePath = VANILLA_DIR
            End If

            InitialFilePath = strGamePath & "\realmlist.wtf"
            NewFilePath = "realmlist_Temp.wtf"

            WrittenFile = New System.IO.FileInfo(NewFilePath)

            If File.Exists(NewFilePath) Then
                WrittenFile.Attributes = IO.FileAttributes.Normal
            End If

            Writer = New StreamWriter(NewFilePath, False)

            Writer.Write("SET realmList " & Me.wRealmList & vbCrLf)

            Writer.Close()

            FileCopy(NewFilePath, InitialFilePath)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

    End Sub






End Class
