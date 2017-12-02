Imports System.Data.SQLite

Module PublicDB


    Public Sub UpgradeDatabase()

        Try

            Dim objConn As SQLiteConnection
            Dim objCommand As SQLiteCommand

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)

                objConn.Open()

                ' -----------------------------------------------------------------------------------

                If DB_VERSION < 1 Then

                    objCommand = objConn.CreateCommand()
                    objCommand.CommandText = "CREATE TABLE macro (MacroID integer primary key, MacroIsFav boolean, MacroName varchar(50), MacroNote varchar(200));"
                    objCommand.ExecuteNonQuery()

                    objCommand = objConn.CreateCommand()
                    objCommand.CommandText = "CREATE TABLE macro_step (StepID integer primary key, MacroID integer REFERENCES macro(MacroID) ON DELETE CASCADE, AccountID integer, CharID integer, StepOrder integer, StepWait integer, StepWindow varchar(10), StepX integer, StepY integer,  StepW integer, StepH integer );"
                    objCommand.ExecuteNonQuery()

                    objCommand = objConn.CreateCommand()
                    objCommand.CommandText = "ALTER TABLE weblink ADD COLUMN IsGuild boolean;"
                    objCommand.ExecuteNonQuery()


                    objCommand = objConn.CreateCommand()
                    objCommand.CommandText = "INSERT INTO param (ParamID, ParamName, ParamValue) VALUES " &
                                             " (10, 'DB_Version', '1') , " &
                                             " (11, 'Char_Image', 'Race') , " &
                                             " (12, 'ShowKill', '1') , " &
                                             " (13, 'ShowTaskBar', '1') , " &
                                             " (14, 'GuildLinks', '1') , " &
                                             " (15, 'GuildName', 'My Guild') , " &
                                             " (16, 'ShowSubWindowed', '1') , " &
                                             " (17, 'ShowSubSubWindowed', '0') ; "
                    objCommand.ExecuteNonQuery()


                    objCommand.CommandText = "DELETE FROM weblink  ; "
                    objCommand.ExecuteNonQuery()

                    objCommand.CommandText = "INSERT INTO weblink (LinkID, IsGuild, LinkOrder, LinkLabel, LinkURL, LinkImage) VALUES " &
                                         " (1, 0, 1, 'Nostalgeek', 'https://www.nostalgeek-serveur.com','13') , " &
                                         " (2, 0, 2, 'The Geek Crusade', 'https://www.thegeekcrusade-serveur.com/','13') , " &
                                         " (3, 0, 3, 'Elysium', 'https://elysium-project.org/','13') , " &
                                         " (4, 0, 4, 'WowHead', 'http://www.wowhead.com/','13') , " &
                                         " (5, 1, 5, 'J&D Forums', 'http://jambonetdragon.xooit.be/index.php','17') , " &
                                         " (6, 1, 6, 'J&D Agenda', 'https://www.allgenda.com/?agendaView=2&agendaType=1&aId=50893','4') , " &
                                         " (7, 1, 7, 'J&D Bank', 'http://www.jetd.fr/jetd/roster/guildbank2.php','21') , " &
                                         " (8, 1, 8, 'J&D RaidStats', 'http://realmplayers.com/RaidStats/RaidList.aspx?realm=NG&Guild=Jambon%20et%20Dragon','15') , " &
                                         " (9, 0, 9, 'WoW Classic Addon library', 'http://addons.us.to/home','13') , " &
                                         " (10, 0, 10, 'Vanilla-Addons', 'http://www.vanilla-addons.com/','13') , " &
                                         " (11, 0, 11, 'Nostalgeek Database', 'https://www.nostalgeek-serveur.com/db/','13') , " &
                                         " (12, 0, 12, 'Nostalgeek Talents', 'https://www.nostalgeek-serveur.com/template/','13') ; "

                    objCommand.ExecuteNonQuery()


                    DB_VERSION = 1
                    CHAR_IMAGE = "Race"
                    TRAY_SHOW_KILL = True
                    TRAY_SHOW_SUB_WINDOWED = True
                    TRAY_SHOW_SUBSUB_WINDOWED = True
                    SHOW_TASKBAR = False
                    GUILD_LINKS = True
                    GUILD_NAME = "My Guild"

                End If



                ' -----------------------------------------------------------------------------------

                If DB_VERSION < 2 Then

                    objCommand = objConn.CreateCommand()
                    objCommand.CommandText = "INSERT INTO param (ParamID, ParamName, ParamValue) VALUES " &
                                             " (18, 'Monitoring', '1') , " &
                                             " (19, 'SoftKill', '1') ; "
                    objCommand.ExecuteNonQuery()

                    SOFT_KILL = True
                    MONITORING = True
                    DB_VERSION = 2

                    Set_Param("DB_Version", 2)

                End If


            End Using


        Catch ex As Exception

            MessageBox.Show("Error: " & ex.Message)
        End Try


    End Sub


    Public Sub CreateDatabase()

        Dim objConn As SQLiteConnection
        Dim objCommand As SQLiteCommand

        Try
            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)

                objConn.Open()

                ' Structure creation

                objCommand = objConn.CreateCommand()
                objCommand.CommandText = "CREATE TABLE param (ParamID integer primary key, ParamName varchar(50), ParamValue varchar(200));"
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "CREATE TABLE wowversion (WoWversionID integer primary key, WoWversionName varchar(30), WoWversionActive boolean, WoWversionExe varchar(300));"
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "CREATE TABLE server (ServerID integer primary key, WoWversionID integer, ServerName varchar(50), ServerRealmList varchar(100));"
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "CREATE TABLE realm (RealmID integer primary key, ServerID integer REFERENCES server(ServerID) ON DELETE CASCADE, RealmName varchar(50), RealmType varchar(6), RealmLang char(2));"
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "CREATE TABLE account (AccountID integer primary key, RealmID integer REFERENCES realm(RealmID) ON DELETE CASCADE, AccountName varchar(50), AccountPassword varchar(50), AccountNote varchar(100), AccountIsFav boolean, AccountIsCustom boolean, AccountX integer, AccountY integer, AccountW integer, AccountH integer);"
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "CREATE TABLE character (CharID integer primary key, AccountID integer REFERENCES account(AccountID) ON DELETE CASCADE, CharIndex integer, CharIsFav boolean, CharName varchar(30),  CharFaction char(1), CharLevel integer, RaceID integer, ClassID integer, CharGender char(1), CharNote varchar(100));"
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "CREATE TABLE class (ClassID integer primary key, WoWversionID integer, ClassName varchar(30), ClassColor char(6));"
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "CREATE TABLE race (RaceID integer primary key, WoWversionID integer, RaceFaction char(1), RaceName varchar(30));"
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "CREATE TABLE weblink (LinkID integer primary key, LinkOrder integer, LinkLabel varchar(30), LinkURL varchar(500), LinkImage varchar(500));"
                objCommand.ExecuteNonQuery()

                ' Data creation


                objCommand.CommandText = "INSERT INTO server (ServerID, WoWversionID, ServerName, ServerRealmList) VALUES " &
                                        " (1,1,'Nostalgeek','auth.nostalgeek-serveur.com'),              " &
                                        " (2,2,'The Geek Crusade','auth.thegeekcrusade-serveur.com'),    " &
                                        " (3,1,'Light''s Hope','logon.lightshope.org'),                  " &
                                        " (4,1,'Elysium Project','logon.elysium-project.org'),           " &
                                        " (5,1,'Kronos','login.kronos-wow.com'),                         " &
                                        " (6,1,'Retro-WoW','logon.retro-wow.com'),                       " &
                                        " (7,1,'Classic-WoW','logon.classic-wow.org '),                  " &
                                        " (8,1,'Nostralia','login.nostralia.org');                       "
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "INSERT INTO realm (RealmID, ServerID, RealmName, RealmType, RealmLang) VALUES " &
                                        " (1,1,'Nostalgeek 1.12','PVP','FR'),      " &
                                        " (2,2,'The Geek Crusade BC','PVP','FR'),  " &
                                        " (3,3,'Elysium','PVP','EN'),              " &
                                        " (4,3,'Darrowshire','PVE','EN'),          " &
                                        " (5,3,'Anathema','PVP','EN'),             " &
                                        " (6,4,'Nighthaven','PVP','EN'),           " &
                                        " (7,4,'Stratholme','PVP','EN'),           " &
                                        " (8,5,'Kronos','PVP','EN'),               " &
                                        " (9,6,'RetroWoW','PVP','EN'),             " &
                                        " (10,7,'Nefarian','PVP','DE'),            " &
                                        " (11,8,'Nostralia','PVP','EN');           "
                objCommand.ExecuteNonQuery()



                objCommand.CommandText = "INSERT INTO param (ParamID, ParamName, ParamValue) VALUES " &
                                         " (1, 'VANILLA_EXE', '') , " &
                                         " (2, 'BC_EXE', '') , " &
                                         " (3, 'Wait_Load', '3000') , " &
                                         " (4, 'Wait_Login', '2000') , " &
                                         " (5, 'Wait_Down', '200') , " &
                                         " (6, 'IsCrypt', '0') , " &
                                         " (7, 'Default_Window', '0') , " &
                                         " (8, 'ShowFolder', '1') , " &
                                         " (9, 'ShowLinks', '1') ; "

                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "INSERT INTO wowversion (WoWversionID, WoWversionName, WoWVersionActive) VALUES " &
                                         "(1, 'Vanilla', 1)," &
                                         "(2, 'Burning Crusade', 1), " &
                                         "(3, 'Wrath of the Lich King', 0), " &
                                         "(4, 'Cataclysm', 0), " &
                                         "(5, 'Mists of Pandaria', 0), " &
                                         "(6, 'Warlords of Draenor', 0), " &
                                         "(7, 'Legion', 0); "
                objCommand.ExecuteNonQuery()



                objCommand.CommandText = "INSERT INTO race (RaceID, WoWversionID, RaceFaction, RaceName) VALUES " &
                                         "(1,1,'A','Human'), " &
                                         "(2,1,'A','Dwarf'), " &
                                         "(3,1,'A','Night Elf'), " &
                                         "(4,1,'A','Gnome'), " &
                                         "(5,2,'A','Dranei'), " &
                                         "(6,4,'A','Worgen'), " &
                                         "(7,1,'H','Orc'), " &
                                         "(8,1,'H','Undead'), " &
                                         "(9,1,'H','Tauren'), " &
                                         "(10,1,'H','Troll'), " &
                                         "(11,2,'H','Blood Elf'), " &
                                         "(12,4,'H','Goblin'), " &
                                         "(13,5,'N','Pandaren'); "
                objCommand.ExecuteNonQuery()

                objCommand.CommandText = "INSERT INTO class (ClassID, WoWversionID, ClassName, ClassColor) VALUES " &
                                         "(13, 1, 'Mule', 'FFD700'), " &
                                         "(1, 1, 'Druid', 'FF7D0A'), " &
                                         "(2, 1, 'Hunter', 'ABD473'), " &
                                         "(3, 1, 'Mage', '69CCF0'), " &
                                         "(4, 1, 'Paladin', 'F58CBA'), " &
                                         "(5, 1, 'Priest', 'FFFFFF'), " &
                                         "(6, 1, 'Rogue', 'FFF569'), " &
                                         "(7, 1, 'Shaman', '0070DE'), " &
                                         "(8, 1, 'Warlock','9482C9'), " &
                                         "(9, 1, 'Warrior', 'C79C6E'), " &
                                         "(10, 4, 'Death Knight', 'C41F3B'), " &
                                         "(11, 5, 'Monk','00FF96'), " &
                                         "(12, 7, 'Demon Hunter', 'A330C9') "
                objCommand.ExecuteNonQuery()



            End Using

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Error " & Err.Number)
        End Try

    End Sub



    ' Regrouping all DELETE here in case referential integrity doest work

    Public Sub DB_Delete_Char(ID As Long)


        Dim objConn As SQLiteConnection
        Dim objCommand As SQLiteCommand

        Try

            If ID = 0 Then Exit Sub

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)
                objConn.Open()
                objCommand = objConn.CreateCommand()
                objCommand.CommandText = "DELETE FROM character WHERE CharID=$ID"
                objCommand.Parameters.Add("$ID", DbType.Int32).Value = ID
                objCommand.ExecuteNonQuery()

            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        End Try


    End Sub


    Public Sub DB_Delete_Account(ID As Long)


        Dim objConn As SQLiteConnection
        Dim objCommand As SQLiteCommand

        Try

            If ID = 0 Then Exit Sub

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)
                objConn.Open()
                objCommand = objConn.CreateCommand()

                objCommand.CommandText = "DELETE FROM Account WHERE AccountID=$AccountID"
                objCommand.Parameters.Add("$AccountID", DbType.Int32).Value = ID
                objCommand.ExecuteNonQuery()

            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        End Try


    End Sub


    Public Sub DB_Delete_Server(ID As Long)


        Dim objConn As SQLiteConnection
        Dim objCommand As SQLiteCommand

        Try

            If ID = 0 Then Exit Sub

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)
                objConn.Open()
                objCommand = objConn.CreateCommand()

                objCommand.CommandText = "DELETE FROM server WHERE ServerID=$ServerID"
                objCommand.Parameters.Add("$ServerID", DbType.Int32).Value = ID
                objCommand.ExecuteNonQuery()

            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        End Try


    End Sub


    Public Sub DB_Delete_Realm(ID As Long)


        Dim objConn As SQLiteConnection
        Dim objCommand As SQLiteCommand

        Try

            If ID = 0 Then Exit Sub

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)
                objConn.Open()
                objCommand = objConn.CreateCommand()

                objCommand.CommandText = "DELETE FROM realm WHERE RealmID=$RealmID"
                objCommand.Parameters.Add("$RealmID", DbType.Int32).Value = ID
                objCommand.ExecuteNonQuery()

            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        End Try


    End Sub

    Public Sub DB_Delete_Macro(ID As Long)


        Dim objConn As SQLiteConnection
        Dim objCommand As SQLiteCommand

        Try

            If ID = 0 Then Exit Sub

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)
                objConn.Open()
                objCommand = objConn.CreateCommand()

                objCommand.CommandText = "DELETE FROM macro WHERE MacroID=$ID"
                objCommand.Parameters.Add("$ID", DbType.Int32).Value = ID
                objCommand.ExecuteNonQuery()

            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        End Try


    End Sub

    Public Sub DB_Delete_Step(ID As Long)


        Dim objConn As SQLiteConnection
        Dim objCommand As SQLiteCommand

        Try

            If ID = 0 Then Exit Sub

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)
                objConn.Open()
                objCommand = objConn.CreateCommand()

                objCommand.CommandText = "DELETE FROM macro_step WHERE StepID=$ID"
                objCommand.Parameters.Add("$ID", DbType.Int32).Value = ID
                objCommand.ExecuteNonQuery()

            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        End Try


    End Sub



    Public Sub DB_Delete_Link(ID As Long)


        Dim objConn As SQLiteConnection
        Dim objCommand As SQLiteCommand

        Try

            If ID = 0 Then Exit Sub

            objConn = New SQLiteConnection(CONSTRING & "New=True;")

            Using (objConn)
                objConn.Open()
                objCommand = objConn.CreateCommand()

                objCommand.CommandText = "DELETE FROM weblink WHERE LinkID=$LinkID"
                objCommand.Parameters.Add("$LinkID", DbType.Int32).Value = ID
                objCommand.ExecuteNonQuery()

            End Using

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        End Try


    End Sub


End Module
