Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CEventChatReply) = GetAddress(AddressOf HandleEventChatReply)
    HandleDataSub(CEvent) = GetAddress(AddressOf HandleEvent)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
End Sub

Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long
        
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), index, buffer.ReadBytes(buffer.Length), 0, 0
End Sub

Private Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar index
                Else
                    ' send new char shit
                    If Not IsPlaying(index) Then
                        Call SendNewCharClasses(index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            Else
                Call AlertMsg(index, "Sorry, that account name is already taken!")
            End If
            
            Set buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(index, Name)

            If LenB(Trim$(Player(index).Name)) > 0 Then
                Call DeleteName(Player(index).Name)
            End If

            Call ClearPlayer(index)
            ' Everything went ok
            Call Kill(App.path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(index, "Your account has been deleted.")
            
            Set buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Check versions
            If buffer.ReadLong < CLIENT_MAJOR Or buffer.ReadLong < CLIENT_MINOR Or buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(index, Name)
            ClearBank index
            LoadBank index, Name
            
            ' Check if character data has been created
            If LenB(Trim$(Player(index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar index
            Else
                ' send new char shit
                If Not IsPlaying(index) Then
                    Call SendNewCharClasses(index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            
            Set buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        Name = buffer.ReadString
        Sex = buffer.ReadLong
        Class = buffer.ReadLong
        Sprite = buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(index) Then
            Call AlertMsg(index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(index, Name, Sex, Class, Sprite)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar index
        
        Set buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(index), index, Msg, QBColor(White))
    Call SendChatBubble(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, Msg, White)
    
    Set buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(index) & ": " & Msg
    Call SayMsg_Global(index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgTo = FindPlayer(buffer.ReadString)
    Msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(index) & " tells you, '" & Msg & "'", TellColor)
            Call PlayerMsg(index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Cannot message yourself.", BrightRed)
    End If
    
    Set buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = buffer.ReadLong 'CLng(Parse(1))
    movement = buffer.ReadLong 'CLng(Parse(2))
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(index).StunDuration > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(index).InShop > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(index) <> tmpX Then
        SendPlayerXY (index)
        Exit Sub
    End If

    If GetPlayerY(index) <> tmpY Then
        SendPlayerXY (index)
        Exit Sub
    End If

    Call PlayerMove(index, Dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerDir
    buffer.WriteLong index
    buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long
Dim buffer As clsBuffer
    
    ' get inventory slot number
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    invNum = buffer.ReadLong
    Set buffer = Nothing

    UseItem index, invNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long
    
    ' can't attack whilst casting
    If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> index Then
            TryPlayerAttackPlayer index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc index, i
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    CheckResource index, x, y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim buffer As clsBuffer
Dim sMes As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PointType = buffer.ReadByte 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(index, PointType) >= 255 Then
            PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.Strength
                Call SetPlayerStat(index, Stats.Strength, GetPlayerRawStat(index, Stats.Strength) + 1)
                sMes = "Strength"
            Case Stats.Endurance
                Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + 1)
                sMes = "Endurance"
            Case Stats.Intelligence
                Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + 1)
                sMes = "Intelligence"
            Case Stats.Agility
                Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + 1)
                sMes = "Agility"
            Case Stats.Willpower
                Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + 1)
                sMes = "Willpower"
        End Select
        
        SendActionMsg GetPlayerMap(index), "+1 " & sMes, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendPlayerData index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Name = buffer.ReadString 'Parse(1)
    Set buffer = Nothing
    i = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(index) & ".", BrightBlue)
            Call PlayerMsg(index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(index, n, GetPlayerX(index), GetPlayerY(index))
    Call PlayerMsg(index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    Call SetPlayerSprite(index, n)
    Call SendPlayerData(index)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim mapnum As Long
    Dim x As Long
    Dim y As Long, z As Long, w As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(index)
    i = Map(mapnum).Revision + 1
    Call ClearMap(mapnum)
    
    Map(mapnum).Name = buffer.ReadString
    Map(mapnum).Music = buffer.ReadString
    Map(mapnum).BGS = buffer.ReadString
    Map(mapnum).Revision = i
    Map(mapnum).Moral = buffer.ReadByte
    Map(mapnum).Up = buffer.ReadLong
    Map(mapnum).Down = buffer.ReadLong
    Map(mapnum).Left = buffer.ReadLong
    Map(mapnum).Right = buffer.ReadLong
    Map(mapnum).BootMap = buffer.ReadLong
    Map(mapnum).BootX = buffer.ReadByte
    Map(mapnum).BootY = buffer.ReadByte
    
    Map(mapnum).Weather = buffer.ReadLong
    Map(mapnum).WeatherIntensity = buffer.ReadLong
    
    Map(mapnum).Fog = buffer.ReadLong
    Map(mapnum).FogSpeed = buffer.ReadLong
    Map(mapnum).FogOpacity = buffer.ReadLong
    
    Map(mapnum).Red = buffer.ReadLong
    Map(mapnum).Green = buffer.ReadLong
    Map(mapnum).Blue = buffer.ReadLong
    Map(mapnum).Alpha = buffer.ReadLong
    
    Map(mapnum).MaxX = buffer.ReadByte
    Map(mapnum).MaxY = buffer.ReadByte
    ReDim Map(mapnum).Tile(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(mapnum).Tile(x, y).Layer(i).x = buffer.ReadLong
                Map(mapnum).Tile(x, y).Layer(i).y = buffer.ReadLong
                Map(mapnum).Tile(x, y).Layer(i).Tileset = buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map(mapnum).Tile(x, y).Autotile(z) = buffer.ReadLong
            Next
            Map(mapnum).Tile(x, y).Type = buffer.ReadByte
            Map(mapnum).Tile(x, y).Data1 = buffer.ReadLong
            Map(mapnum).Tile(x, y).Data2 = buffer.ReadLong
            Map(mapnum).Tile(x, y).Data3 = buffer.ReadLong
            Map(mapnum).Tile(x, y).Data4 = buffer.ReadString
            Map(mapnum).Tile(x, y).DirBlock = buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map(mapnum).Npc(x) = buffer.ReadLong
        Map(mapnum).NpcSpawnType(x) = buffer.ReadLong
        Call ClearMapNpc(x, mapnum)
    Next
    
    'Event Data!
    Map(mapnum).EventCount = buffer.ReadLong
        
    If Map(mapnum).EventCount > 0 Then
        ReDim Map(mapnum).Events(0 To Map(mapnum).EventCount)
        For i = 1 To Map(mapnum).EventCount
            With Map(mapnum).Events(i)
                .Name = buffer.ReadString
                .Global = buffer.ReadLong
                .x = buffer.ReadLong
                .y = buffer.ReadLong
                .PageCount = buffer.ReadLong
            End With
            If Map(mapnum).Events(i).PageCount > 0 Then
                ReDim Map(mapnum).Events(i).Pages(0 To Map(mapnum).Events(i).PageCount)
                For x = 1 To Map(mapnum).Events(i).PageCount
                    With Map(mapnum).Events(i).Pages(x)
                        .chkVariable = buffer.ReadLong
                        .VariableIndex = buffer.ReadLong
                        .VariableCondition = buffer.ReadLong
                        .VariableCompare = buffer.ReadLong
                            
                        .chkSwitch = buffer.ReadLong
                        .SwitchIndex = buffer.ReadLong
                        .SwitchCompare = buffer.ReadLong
                            
                        .chkHasItem = buffer.ReadLong
                        .HasItemIndex = buffer.ReadLong
                            
                        .chkSelfSwitch = buffer.ReadLong
                        .SelfSwitchIndex = buffer.ReadLong
                        .SelfSwitchCompare = buffer.ReadLong
                            
                        .GraphicType = buffer.ReadLong
                        .Graphic = buffer.ReadLong
                        .GraphicX = buffer.ReadLong
                        .GraphicY = buffer.ReadLong
                        .GraphicX2 = buffer.ReadLong
                        .GraphicY2 = buffer.ReadLong
                            
                        .MoveType = buffer.ReadLong
                        .MoveSpeed = buffer.ReadLong
                        .MoveFreq = buffer.ReadLong
                            
                        .MoveRouteCount = buffer.ReadLong
                        
                        .IgnoreMoveRoute = buffer.ReadLong
                        .RepeatMoveRoute = buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map(mapnum).Events(i).Pages(x).MoveRoute(0 To .MoveRouteCount)
                            For y = 1 To .MoveRouteCount
                                .MoveRoute(y).index = buffer.ReadLong
                                .MoveRoute(y).Data1 = buffer.ReadLong
                                .MoveRoute(y).Data2 = buffer.ReadLong
                                .MoveRoute(y).Data3 = buffer.ReadLong
                                .MoveRoute(y).Data4 = buffer.ReadLong
                                .MoveRoute(y).data5 = buffer.ReadLong
                                .MoveRoute(y).data6 = buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = buffer.ReadLong
                        .DirFix = buffer.ReadLong
                        .WalkThrough = buffer.ReadLong
                        .ShowName = buffer.ReadLong
                        .Trigger = buffer.ReadLong
                        .CommandListCount = buffer.ReadLong
                            
                        .Position = buffer.ReadLong
                    End With
                        
                    If Map(mapnum).Events(i).Pages(x).CommandListCount > 0 Then
                        ReDim Map(mapnum).Events(i).Pages(x).CommandList(0 To Map(mapnum).Events(i).Pages(x).CommandListCount)
                        For y = 1 To Map(mapnum).Events(i).Pages(x).CommandListCount
                            Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount = buffer.ReadLong
                            Map(mapnum).Events(i).Pages(x).CommandList(y).ParentList = buffer.ReadLong
                            If Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount > 0 Then
                                ReDim Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount)
                                For z = 1 To Map(mapnum).Events(i).Pages(x).CommandList(y).CommandCount
                                    With Map(mapnum).Events(i).Pages(x).CommandList(y).Commands(z)
                                        .index = buffer.ReadLong
                                        .Text1 = buffer.ReadString
                                        .Text2 = buffer.ReadString
                                        .Text3 = buffer.ReadString
                                        .Text4 = buffer.ReadString
                                        .Text5 = buffer.ReadString
                                        .Data1 = buffer.ReadLong
                                        .Data2 = buffer.ReadLong
                                        .Data3 = buffer.ReadLong
                                        .Data4 = buffer.ReadLong
                                        .data5 = buffer.ReadLong
                                        .data6 = buffer.ReadLong
                                        .ConditionalBranch.CommandList = buffer.ReadLong
                                        .ConditionalBranch.Condition = buffer.ReadLong
                                        .ConditionalBranch.Data1 = buffer.ReadLong
                                        .ConditionalBranch.Data2 = buffer.ReadLong
                                        .ConditionalBranch.Data3 = buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = buffer.ReadLong
                                        .MoveRouteCount = buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).index = buffer.ReadLong
                                                .MoveRoute(w).Data1 = buffer.ReadLong
                                                .MoveRoute(w).Data2 = buffer.ReadLong
                                                .MoveRoute(w).Data3 = buffer.ReadLong
                                                .MoveRoute(w).Data4 = buffer.ReadLong
                                                .MoveRoute(w).data5 = buffer.ReadLong
                                                .MoveRoute(w).data6 = buffer.ReadLong
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    
    'End Event Data

    Call SendMapNpcsToMap(mapnum)
    Call SpawnMapNpcs(mapnum)
    Call SpawnGlobalEvents(mapnum)
    
    For i = 1 To Player_HighIndex
        If Player(i).Map = mapnum Then
            SpawnMapEventsFor i, mapnum
        End If
    Next

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))
    ' Save the map
    Call SaveMap(mapnum)
    Call MapCache_Create(mapnum)
    Call ClearTempTile(mapnum)
    Call CacheResources(mapnum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            Call PlayerWarp(i, mapnum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i
    
    Call CacheMapBlocks(mapnum)

    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Get yes/no value
    s = buffer.ReadLong 'Parse(1)
    Set buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(index, GetPlayerMap(index))
    End If

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SpawnMapEventsFor(index, GetPlayerMap(index))
    Call SendJoinMap(index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next

    TempPlayer(index).GettingMap = NO
    Set buffer = New clsBuffer
    buffer.WriteLong SMapDone
    SendDataTo index, buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim amount As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    invNum = buffer.ReadLong 'CLng(Parse(1))
    amount = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing
    
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(index, invNum) < 1 Or GetPlayerInvItemNum(index, invNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(index, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, invNum)).Stackable > 0 Then
        If amount < 1 Or amount > GetPlayerInvItemValue(index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(index, invNum, amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(index))
    Next

    CacheResources GetPlayerMap(index)
    Call PlayerMsg(index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(index) & "!", White)
                Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim File As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    Kill filename
    Call PlayerMsg(index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call BanIndex(n, index)
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    SendMapEventData (index)

    Set buffer = New clsBuffer
    buffer.WriteLong SEditMap
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SItemEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SAnimationEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim npcNum As Long
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    npcNum = buffer.ReadLong

    ' Prevent hacking
    If npcNum < 0 Or npcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(Npc(npcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(Npc(npcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(npcNum)
    Call SaveNpc(npcNum)
    Call AddLog(GetPlayerName(index) & " saved Npc #" & npcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SResourceEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ResourceNum = buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SShopEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call AddLog(GetPlayerName(index) & " saving shop #" & shopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SSpellEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellnum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    spellnum = buffer.ReadLong

    ' Prevent hacking
    If spellnum < 0 Or spellnum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(spellnum)
    Call SaveSpell(spellnum)
    Call AddLog(GetPlayerName(index) & " saved Spell #" & spellnum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    ' The access
    i = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(index) Then
                Call PlayerMsg(index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Invalid access level.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(buffer.ReadString) 'Parse(1))
    SaveOptions
    Set buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadLong 'CLng(Parse(1))
    y = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        ' Change target
                        If TempPlayer(index).targetType = TARGET_TYPE_PLAYER And TempPlayer(index).target = i Then
                            TempPlayer(index).target = 0
                            TempPlayer(index).targetType = TARGET_TYPE_NONE
                            ' send target to player
                            SendTarget index
                        Else
                            TempPlayer(index).target = i
                            TempPlayer(index).targetType = TARGET_TYPE_PLAYER
                            ' send target to player
                            SendTarget index
                        End If
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(index)).Npc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(index)).Npc(i).x = x Then
                If MapNpc(GetPlayerMap(index)).Npc(i).y = y Then
                    If TempPlayer(index).target = i And TempPlayer(index).targetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(index).target = 0
                        TempPlayer(index).targetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget index
                    Else
                        ' Change target
                        TempPlayer(index).target = i
                        TempPlayer(index).targetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget index
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Spell slot
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    Set buffer = Nothing
    PlayerSwitchInvSlots index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        PlayerMsg index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(n) > GetTickCount Then
            PlayerMsg index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    Set buffer = Nothing
    PlayerSwitchSpellSlots index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSendPing
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleUnequip(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PlayerUnequipItem index, buffer.ReadLong
    Set buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData index
End Sub

Sub HandleRequestItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems index
End Sub

Sub HandleRequestAnimations(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations index
End Sub

Sub HandleRequestNPCS(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs index
End Sub

Sub HandleRequestResources(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources index
End Sub

Sub HandleRequestSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells index
End Sub

Sub HandleRequestShops(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops index
End Sub

Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' item
    tmpItem = buffer.ReadLong
    tmpAmount = buffer.ReadLong
        
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index)
    Set buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess(index) < 4 Then Exit Sub
    SetPlayerExp index, GetPlayerNextLevel(index)
    CheckPlayerLevelUp index
End Sub

Sub HandleForgetSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim spellslot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    spellslot = buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(index).spellBuffer.Spell = spellslot Then
        PlayerMsg index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(index).Spell(spellslot) = 0
    SendPlayerSpells index
    
    Set buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemamount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopslot = buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        itemamount = HasItem(index, .costitem)
        If itemamount = 0 Or itemamount < .costvalue Then
            PlayerMsg index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem index, .costitem, .costvalue
        GiveInvItem index, .Item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set buffer = Nothing
End Sub

Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim itemnum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(index, invSlot) < 1 Or GetPlayerInvItemNum(index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    itemnum = GetPlayerInvItemNum(index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(index).InShop).BuyRate / 100
    price = Item(itemnum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg index, "The shop doesn't want that item.", BrightRed
        ResetShopAction index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem index, itemnum, 1
    GiveInvItem index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    
    PlayerSwitchBankSlots index, oldSlot, newSlot
    
    Set buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim BankSlot As Long
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    BankSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    TakeBankItem index, BankSlot, amount
    
    Set buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    GiveBankItem index, invSlot, amount
    
    Set buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SaveBank index
    SavePlayer index
    
    TempPlayer(index).InBank = False
    
    Set buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim x As Long
    Dim y As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    x = buffer.ReadLong
    y = buffer.ReadLong
    
    If GetPlayerAccess(index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX index, x
        SetPlayerY index, y
        SendPlayerXYToMap index
    End If
    
    Set buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(index).target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = index Then
        PlayerMsg index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).y
    sX = Player(index).x
    sY = Player(index).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = index
    SendTradeRequest tradeTarget, index
End Sub

Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long

    If TempPlayer(index).InTrade > 0 Then
        TempPlayer(index).TradeRequest = 0
    Else
        tradeTarget = TempPlayer(index).TradeRequest
        ' let them know they're trading
        PlayerMsg index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
        PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has accepted your trade request.", BrightGreen
        ' clear the tradeRequest server-side
        TempPlayer(index).TradeRequest = 0
        TempPlayer(tradeTarget).TradeRequest = 0
        ' set that they're trading with each other
        TempPlayer(index).InTrade = tradeTarget
        TempPlayer(tradeTarget).InTrade = index
        ' clear out their trade offers
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next
        ' Used to init the trade window clientside
        SendTrade index, tradeTarget
        SendTrade tradeTarget, index
        ' Send the offer data - Used to clear their client
        SendTradeUpdate index, 0
        SendTradeUpdate index, 1
        SendTradeUpdate tradeTarget, 0
        SendTradeUpdate tradeTarget, 1
    End If
End Sub

Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(index).TradeRequest, GetPlayerName(index) & " has declined your trade request.", BrightRed
    PlayerMsg index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim itemnum As Long
    
    TempPlayer(index).AcceptTrade = True
    
    tradeTarget = TempPlayer(index).InTrade
        
    If tradeTarget > 0 Then
    
        ' if not both of them accept, then exit
        If Not TempPlayer(tradeTarget).AcceptTrade Then
            SendTradeStatus index, 2
            SendTradeStatus tradeTarget, 1
            Exit Sub
        End If
    
        ' take their items
        For i = 1 To MAX_INV
            ' player
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                itemnum = Player(index).Inv(TempPlayer(index).TradeOffer(i).Num).Num
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem(i).Num = itemnum
                    tmpTradeItem(i).Value = TempPlayer(index).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot index, TempPlayer(index).TradeOffer(i).Num, tmpTradeItem(i).Value
                End If
            End If
            ' target
            If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
                itemnum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem2(i).Num = itemnum
                    tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).Value
                End If
            End If
        Next
    
        ' taken all items. now they can't not get items because of no inventory space.
        For i = 1 To MAX_INV
            ' player
            If tmpTradeItem2(i).Num > 0 Then
                ' give away!
                GiveInvItem index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False
            End If
            ' target
            If tmpTradeItem(i).Num > 0 Then
                ' give away!
                GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False
            End If
        Next
    
        SendInventory index
        SendInventory tradeTarget
    
        ' they now have all the items. Clear out values + let them out of the trade.
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg index, "Trade completed.", BrightGreen
        PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
        SendCloseTrade index
        SendCloseTrade tradeTarget
            
    End If
End Sub

Sub HandleDeclineTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(index).InTrade
    
    If tradeTarget > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg index, "You declined the trade.", BrightRed
        PlayerMsg tradeTarget, GetPlayerName(index) & " has declined the trade.", BrightRed
    
        SendCloseTrade index
        SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleTradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    Dim EmptySlot As Long
    Dim itemnum As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    amount = buffer.ReadLong
    
    Set buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    itemnum = GetPlayerInvItemNum(index, invSlot)
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = invSlot Then
                ' add amount
                TempPlayer(index).TradeOffer(i).Value = TempPlayer(index).TradeOffer(i).Value + amount
                ' clamp to limits
                If TempPlayer(index).TradeOffer(i).Value > GetPlayerInvItemValue(index, invSlot) Then
                    TempPlayer(index).TradeOffer(i).Value = GetPlayerInvItemValue(index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = invSlot Then
                PlayerMsg index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(index).TradeOffer(EmptySlot).Num = invSlot
    TempPlayer(index).TradeOffer(EmptySlot).Value = amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    tradeSlot = buffer.ReadLong
    
    Set buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    sType = buffer.ReadLong
    Slot = buffer.ReadLong
    hotbarNum = buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(index).Hotbar(hotbarNum).Slot = 0
            Player(index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(index).Inv(Slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(index, Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Inv(Slot).Num
                        Player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(index).Spell(Slot) > 0 Then
                    If Len(Trim$(Spell(Player(index).Spell(Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Spell(Slot)
                        Player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar index
    
    Set buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Slot = buffer.ReadLong
    
    Select Case Player(index).Hotbar(Slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(index).Inv(i).Num > 0 Then
                    If Player(index).Inv(i).Num = Player(index).Hotbar(Slot).Slot Then
                        If Item(Player(index).Inv(i).Num).Type = ITEM_TYPE_CONSUME Then
                            Player(index).Hotbar(Slot).Slot = 0
                            Player(index).Hotbar(Slot).sType = 0
                            SendHotbar index
                        End If
                        UseItem index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(index).Spell(i) > 0 Then
                    If Player(index).Spell(i) = Player(index).Hotbar(Slot).Slot Then
                        BufferSpell index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(index).target = index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(index).target) Or Not IsPlaying(TempPlayer(index).target) Then Exit Sub
    If GetPlayerMap(TempPlayer(index).target) <> GetPlayerMap(index) Then Exit Sub
    
    ' init the request
    Party_Invite index, TempPlayer(index).target
End Sub

Sub HandleAcceptParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(index).partyInvite, index
End Sub

Sub HandleDeclineParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(index).partyInvite, index
End Sub

Sub HandlePartyLeave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave index
End Sub

Sub HandleEventChatReply(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim eventID As Long, pageID As Long, reply As Long, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    eventID = buffer.ReadLong
    pageID = buffer.ReadLong
    reply = buffer.ReadLong
    
    If TempPlayer(index).EventProcessingCount > 0 Then
        For i = 1 To TempPlayer(index).EventProcessingCount
            If TempPlayer(index).EventProcessing(i).eventID = eventID And TempPlayer(index).EventProcessing(i).pageID = pageID Then
                If TempPlayer(index).EventProcessing(i).WaitingForResponse = 1 Then
                    If reply = 0 Then
                        If Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).index = EventType.evShowText Then
                            TempPlayer(index).EventProcessing(i).WaitingForResponse = 0
                        End If
                    ElseIf reply > 0 Then
                        If Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).index = EventType.evShowChoices Then
                            Select Case reply
                                Case 1
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data1
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 2
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data2
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 3
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data3
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 4
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data4
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                            End Select
                        End If
                        TempPlayer(index).EventProcessing(i).WaitingForResponse = 0
                    End If
                End If
            End If
        Next
    End If
    
    
    
    Set buffer = Nothing
End Sub

Sub HandleEvent(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long, begineventprocessing As Boolean, z As Long, buffer As clsBuffer

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    i = buffer.ReadLong
    Set buffer = Nothing
    
    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
            If TempPlayer(index).EventMap.EventPages(z).eventID = i Then
                i = z
                begineventprocessing = True
                Exit For
            End If
        Next
    End If
    
    If begineventprocessing = True Then
        If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
            'Process this event, it is action button and everything checks out.
            TempPlayer(index).EventProcessingCount = TempPlayer(index).EventProcessingCount + 1
            ReDim Preserve TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
            With TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
                .ActionTimer = GetTickCount
                .CurList = 1
                .CurSlot = 1
                .eventID = TempPlayer(index).EventMap.EventPages(i).eventID
                .pageID = TempPlayer(index).EventMap.EventPages(i).pageID
                .WaitingForResponse = 0
                ReDim .ListLeftOff(0 To Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).eventID).Pages(TempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount)
            End With
        End If
        begineventprocessing = False
    End If
End Sub

Sub HandleRequestSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSwitchesAndVariables (index)
End Sub

Sub HandleSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set buffer = Nothing
    
    SendSwitchesAndVariables 0, True
End Sub
