VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//    Title: FMenu - v1.0
'//       By: Kevin Slaughter
'// Released: 5/12/04




'Declares
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long

'This is a small expansion to the original intent of FMenu, to support small
'  13x13 bitmaps for icons. PLEASE be careful using this!! Icons add to the menu
'  creation time, and can eat massive amounts of memory if over-used. Larger BMP's
'  could be used, but it would require subclassing.. And that would defeat the
'  purpose of this being a fast 'n simple menu, which is rather unlikely to crash. Bring
'  subclassing into the picture, and you're asking for trouble.
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0
'--


'Constants
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_CHECKED = &H8&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_POPUP = &H10&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RETURNCMD = &H100&


'Types
Private Type FMENUITEM
    lIcon As Long           'This is the Icon's index in the icon array, not the actual icon handle.
    lID As Long
    lFlags As Long
    sCaption As String
    sCommand As String
    sMenuName As String
End Type
Private Type FMENUITEMICON
    lIcon As Long
    sFilename As String     'Allows us to cache, kinda. Re-use icons if they're requested more than 1x
End Type
Private Type FMENUITEMCMDSTR
    bGoteExecWinStyle As Boolean
    bGoteMBType As Boolean
    bGotlMBButtons As Boolean
    bGotsCmd As Boolean
    bGotsExec As Boolean
    bGotsMBCaption As Boolean
    bGotsMBText As Boolean
    
    eExecWinStyle As VbAppWinStyle
    eMBType As VbMsgBoxStyle
    lMBButtons As Long
    sCmd As String
    sExec As String
    sMBCaption As String
    sMBText As String
End Type


'Events
Public Event ItemClicked(sKey As String)


'Variables
Private m_tMenuItems() As FMENUITEM     'All items in the menu
Private m_lIcons() As FMENUITEMICON     'If you -must- have icons, atleast be tidy about it!
Private m_lPopupMenus() As Long         'Array of all created popup menus, so all can be destroyed.
Private m_lNextAvailableItemID As Long  'Duh?
Private m_lParentHWND As Long           'Window calling this class
Private m_lMainHMENU As Long            'Root menu HMENU
Private m_lNumIcons As Long             'Easier error checking





'Properties
Public Property Get OwnerHWND() As Long
    OwnerHWND = m_lParentHWND
End Property
Public Property Let OwnerHWND(ByVal lVal As Long)
    m_lParentHWND = lVal
End Property





'Sub routines
Private Sub Class_Initialize()
    'Init arrays
    ReDim Preserve m_tMenuItems(0 To 0)
    ReDim Preserve m_lIcons(0 To 0)
    ReDim Preserve m_lPopupMenus(0 To 0)
    
    m_lIcons(0).lIcon = -1              'Invalidate
    m_lMainHMENU = CreatePopupMenu()    'Create root menu
    m_lNextAvailableItemID = 1          'Don't start this at 0.. TrackPopupMenu()
                                        '  returns 0 if nothing was selected.
End Sub
Private Sub Class_Terminate()
    Call DestroyAllMenus    'Free memory!
End Sub





'Functions
Private Function CacheIcon(ByVal sFile As String)
    Dim a As Long
    sFile = UCase$(sFile)
    
    'Check if it has already been cached. If so, exit.
    For a = 0 To UBound(m_lIcons)
        If (m_lIcons(a).sFilename = UCase$(sFile)) Then Exit Function
    Next a
    
    'Load it into memory
    ReDim Preserve m_lIcons(0 To IIf(m_lIcons(0).lIcon = -1, 0, UBound(m_lIcons) + 1))
    With m_lIcons(UBound(m_lIcons))
        .lIcon = LoadImage(-1, sFile, IMAGE_BITMAP, 13, 13, LR_LOADFROMFILE)
        .sFilename = sFile
    End With
    
    m_lNumIcons = (m_lNumIcons + 1)
End Function

Private Function ReadINI(ByVal sFile As String, ByVal sSection As String, ByVal sKey As String) As String
    Dim sTmpStr As String, lRet As String
    
    sTmpStr = String$(1024, Chr$(0))    'Buffer
    lRet = GetPrivateProfileString(sSection, sKey, "", sTmpStr, 1024, sFile)    'Get entry
    sTmpStr = Mid$(sTmpStr, 1, lRet)    'Truncate trailing spaces
    
    'Remove tabs, as they only show on the menu as a | symbol
    sTmpStr = Replace$(sTmpStr, vbTab, "", 1, -1, vbTextCompare)
    
    'Check for comments
    lRet = InStr(1, sTmpStr, "//", vbTextCompare)
    If (lRet > 0) Then
        sTmpStr = Mid$(sTmpStr, 1, lRet - 1)
    End If
    
    ReadINI = sTmpStr
End Function

Public Function LoadMenus(ByVal sFile As String)
    Dim tTMI As FMENUITEM
    Dim a As Long, b As Long, c As Long, sTmpStr As String, lRet As Long
    Dim lNumSections As Long, lNumItems As Long, lNumIcons As Long
    
    'Get total # of menu sections in the file
    lNumSections = CLng(GetPrivateProfileInt("INDEX", "NumSections", 0, sFile))
    
    'Load all icons into cache. The cache function will weed out duplicate icons, also.
    lNumIcons = CLng(GetPrivateProfileInt("ICONS", "NumIcons", 0, sFile))
    For a = 1 To lNumIcons
        sTmpStr = ReadINI(sFile, "ICONS", "Icon" & CStr(a)) 'Get icon path
        If (sTmpStr <> "") Then Call CacheIcon(sTmpStr)
    Next a
    
    'Parse sections
    If (lNumSections > 0) Then
        'Loop thru each section, and load it's items/commands
        For a = 0 To (lNumSections - 1)
            'Get section name
            tTMI.sMenuName = UCase$(ReadINI(sFile, CStr(a), "MenuName"))
            
            'Early error checking. Don't load items if the name is invalid. There is no point in
            '  continuing if no other menu can inherit this menu, due to an invalid name.
            If (tTMI.sMenuName <> "") Then
                'Get number of items
                lNumItems = GetPrivateProfileInt(CStr(a), "NumItems", 0, sFile)
                
                'Load items
                If (lNumItems > 0) Then
                    For b = 1 To lNumItems
                        'Get item name
                        tTMI.sCaption = ReadINI(sFile, CStr(a), "Item" & CStr(b))
                        
                        'Get item command
                        tTMI.sCommand = ReadINI(sFile, CStr(a), "Item" & CStr(b) & "Command")
                        
                        'Get item state.
                        tTMI.lFlags = MF_STRING
                        sTmpStr = UCase$(ReadINI(sFile, CStr(a), "Item" & CStr(b) & "State"))
                        If (InStr(1, sTmpStr, "CHECKED", vbTextCompare) > 0) Then tTMI.lFlags = (tTMI.lFlags Or MF_CHECKED)
                        If (InStr(1, sTmpStr, "DISABLED", vbTextCompare) > 0) Then tTMI.lFlags = (tTMI.lFlags Or MF_DISABLED)
                        If (InStr(1, sTmpStr, "GRAYED", vbTextCompare) > 0) Then tTMI.lFlags = (tTMI.lFlags Or MF_GRAYED)
                        
                        'Set icon. Icon #0 is nothing. Items that do not specify an Icon #, will
                        '  default to 0.
                        tTMI.lIcon = GetPrivateProfileInt(CStr(a), "Item" & CStr(b) & "Icon", 0, sFile)
                        
                        'Avoid an obviously potential "Out of bounds" array error.
                        If (tTMI.lIcon > m_lNumIcons) Then
                            tTMI.lIcon = 0
                            Call MsgBox("The requested icon # for [" & tTMI.sMenuName & "]Item" & CStr(b) & " is larger than the amount of icons specified in [ICONS]. Defaulting to icon #0 (no icon). Please fix this error in your menu.", vbExclamation, "FMenu")
                        End If
                        
                        'If the caption is "-", then make the item a separator bar. Doing this
                        '  after the state flag checks will save us from numerous repetitions of
                        '  checking to make sure the caption is not "-"
                        If (tTMI.sCaption = "-") Then
                            tTMI.lFlags = MF_SEPARATOR
                            tTMI.sCaption = ""
                        End If
                        
                        'Error checking..
                        'If the text has a length of 0, Windows will not display it at all.
                        If (tTMI.sCaption = "") Then tTMI.sCaption = " "
                        
                        'If no command is specified, default to FM_NONE.
                        If (tTMI.sCommand = "") Then
                            tTMI.sCommand = "FM_NONE();"
                        End If
                        
                        'Add the new menu to the array
                        ReDim Preserve m_tMenuItems(0 To (UBound(m_tMenuItems) + 1))
                        m_tMenuItems(UBound(m_tMenuItems)) = tTMI  'Copy new data into permanent array
                        
                        'Clear out current item's data.. Not doing so can result in false parsing for
                        '  the next item.
                        With tTMI
                            .lFlags = 0&
                            .lIcon = 0&
                            .lID = 0&
                            .sCaption = ""
                            .sCommand = ""
                        End With
                    Next b
                End If
            Else
                'If section 0 (ROOT) is invalid, cancel the entire process. Everything else
                '  will fail if we don't stop it here.
                If (a = 0) Then
                    Call MsgBox("ROOT menu was not the first section encountered." & vbCrLf & _
                                "ROOT must -always- be the first section in the file." & vbCrLf & _
                                vbCrLf & _
                                "Menu creation terminated. Please fix this error in your menu.", vbCritical, "FMenu")
                    ReDim m_tMenuItems(0 To 0) As FMENUITEM     'This menu is invalid. Clear the array.
                    Exit Function
                End If
            End If
        Next a
    End If
End Function

Private Function GetIndexesOfAllChildren(ByVal sMenuName As String, ByRef lIndexes() As Long) As Long
    Dim a As Long
    ReDim lIndexes(0 To 0) As Long  'Clear array
    lIndexes(0) = -1    'Invalidate so we can accurately report how many items were found
    sMenuName = UCase$(sMenuName)   'Case-Insensitive search
    
    For a = 0 To UBound(m_tMenuItems)
        If (m_tMenuItems(a).sMenuName = sMenuName) Then
            ReDim Preserve lIndexes(0 To IIf(lIndexes(0) = -1, 0, (UBound(lIndexes) + 1)))
            lIndexes(UBound(lIndexes)) = a
        End If
    Next a
    
    'Return how many children were found. Will be -1 for an error or none
    GetIndexesOfAllChildren = IIf(lIndexes(0) > -1, UBound(lIndexes) + 1, -1)
End Function

Private Function BuildMenu(ByVal DestHMENU As Long, ByVal sMenuName As String)
    Dim a As Long, lItems() As Long, sTmp As String
    
    'By re-using the item ID's, this allows us to inherit a menu numerous times,
    '  but only have 1 of each of it's items in memory. :)
    Call GetIndexesOfAllChildren(sMenuName, lItems())
    If (lItems(0) > -1) Then
        For a = 0 To UBound(lItems)
            With m_tMenuItems(lItems(a))
                'Set ID and add item to menu.
                .lID = m_lNextAvailableItemID
                Call AppendMenu(DestHMENU, ByVal .lFlags, ByVal .lID, .sCaption)
                
                'Add bitmap icon, if requested.
                If (.lIcon >= 1) Then
                    Call SetMenuItemBitmaps(DestHMENU, ByVal .lID, MF_BYCOMMAND, ByVal m_lIcons(.lIcon - 1).lIcon, ByVal m_lIcons(.lIcon - 1).lIcon)
                End If
                
                'Note:: Since an item with a submenu technically cannot be clicked, the InheritMenu
                '         action will invalidate the menu item as a selectable item.
                If (UCase$(Mid$(.sCommand, 1, 14)) = "FM_INHERITMENU") Then
                    sTmp = UCase$(.sCommand)
                    
                    'Remove ending ");"
                    If (Right$(sTmp, 2) = ");") Then
                        sTmp = Mid$(sTmp, 1, Len(sTmp) - 2)
                    End If
                
                    'Create child menu..
                    Dim lUB As Long  'For internal reading clarity only.
                    lUB = (UBound(m_lPopupMenus) + 1)
                    ReDim Preserve m_lPopupMenus(0 To lUB)
                    m_lPopupMenus(lUB) = CreatePopupMenu()
                    
                    'Get child menu name
                    sTmp = Mid$(sTmp, 16)
                    
                    If (m_lPopupMenus(lUB) > 0) Then
                        Call ModifyMenu(m_lMainHMENU, .lID, .lFlags Or MF_POPUP, m_lPopupMenus(lUB), .sCaption)
                        .lID = m_lPopupMenus(lUB)
                        Call BuildMenu(m_lPopupMenus(lUB), sTmp)
                    End If
                    
                    .sCommand = "FM_NONE();"    'Avoid errors in the parser, if this item gets clicked for some un-Godly reason.
                End If
            End With
            
            m_lNextAvailableItemID = (m_lNextAvailableItemID + 1)   'Inc ID
        Next a
    Else
        'No children were found. This was most likely an invalid name entry.
        '  Since InheritMenu renders the item without a command.. Replace the
        '  command with an alert for our user. Setting the ID to 0, will make the
        '  command interpreter think the menu was just exited, even on the very
        '  unlikely event these items end up being able to be clicked.
        Call AppendMenu(DestHMENU, MF_STRING Or MF_DISABLED, 0, "** FMenu Error **")
        Call AppendMenu(DestHMENU, MF_SEPARATOR, 0, "-")
        Call AppendMenu(DestHMENU, MF_STRING Or MF_DISABLED, 0, "InheritMenu failed. The")
        Call AppendMenu(DestHMENU, MF_STRING Or MF_DISABLED, 0, "requested section cannot")
        Call AppendMenu(DestHMENU, MF_STRING Or MF_DISABLED, 0, "be located.")
    End If
End Function

Private Function FindItemByID(ByVal lID As Long) As Long
    Dim a As Long
    
    'Find an item based on it's ID, and return the first match.
    '  ID's -shouldn't- be duplicated, but it can happen with
    '  an extremely large menu. I would implement an ID checking
    '  function, to make sure they're not duplicated.. But it would
    '  be a big waste of processor usage. If this even happens at
    '  all, it wouldn't start until atleast over 100,000 items
    '  were in the menu, so I'm not really concerned with it.
    '
    'When you make a menu item have a sub-menu, the item's ID becomes
    '  the hMenu of the sub-menu. An hMenu is usually an extremely
    '  large number anyway, as it's really an address in memory
    '  (100,000 or higher, usually, sometimes even that in negative).
    
    For a = 0 To UBound(m_tMenuItems)
        With m_tMenuItems(a)
            If .lID = lID Then '
                FindItemByID = a
                Exit Function
            End If
        End With
    Next a
    
    FindItemByID = -1
End Function

Private Function DestroyAllMenus()
    Dim a As Long
    
    'Destroy all possible child menus. It doesn't matter if this function's
    '  calling to DestroyMenu() fails, as this is just a quick-'n-dirty way
    '  of trying to free up as much memory as we can. I didn't know if
    '  Windows would delete all of the sub-menu's attached to the main menu,
    '  when the main menu is destroyed.. So better safe than sorry. Memory
    '  leaks are -not- good!  ^_^
    
    'Destroy main menu first
    Call DestroyMenu(m_lMainHMENU)
    
    'Kill each sub-menu
    For a = 0 To UBound(m_lPopupMenus)
        If (m_lPopupMenus(a) > 0) Then Call DestroyMenu(m_lPopupMenus(a))
    Next a
    
    'Kill any icons in memory
    For a = 0 To UBound(m_tMenuItems)
        If (m_tMenuItems(a).lIcon > 0) Then Call DeleteObject(m_tMenuItems(a).lIcon)
    Next a
End Function

Private Function RunCommand(ByVal sCmd As String) As Boolean
On Error GoTo ErrHndlRC
    Dim sTmp() As String, sTmp2() As String, tII As FMENUITEMCMDSTR, a As Long
    
    'Remove ending ");" marks. 'C' style coding is used to have a standard appearance
    If (Right$(sCmd, 2) = ");") Then
        sCmd = Mid$(sCmd, 1, Len(sCmd) - 2)
    End If
    
    'Get command to execute internally (FM_??)
    a = InStr(1, sCmd, "(", vbTextCompare)
    If (a > 0) Then
        tII.sCmd = Mid$(sCmd, 1, (a - 1))
        tII.bGotsCmd = CBool((tII.sCmd <> "FM_NONE") And (tII.sCmd <> "FM_EXEC") And (tII.sCmd <> "FM_ALERT"))
        
        'Save parsing errors..
        If (tII.sCmd = "FM_NONE") Then
            tII.bGotsCmd = True
            tII.bGotsExec = True
        End If
        
        sCmd = Mid$(sCmd, (a + 1))  'Remove this part too.
    End If
    
    'Get all tokens for the command, and parse accordingly. This way, the command string
    '  doesn't have to be in any particular order (since some people are lazy with syntax.)
    sTmp = Split(sCmd, "|", -1, vbTextCompare)
    
    For a = 0 To UBound(sTmp)
        sTmp2 = Split(sTmp(a), "::", 2, vbTextCompare)  'Split at the :: mark
        
        If (UBound(sTmp2) = 1) Then 'Make sure there are 2 parts "PART1::PART2".. If not, we'll bitch.
            If ((sTmp2(0) <> "") And (sTmp2(1) <> "")) Then
                Select Case UCase$(sTmp2(0))
                    'MessageBox Buttons
                    Case "BUTTONS"
                        tII.bGotlMBButtons = True
                        Select Case UCase$(sTmp2(1))
                            Case "OK_CANCEL":   tII.lMBButtons = vbOKCancel
                            Case "OK_ONLY":     tII.lMBButtons = vbOKOnly
                            Case "YES_NO":      tII.lMBButtons = vbYesNo
                            Case Else:
                                tII.bGotlMBButtons = False
                                Call MsgBox("Invalid button type specified. Valid types are:" & vbCrLf & _
                                            vbCrLf & _
                                            "OK_CANCEL" & vbCrLf & _
                                            "OK_ONLY" & vbCrLf & _
                                            "YES_NO" & vbCrLf & _
                                            vbCrLf & _
                                            "Please correct this error in your menu. Defaulting to OK_CANCEL.", vbExclamation, "FMenu")
                        End Select
                    
                    'MessageBox Caption
                    Case "CAPTION"
                        tII.bGotsMBCaption = True
                        tII.sMBCaption = sTmp2(1)
                    
                    'App to launch
                    Case "EXEC"
                        tII.bGotsExec = True
                        tII.sExec = sTmp2(1)
                    
                    'Type of MessageBox to display
                    Case "TYPE"
                        tII.bGoteMBType = True
                        Select Case UCase(sTmp2(1))
                            Case "CRITICAL":    tII.eMBType = vbCritical
                            Case "EXCLAMATION": tII.eMBType = vbExclamation
                            Case "INFORMATION": tII.eMBType = vbInformation
                            Case "QUESTION":    tII.eMBType = vbQuestion
                            Case Else:
                                tII.bGoteMBType = False
                                Call MsgBox("Invalid alert type specified. Valid types are:" & vbCrLf & _
                                            vbCrLf & _
                                            "CRITICAL" & vbCrLf & _
                                            "EXCLAMATION" & vbCrLf & _
                                            "INFORMATION" & vbCrLf & _
                                            "QUESTION" & vbCrLf & _
                                            vbCrLf & _
                                            "Please correct this error in your menu. Defaulting to INFORMATION.", vbExclamation, "FMenu")
                        End Select
                    
                    'MessageBox Text
                    Case "TEXT"
                        tII.bGotsMBText = True
                        tII.sMBText = sTmp2(1)
                    
                    'Window style for a launched app
                    Case "WINSTYLE"
                        tII.bGoteExecWinStyle = True
                        Select Case UCase$(sTmp2(1))
                            Case "HIDE":            tII.eExecWinStyle = vbHide
                            Case "MAX":             tII.eExecWinStyle = vbMaximizedFocus
                            Case "MIN-FOCUS":       tII.eExecWinStyle = vbMinimizedFocus
                            Case "MIN-NOFOCUS":     tII.eExecWinStyle = vbMinimizedNoFocus
                            Case "NORMAL-FOCUS":    tII.eExecWinStyle = vbNormalFocus
                            Case "NORMAL-NOFOCUS":  tII.eExecWinStyle = vbNormalNoFocus
                            Case Else:
                                tII.bGoteExecWinStyle = False
                                Call MsgBox("Invalid window style type specified. Valid types are:" & vbCrLf & _
                                            vbCrLf & _
                                            "HIDE" & vbCrLf & _
                                            "MAX" & vbCrLf & _
                                            "MIN-FOCUS" & vbCrLf & _
                                            "MIN-NOFOCUS" & vbCrLf & _
                                            "NORMAL-FOCUS" & vbCrLf & _
                                            "NORMAL-NOFOCUS" & vbCrLf & _
                                             vbCrLf & _
                                             "Please correct this error in your menu. Defaulting to NORMAL-FOCUS.", vbExclamation, "FMenu")
                        End Select
                        
                    Case Else
                        Call MsgBox("Invalid tag '" & UCase$(sTmp2(0)) & "' was encountered in:" & vbCrLf & _
                                    vbCrLf & _
                                    sCmd & vbCrLf & _
                                    vbCrLf & _
                                    "This command may not execute properly. Please fix this error in your menu .", vbExclamation, "FMenu")
                End Select
            End If
        End If
    Next a

    'Fix with default properties, if a tag was left out.
    If (Not tII.bGoteExecWinStyle) Then tII.eExecWinStyle = vbNormalFocus
    If (Not tII.bGotlMBButtons) Then tII.lMBButtons = vbOKCancel
    If (Not tII.bGoteMBType) Then tII.eMBType = vbInformation
    If (Not tII.bGotsCmd) Then
        Call MsgBox("An invalid command was specified for this item, or a command was not specified at all. Please fix this error in your menu.", vbExclamation, "FMenu")
        tII.sCmd = "FM_NONE"
    End If
    If (Not tII.bGotsExec) Then
        Call MsgBox("An invalid execution command was specified for this item, or a execution command was not specified at all. Please fix this error in your menu.", vbExclamation, "FMenu")
        tII.sExec = ""
    End If
    If (Not tII.bGotsMBCaption) Then tII.sMBCaption = "FMenu (No caption specified)"
    If (Not tII.bGotsMBText) Then tII.sMBText = "FMenu (No text specified)"

    'FM_ALERT
    Select Case UCase$(tII.sCmd)
        Case "FM_ALERT"
            'Execute system command based on user's button press.
            '** The 'decision' commands to execute, will be treated as a system
            '     command to be passed to the Shell() function. Please keep this
            '     in mind, otherwise the command will fail (of course). If the
            '     user clicks 'No', or 'Cancel', the function will just exit.
            '
            '** The command resulting from a yes, MUST be a shell executable command.
            mbr = MsgBox(tII.sMBText, tII.lMBButtons Or tII.eMBType, tII.sMBCaption)
            Select Case mbr
                Case vbOK, vbYes
                    If (tII.sExec <> "") Then
                        RunCommand = IIf(Shell(tII.sExec, tII.eExecWinStyle) > 0, True, False)
                    Else
                        RunCommand = True   'No command. We already bitched earlier about it.
                    End If
                Case vbCancel, vbNo
                    Exit Function
            End Select
            
            'Return good news :)
            RunCommand = True
            Exit Function
    
        Case "FM_EXEC"
            If (tII.sExec <> "") Then
                RunCommand = CBool(Shell(tII.sExec, tII.eExecWinStyle) > 0)
            Else
                RunCommand = True   'No command. We already bitched earlier about it.
            End If
            Exit Function
       
        Case "FM_NONE"
            RunCommand = True
            Exit Function
        
        Case Else
            RunCommand = False
            Exit Function
    End Select
    Exit Function

ErrHndlRC:
    Call MsgBox("Error executing the command '" & tII.sExec & "'" & vbCrLf & _
                vbCrLf & _
                "Error #" & CStr(Err.Number) & ":" & vbCrLf & _
                Err.Description & vbCrLf & _
                vbCrLf & _
                "Please check the menu's file to resolve this issue.", vbExclamation, "FMenu Error")
    RunCommand = False
    Exit Function
End Function

Public Function ShowMenu(ByVal X As Long, ByVal Y As Long)
    On Error Resume Next    'If shell fails, ignore it
    Dim lRet As Long, sCmdSplit() As String
    
    'Create the menu, starting from root level. The function will recursively build
    '  the entire menu system from that point onward.
    Call BuildMenu(m_lMainHMENU, "ROOT")
    
    'Show the menu
    lRet = TrackPopupMenu(m_lMainHMENU, TPM_LEFTALIGN Or TPM_RETURNCMD, ByVal X, ByVal Y, 0, m_lParentHWND, ByVal 0&)
    lRet = FindItemByID(lRet)   'Var re-using again.
    If (lRet > 0) Then
        With m_tMenuItems(lRet)
            'Replace tokens in commands
            .sCommand = Replace(.sCommand, "$LF$", vbCrLf, 1, -1, vbTextCompare)
            .sCommand = Replace(.sCommand, "$DATE$", Date$, 1, -1, vbTextCompare)
            .sCommand = Replace(.sCommand, "$TIME$", Time$, 1, -1, vbTextCompare)
            
            Call RunCommand(.sCommand)
        End With
    Else
        'We do nothing at this point. An ID of 0 reports that either::
        '  A.) Menu was exited without clicking on anything
        '  B.) Somehow, a disabled item was clicked (All of their ID's are 0)
        Exit Function
    End If
End Function
