Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : Main module containing sub main
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Private Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
'------------------------------------------------------ ENDS


Public fMain As New cfMain
Public revealWidgetTimerCount As Integer

Public saturnWidget As cwSaturn
Public aboutWidget As cwAbout




'--------------------------------------------------------------------------------------------------------------
' BUILD:
'
'
' Credits : Standing on the shoulders of the following giants:
'
'           Olaf Schmidt for his Rich Client 5 Cairo wrapper.
'           Shuja Ali (codeguru.com) for his settings.ini code.
'           Registry reading code from ALLAPI.COM.
'           Rxbagain on codeguru for his Open File common dialog code without dependent OCX
'           Krool on the VBForums for his impressive common control replacements
'           si_the_geek for his special folder code
'           Elroy for the balloon tooltips
'           Rod Stephens vb-helper.com for ResizeControls
'
' Tools:    Built on a 3.3ghz Dell Latitude E6410 running Windows 7 Ultimate 64bit using VB6 SP6, VbAdvance, MZ-TOOLS 3.0,
'           CodeHelp Core IDE Extender Framework 2.2 & Rubberduck 2.4.1
'
'           MZ-TOOLS https://www.mztools.com/
'           CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1
'           Rubberduck http://rubberduckvba.com/
'           Registry code ALLAPI.COM
'           VbAdvance
'           La Volpe  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
'           Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain
'           Open font dialog code without dependent OCX - unknown URL
'           Krool's replacement Controls http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29
'
'   Tested on :
'           ReactOS 0.4.14 32bit on virtualBox
'           Windows 7 Professional 32bit on Intel - wip
'           Windows 7 Ultimate 64bit on Intel
'           Windows 7 Professional 64bit on Intel - done
'           Windows XP SP3 32bit on Intel - wip
'           Windows 10 Home 64bit on Intel - done
'           Windows 10 Home 64bit on AMD
'           Windows 11 64bit on Intel - done
'
' Dependencies:
'
'           Requires a Saturn folder in C:\Users\<user>\AppData\Roaming\ eg: C:\Users\<user>\AppData\Roaming\Saturn
'           Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\Saturn
'           The above will be created automatically by the compiled program when run for the first time.
'
'           Uses just one OCX control extracted from Krool's mega pack (slider). This is part of Krool's replacement for the
'           whole of Microsoft Windows Common Controls found in mscomctl.ocx. The slider control OCX file is shipped with
'           this package.
'
'           * CCRSlider.ocx
'
'           This OCX will reside in the program folder. The program reference to this OCX is contained within the
'           supplied resource file Saturn.RES. It is compiled into the binary.
'
'           * OLEGuids.tlb
'
'           This is a type library that defines types, object interfaces, and more specific API definitions
'           needed for COM interop / marshalling. It is only used at design time (IDE). This is a Krool-modified
'           version of the original .tlb from the vbaccelerator website. The .tlb is compiled into the executable.
'           For the compiled .exe this is NOT a dependency, only during design time.
'
'           From the command line, copy the tlb to a central location (system32 or wow64 folder) and register it.
'
'           COPY OLEGUIDS.TLB %SystemRoot%\System32\
'           REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB
'
'           In the VB6 IDE - project - references - browse - select the OLEGuids.tlb
'
'           * Uses the RC5 Cairo wrapper from Olaf Schmidt.
'
'           During development the RC5 components need to be registered. These scripts are used to register. Run each by
'           double-clicking on them.
'
'           RegisterRC5inPlace.vbs
'           RegisterVBWidgetsInPlace.vbs
'
'           During runtime on the user's system, the RC5 components are dynamically referenced using modRC5regfree.bas
'           which is compiled into the binary.
'
'           * SETUP.EXE - The program is currently distributed using setup2go, a very useful and comprehensive installer program
'           that builds a .exe installer. You'll have to find a copy of setup2go on the web as it is now abandonware. Contact me
'           directly for a copy. The file "install Saturn 0.1.0.s2g" is the configuration file for setup2go. When you build it will
'           report any errors in the build.
'
'           * HELP.CHM - the program documentation is built using the NVU HTML editor and compiled using the Microsoft supplied
'           CHM builder tools (HTMLHelp Workshop) and the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
'           the job admirably. The HTML files exist alongside the compiled CHM file in the HELP folder.
'
' Project References:
'
'           VisualBasic for Applications
'           VisualBasic Runtime Objects and Procedures
'           VisualBasic Objects and Procedures
'           OLE Automation
'           vbRichClient5
'
' Summary:
'
'           The program is quite simple but forms the structure for other similar programs yet to come. These
'           will be funcxtional replicas of the graphical Steam/Dieselpunk javascript widgets I previously built using the
'           Yahoo widget engine.
'
'           This program is a mix of native VB6 forms and controls and 3rd party additions.
'           The superb RC5 Cairo wrapper from Olaf Schmidt is used in a very limited manner.
'           RC5's transparency capability is used for the main saturn and the about image only. I haven't used
'           Olaf's other Cairo controls to build forms as I need a graphical IDE to operate. Only testing RC5
'           at the moment, there should be no problems upgrading to RC6.
'
'           The helpForm is a standard VB6 window without a titlebar nor controls, displaying a fullsize image.
'           Standard VB6 forms are used for the preference and licence windows.
'           The standard VB6 timers are located on an invisible standard VB6 form - frmTimer.
'           The standard VB6 menu is located on it's own invisible VB6 form - menuForm.
'           The frmTimer invisible form is also the container for the large 128x128px overall project icon.

'           The utility itself has some configuration details that it stores in the settings.ini file
'           within the user appdata roaming folder.
'
'           The font selection and file/folder dialogs are generated using Win32 APIs rather than the
'           common dialog OCX which dispensed with another MS OCX.
'
'           As stated above, I have used Krool's amazing control replacement project. The specific code for
'           just one of the controls (slider) has been incorporated rather than all 32 of
'           Krool's complete package.
'
'           The preference form is resizable - which allows it to run on high DPI systems. In my mind, this is a poor
'           man's method of handling high DPI, 4K screens. I find the creation of DPI aware VB6 programs with a working
'           side by side configuration using manifests difficult with VB6 on modern systems. Instead the controls are
'           resized dynamically when the form is dragged. The images are reloaded with higher res. versions on 1500 twip
'           intervals. I could have done with GDI+ using multiple embedded icons but it is OK for the moment. All in all,
'           it's a bit sh1t but it works well enough, so it'll do...
'
'           There is one useful .BAT file - unhide.bat which will reveal a Saturn saturn mistakenly 'hidden' for an
'           extended period of time from the right click menu. This will allow you to open the prefs and unset the hidden
'           configuration option.
'
'
'    LICENCE AGREEMENTS:
'
'    Copyright © 2023 Dean Beedell
'
'    Using this program implies you have accepted the licence. The GPL licence applies to the code
'    this software Is provided 'as-is', without any express or implied warranty. In no event will the
'    author be held liable for any damages arising from the use of this software. Permission is granted to
'    anyone to use this software for any purpose, including commercial applications, and to alter it and
'    redistribute it freely, subject to the following restrictions:
'
'    1. The origin of this software must not be misrepresented; you must not claim that you wrote the original software. If you use this software in a product, an acknowledgment in the product documentation is required.
'    2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original software.
'    3. This notice may not be removed or altered from any source distribution.
'
'    This program is free software; you can redistribute it and/or modify it under the terms of the
'    GNU General Public Licence as published by the Free Software Foundation; either version 2 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
'    even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'    General Public Licence for more details.
'
'    You should have received a copy of the GNU General Public Licence along with this program; if not,
'    write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
'    USA
'
'    If you use this software in any way whatsoever then that implies acceptance of the licence. If you
'    do not wish to comply with the licence terms then please remove the download, binary and source code
'    from your systems immediately.
'
'    FYI - I like CALLing subroutines, it may be old-fashioned but its what I do.
'
'--------------------------------------------------------------------------------------------------------------




'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : beededea
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Main()
    On Error GoTo Main_Error
    
    Call mainRoutine(False)

   On Error GoTo 0
   Exit Sub

Main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : main_routine
' Author    : beededea
' Date      : 27/06/2023
' Purpose   : called by sub main() to allow this routine to be called elsewhere,
'             a reload for example.
'---------------------------------------------------------------------------------------
'
Public Sub mainRoutine(ByVal restart As Boolean)
    Dim extractCommand As String: extractCommand = vbNullString

   On Error GoTo main_routine_Error

    extractCommand = Command$ ' capture any parameter passed
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    'add Resources to the global ImageList
    Call addImagesToImageList
    
    ' check the Windows version
    classicThemeCapable = fTestClassicThemeCapable
  
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the dock settings from the new configuration file
    Call readSettingsFile(softwarePlanet, StSettingsFile)
        
    ' check first usage and display licence screen
    Call checkLicenceState

    ' initialise and create the main forms on the current display
    Call createFormOnCurrentDisplay
    
    ' place the form at the saved location
    Call makeVisibleFormElements
    
    ' resolve VB6 sizing width bug
    Call determineScreenDimensions
    
    ' run the functions that are also called at reload time.
    Call adjustMainControls ' this needs to be here after the initialisation of the Cairo forms and widgets
    
    ' check the selected monitor properties to determine form placement
    'Call monitorProperties(frmHidden) - might use RC5 for this?
    
    ' move/hide onto/from the main screen
    Call mainScreen
        
    ' if the program is run in unhide mode, write the settings and exit
    Call handleUnhideMode(extractCommand)
    
    ' if a first time run shows prefs
    If PrFirstTimeRun = "true" Then     'parse the command line
        Call makeProgramPreferencesAvailable
    End If
    
    ' check for first time running
    Call checkFirstTime

    ' configure any global timers here
    Call configureTimers
    
    'load the preferences form but don't yet show it, speeds up access to the prefs via the menu
    Load saturnPrefs

    ' RC message pump will auto-exit when Cairo Forms = 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload.
    If Cairo.WidgetForms.Count = 0 Then Cairo.WidgetForms.EnterMessageLoop

   On Error GoTo 0
   Exit Sub

main_routine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main_routine of Module modMain"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkFirstTime
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : check for first time running
'---------------------------------------------------------------------------------------
'
Private Sub checkFirstTime()

   On Error GoTo checkFirstTime_Error

    If PrFirstTimeRun = "true" Then
        PrFirstTimeRun = "false"
        sPutINISetting softwarePlanet, "firstTimeRun", PrFirstTimeRun, StSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

checkFirstTime_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkFirstTime of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : initialiseGlobalVars
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : initialise global vars
'---------------------------------------------------------------------------------------
'
Private Sub initialiseGlobalVars()
      
    On Error GoTo initialiseGlobalVars_Error

    thisPlanet = "\Saturn"
    softwarePlanet = "Software" & thisPlanet

    ' general
    PrStartup = vbNullString
    PrGaugeFunctions = vbNullString
    'PrStampSelection = vbNullString
    
    ' config
    PrEnableTooltips = vbNullString
    PrEnableBalloonTooltips = vbNullString
    PrShowTaskbar = vbNullString
    
    PrGaugeSize = vbNullString
    PrScrollWheelDirection = vbNullString
    
    ' position
    PrAspectHidden = vbNullString
    PrWidgetPosition = vbNullString
    PrWidgetLandscape = vbNullString
    PrWidgetPortrait = vbNullString
    PrLandscapeFormHoffset = vbNullString
    PrLandscapeFormVoffset = vbNullString
    PrPortraitHoffset = vbNullString
    PrPortraitYoffset = vbNullString
    PrvLocationPercPrefValue = vbNullString
    PrhLocationPercPrefValue = vbNullString
    
    ' sounds
    PrEnableSounds = vbNullString
    
    ' development
    PrDebug = vbNullString
    PrDblClickCommand = vbNullString
    PrOpenFile = vbNullString
    PrDefaultEditor = vbNullString
         
    ' font
    PrPrefsFont = vbNullString
    PrPrefsFontSize = vbNullString
    PrPrefsFontItalics = vbNullString
    PrPrefsFontColour = vbNullString
    
    ' window
    PrWindowLevel = vbNullString
    PrPreventDragging = vbNullString
    PrOpacity = vbNullString
    PrWidgetHidden = vbNullString
    PrHidingTime = vbNullString
    PrIgnoreMouse = vbNullString
    PrFirstTimeRun = vbNullString
    
    ' general storage variables declared
    PrSettingsDir = vbNullString
    StSettingsFile = vbNullString
    PrMaximiseFormX = vbNullString
    PrMaximiseFormY = vbNullString
    PrLastSelectedTab = vbNullString
    PrSkinTheme = vbNullString
    
    ' general variables declared
    toolSettingsFile = vbNullString
    classicThemeCapable = False
    storeThemeColour = 0
    windowsVer = vbNullString
    
    ' vars to obtain correct screen width (to correct VB6 bug) STARTS
    screenTwipsPerPixelX = 0
    screenTwipsPerPixelY = 0
    screenWidthTwips = 0
    screenHeightTwips = 0
    screenHeightPixels = 0
    screenWidthPixels = 0
    oldScreenHeightPixels = 0
    oldScreenWidthPixels = 0
    
    ' key presses
    CTRL_1 = False
    SHIFT_1 = False
    
    ' other globals
    debugflg = 0
    minutesToHide = 0
    aspectRatio = vbNullString
    revealWidgetTimerCount = 0
    oldPrSettingsModificationTime = #1/1/2000 12:00:00 PM#

   On Error GoTo 0
   Exit Sub

initialiseGlobalVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGlobalVars of Module modMain"
    
End Sub

        
'---------------------------------------------------------------------------------------
' Procedure : addImagesToImageList
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : add Resources to the global ImageList
'---------------------------------------------------------------------------------------
'
Private Sub addImagesToImageList()
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo addImagesToImageList_Error

    Cairo.ImageList.AddImage "about", App.Path & "\Resources\images\about.png"
    
    'add Resources to the global ImageList
    Cairo.ImageList.AddImage "saturn", App.Path & "\Resources\images\saturn.png"

   On Error GoTo 0
   Exit Sub

addImagesToImageList_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToImageList of Module modMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the saturn and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustMainControls()
   
   On Error GoTo adjustMainControls_Error

    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' saturnWidget.RotationSpeed = Val(PrsaturnSelection)
    saturnWidget.Zoom = Val(PrGaugeSize) / 100
    saturnWidget.ZoomDirection = PrScrollWheelDirection
    
    If saturnWidget.Hidden = False Then
        saturnWidget.opacity = Val(PrOpacity) / 100
        saturnWidget.Widget.Refresh
    End If
    
    If PrGaugeFunctions = "1" Then
        'saturnWidget.Rotating = True
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        'saturnWidget.Rotating = False
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If PrDefaultEditor <> vbNullString And PrDebug = "1" Then
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & PrDefaultEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
        
    If PrPreventDragging = "0" Then
        menuForm.mnuLockWidget.Checked = False
        saturnWidget.Locked = False
    Else
        menuForm.mnuLockWidget.Checked = True
        saturnWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
    End If
    
    If PrShowTaskbar = "0" Then
        fMain.saturnForm.ShowInTaskbar = False
    Else
        fMain.saturnForm.ShowInTaskbar = True
    End If
                 
    ' set the z-ordering of the window
    Call setWindowZordering
    
    ' set the tooltips on the main screen
    Call setMainTooltips
    
    ' set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
    Call setHidingTime
    
    'Call setVisibleStampType

    If minutesToHide > 0 Then menuForm.mnuHideWidget.Caption = "Hide Widget for " & minutesToHide & " min."

   On Error GoTo 0
   Exit Sub

adjustMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustMainControls of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setWindowZordering
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setWindowZordering()

   On Error GoTo setWindowZordering_Error

    If Val(PrWindowLevel) = 0 Then
        Call SetWindowPos(fMain.saturnForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(PrWindowLevel) = 1 Then
        Call SetWindowPos(fMain.saturnForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(PrWindowLevel) = 2 Then
        Call SetWindowPos(fMain.saturnForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
    End If

   On Error GoTo 0
   Exit Sub

setWindowZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setWindowZordering of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal location As String, ByVal StSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(StSettingsFile) Then
        
        ' general
        PrStartup = fGetINISetting(location, "startup", StSettingsFile)
        PrGaugeFunctions = fGetINISetting(location, "gaugeFunctions", StSettingsFile)
        PrsaturnSelection = fGetINISetting(location, "saturnSelection", StSettingsFile)
        
        ' configuration
        PrEnableTooltips = fGetINISetting(location, "enableTooltips", StSettingsFile)
        PrEnableBalloonTooltips = fGetINISetting(location, "enableBalloonTooltips", StSettingsFile)
        PrShowTaskbar = fGetINISetting(location, "showTaskbar", StSettingsFile)
        
        PrGaugeSize = fGetINISetting(location, "gaugeSize", StSettingsFile)
        PrScrollWheelDirection = fGetINISetting(location, "scrollWheelDirection", StSettingsFile)
        'PrWidgetSkew = fGetINISetting(location, "widgetSkew", StSettingsFile)
        
        ' position
        PrAspectHidden = fGetINISetting(location, "aspectHidden", StSettingsFile)
        PrWidgetPosition = fGetINISetting(location, "widgetPosition", StSettingsFile)
        PrWidgetLandscape = fGetINISetting(location, "widgetLandscape", StSettingsFile)
        PrWidgetPortrait = fGetINISetting(location, "widgetPortrait", StSettingsFile)
        PrLandscapeFormHoffset = fGetINISetting(location, "landscapeHoffset", StSettingsFile)
        PrLandscapeFormVoffset = fGetINISetting(location, "landscapeYoffset", StSettingsFile)
        PrPortraitHoffset = fGetINISetting(location, "portraitHoffset", StSettingsFile)
        PrPortraitYoffset = fGetINISetting(location, "portraitYoffset", StSettingsFile)
        PrvLocationPercPrefValue = fGetINISetting(location, "vLocationPercPrefValue", StSettingsFile)
        PrhLocationPercPrefValue = fGetINISetting(location, "hLocationPercPrefValue", StSettingsFile)

        ' font
        PrPrefsFont = fGetINISetting(location, "prefsFont", StSettingsFile)
        PrPrefsFontSize = fGetINISetting(location, "prefsFontSize", StSettingsFile)
        PrPrefsFontItalics = fGetINISetting(location, "prefsFontItalics", StSettingsFile)
        PrPrefsFontColour = fGetINISetting(location, "prefsFontColour", StSettingsFile)
        
        ' sound
        PrEnableSounds = fGetINISetting(location, "enableSounds", StSettingsFile)
        
        ' development
        PrDebug = fGetINISetting(location, "debug", StSettingsFile)
        PrDblClickCommand = fGetINISetting(location, "dblClickCommand", StSettingsFile)
        PrOpenFile = fGetINISetting(location, "openFile", StSettingsFile)
        PrDefaultEditor = fGetINISetting(location, "defaultEditor", StSettingsFile)
        
        ' other
        PrMaximiseFormX = fGetINISetting(softwarePlanet, "maximiseFormX", StSettingsFile)
        PrMaximiseFormY = fGetINISetting(softwarePlanet, "maximiseFormY", StSettingsFile)
        PrLastSelectedTab = fGetINISetting(location, "lastSelectedTab", StSettingsFile)
        PrSkinTheme = fGetINISetting(location, "skinTheme", StSettingsFile)
        
        ' window
        PrWindowLevel = fGetINISetting(location, "windowLevel", StSettingsFile)
        PrPreventDragging = fGetINISetting(location, "preventDragging", StSettingsFile)
        PrOpacity = fGetINISetting(location, "opacity", StSettingsFile)
        
        ' we do not want the widget to hide at startup
        'PrWidgetHidden = fGetINISetting(location, "widgetHidden", StSettingsFile)
        PrWidgetHidden = "0"
        
        PrHidingTime = fGetINISetting(location, "hidingTime", StSettingsFile)
        PrIgnoreMouse = fGetINISetting(location, "ignoreMouse", StSettingsFile)
         
        PrFirstTimeRun = fGetINISetting(location, "firstTimeRun", StSettingsFile)
        
    End If

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Module common2"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : beededea
' Date      : 17/06/2020
' Purpose   : validate the relevant entries from the settings.ini file in user appdata
'---------------------------------------------------------------------------------------
'
Public Sub validateInputs()
    
   On Error GoTo validateInputs_Error
            
        ' general
        If PrGaugeFunctions = vbNullString Then PrGaugeFunctions = "1" ' always operate
        If PrsaturnSelection = vbNullString Then PrsaturnSelection = "0"
        If PrStartup = vbNullString Then PrStartup = "1"
        
        ' Config
        If PrEnableTooltips = vbNullString Then PrEnableTooltips = "1"
        If PrEnableBalloonTooltips = vbNullString Then PrEnableBalloonTooltips = "1"
        If PrShowTaskbar = vbNullString Then PrShowTaskbar = "0"
        
        If PrGaugeSize = vbNullString Then PrGaugeSize = "25"
        If PrScrollWheelDirection = vbNullString Then PrScrollWheelDirection = "1"
               
        ' fonts
        If PrPrefsFont = vbNullString Then PrPrefsFont = "times new roman" 'prefsFont", StSettingsFile)
        If PrPrefsFontSize = vbNullString Then PrPrefsFontSize = "8" 'prefsFontSize", StSettingsFile)
        If PrPrefsFontItalics = vbNullString Then PrPrefsFontItalics = "false"
        If PrPrefsFontColour = vbNullString Then PrPrefsFontColour = "0"

        ' sounds
        If PrEnableSounds = vbNullString Then PrEnableSounds = "1"

        ' position
        If PrAspectHidden = vbNullString Then PrAspectHidden = "0"
        If PrWidgetPosition = vbNullString Then PrWidgetPosition = "0"
        If PrWidgetLandscape = vbNullString Then PrWidgetLandscape = "0"
        If PrWidgetPortrait = vbNullString Then PrWidgetPortrait = "0"
        If PrLandscapeFormHoffset = vbNullString Then PrLandscapeFormHoffset = vbNullString
        If PrLandscapeFormVoffset = vbNullString Then PrLandscapeFormVoffset = vbNullString
        If PrPortraitHoffset = vbNullString Then PrPortraitHoffset = vbNullString
        If PrPortraitYoffset = vbNullString Then PrPortraitYoffset = vbNullString
        If PrvLocationPercPrefValue = vbNullString Then PrvLocationPercPrefValue = vbNullString
        If PrhLocationPercPrefValue = vbNullString Then PrhLocationPercPrefValue = vbNullString
                
        ' development
        If PrDebug = vbNullString Then PrDebug = "0"
        If PrDblClickCommand = vbNullString Then PrDblClickCommand = vbNullString
        If PrOpenFile = vbNullString Then PrOpenFile = vbNullString
        If PrDefaultEditor = vbNullString Then PrDefaultEditor = vbNullString
        If PrPreventDragging = vbNullString Then PrPreventDragging = "0"
        
        ' window
        If PrWindowLevel = vbNullString Then PrWindowLevel = "1" 'WindowLevel", StSettingsFile)
        If PrOpacity = vbNullString Then PrOpacity = "100"
        If PrWidgetHidden = vbNullString Then PrWidgetHidden = "0"
        If PrHidingTime = vbNullString Then PrHidingTime = "0"
        If PrIgnoreMouse = vbNullString Then PrIgnoreMouse = "0"
        
        ' other
        If PrFirstTimeRun = vbNullString Then PrFirstTimeRun = "true"
        If PrLastSelectedTab = vbNullString Then PrLastSelectedTab = "general"
        If PrSkinTheme = vbNullString Then PrSkinTheme = "dark"
        
   On Error GoTo 0
   Exit Sub

validateInputs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of form modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file and assign to a global var
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
    On Error GoTo getToolSettingsFile_Error
    ''If debugflg = 1  Then Debug.Print "%getToolSettingsFile"
    
    Dim iFileNo As Integer: iFileNo = 0
    
    PrSettingsDir = fSpecialFolder(feUserAppData) & "\Saturn" ' just for this user alone
    StSettingsFile = PrSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(PrSettingsDir) Then
        MkDir PrSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(StSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open StSettingsFile For Output As #iFileNo
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form modMain"

End Sub



'
'---------------------------------------------------------------------------------------
' Procedure : configureTimers
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : configure any global timers here
'---------------------------------------------------------------------------------------
'
Private Sub configureTimers()

    On Error GoTo configureTimers_Error
    
    oldPrSettingsModificationTime = FileDateTime(StSettingsFile)

    frmTimer.rotationTimer.Enabled = True
    frmTimer.settingsTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

configureTimers_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configureTimers of Module modMain"
            Resume Next
          End If
    End With
 
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : setHidingTime
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
'---------------------------------------------------------------------------------------
'
Private Sub setHidingTime()
    
    On Error GoTo setHidingTime_Error

    If PrHidingTime = "0" Then minutesToHide = 1
    If PrHidingTime = "1" Then minutesToHide = 5
    If PrHidingTime = "2" Then minutesToHide = 10
    If PrHidingTime = "3" Then minutesToHide = 20
    If PrHidingTime = "4" Then minutesToHide = 30
    If PrHidingTime = "5" Then minutesToHide = 60

    On Error GoTo 0
    Exit Sub

setHidingTime_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHidingTime of Module modMain"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : createFormOnCurrentDisplay
' Author    : beededea
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createFormOnCurrentDisplay()
    On Error GoTo createFormOnCurrentDisplay_Error

    With New_c.Displays(1) 'get the current Display
      fMain.InitAndShowAsFreeForm .WorkLeft, .WorkTop, 1000, 1000, "Saturn Widget"
    End With

    On Error GoTo 0
    Exit Sub

createFormOnCurrentDisplay_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createFormOnCurrentDisplay of Module modMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : handleUnhideMode
' Author    : beededea
' Date      : 13/05/2023
' Purpose   : when run in 'unhide' mode it writes the settings file then exits, the other
'             running but hidden process will unhide itself by timer.
'---------------------------------------------------------------------------------------
'
Private Sub handleUnhideMode(ByVal thisUnhideMode As String)
    
    On Error GoTo handleUnhideMode_Error

    If thisUnhideMode = "unhide" Then     'parse the command line
        PrUnhide = "true"
        sPutINISetting softwarePlanet, "unhide", PrUnhide, StSettingsFile
        Call saturnForm_Unload
        End
    End If

    On Error GoTo 0
    Exit Sub

handleUnhideMode_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure handleUnhideMode of Module modMain"
            Resume Next
          End If
    End With
End Sub
