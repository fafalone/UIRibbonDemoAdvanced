# Windows UI Ribbon Framework Demo - Advanced w/ Direct2D

![image](https://github.com/user-attachments/assets/dfcb79a6-04a3-4fc6-836a-4d5dad33cb2d)

The Windows UI Ribbon Framework Demo - Advanced is the long awaited final part of my series on using the UI Ribbon in twinBASIC. With the terrible decline in UI in Windows 11, the ribbon is looking pretty great compared to how it is now.

If you're not already familiar with the basics, you'll want to see the Introduction demo, Intermediate Demo, and Galleries Intro demos, which are all in the following repo: https://github.com/fafalone/UIRibbonDemos

This Advanced demo combines the Intermediate and Galleries demo, and builds on those to cover almost all of the remaining features of the ribbon, and a number of bonus features related to the operation of the RichEdit control so that almost all of the related buttons work. 

This project was developed exclusively in twinBASIC; the code takes advantage of new language features wheverever beneficial. It is possible to backport these techniques to use the ribbon in VB6 via oleexp.tlb; the Intro project [has been backported](https://www.vbforums.com/showthread.php?t=900815) to provide a proof of concept and template of how to do it, but the Intermediate, Galleries, and Advanced Ribbon Demos will remain twinBASIC-only projects.
        
## Requirements
-The UI Ribbon is only available in Windows 7 or newer.\
-twinBASIC Beta 677 or newer is required.\
-For Color Font support (e.g. Color Emojis), the riched20.dll and mtpls.dll from Office 2021 or newer must be included in the same folder as the compiled exe. Signed official Microsoft DLLs for both 32 and 64 bit versions are included with this demo.
-To run from the IDE, restart the compiler and build the exe before running.\
-To avoid visual glitching on resizing the Form to larger sizes on Windows 10/11, set the Form HasDC property to  False. This will not affect anything else in 99% of apps, but see further details below or in the frmMain ReadMe for specifics and alternatives.

 ## Changelog
 (Version 4.0.1, 17 Feb 2025) Initial release of Advanced demo.

 ## New Ribbon Features

### Multiple application modes

The Intermediate and Galleries demo have been combined such that there's now two Application Modes: one shows all of the tabs related to the RichEdit, Color Pickers, and Contextual Tabs/Popups, and the other shows only the Galleries tab that controls the shapes, now rendered in a PictureBox. This is accomplished in two parts, first defining which tab and command groups are part of which modes in the XML with the ApplicationModes property. A tab and its command groups can be a part of any or all of 32 possible modes; this demo only has 2. So we have e.g.:

```
<Tab CommandName="cmdTabMain" ApplicationModes="0">
<Tab CommandName="cmdTabGalleries" ApplicationModes="1">
<Group CommandName="cmdShapesGroup" SizeDefinition="InRibbonGalleryAndBigButton" ApplicationModes="1">
```

A ApplicationModes value must be assigned for each tab and group when using multiple modes. The File menu also supports
controlling which modes commands appear in, but in the demo we're keeping the same menu for both so each command has:
```
<MenuGroup>
  <Button CommandName="cmdButtonNew" ApplicationModes="0,1"/>
```
Once the XML is set up, the active mode is controlled programmatically in code. The File menu contains a "Gallery Mode"
option that toggles whether the Ribbon is showing the Galleries tab or the other tabs. Switching modes is easy, it's just
a single call to the Framework object:

```vba
Private Sub SwitchRibbonModes(newMode As eRibbonModes)
    Select Case newMode
        Case RibbonModeTextAndColors
            pFramework.SetModes(UI_MAKEAPPMODE(RibbonModeTextAndColors))
            mRibbonMode = RibbonModeTextAndColors
        Case RibbonModeGalleries
            pFramework.SetModes(UI_MAKEAPPMODE(RibbonModeGalleries))
            mRibbonMode = RibbonModeGalleries
```

The `UI_MAKEAPPMODE` macro is implemented in WinDevLib since it's part of the SDK headers; it's essential since you can have multiple modes active at the same time, represented in a binary 32bit variable, i.e. having only mode 0 active is represented as 1 in binary, having only mode 1 as 10, and both as 11.

### Scaling Policies
If you ever enabled unrestricted resizing on the other demos, you'd see the ribbon is just cut off with an arrow to scroll as you horizontally resize it past the last commands. Scaling policies allow dealing with this more gracefully: commands can be shown in smaller sizes or different arrangements, or groups can become single buttons that display the group normally as a popup. For the group to become a popup, it's desirable to define an image for it, even though these aren't seen in full sized mode. For  the Shapes in ribbon gallery, it now has an associated regular image too. You'll see why in this animation of how the scaling works:

![RibbonResie1](https://github.com/user-attachments/assets/8b4bed3b-4890-4d27-b2f2-6241aeb228f3)

![RibbonResize2](https://github.com/user-attachments/assets/412261b7-617b-46d9-8ba0-561e49d4ed7c)

The scaling policy is defined in the XML. First, there's a block for the ideal sizes if there's no limit on size. After that, the way the size modes switch is defined in the order for which they happen. See for the main tab:

```
<Tab CommandName="cmdTabMain" ApplicationModes="0">
  <Tab.ScalingPolicy>
    <ScalingPolicy>
      <ScalingPolicy.IdealSizes>
        <Scale Group="cmdGroupMain" Size="Large"/>
        <Scale Group="cmdGroupRichFont" Size="Large"/>
        <Scale Group="cmdGroupParagraph" Size="Large"/>
        <Scale Group="cmdCheckHdr" Size="Large"/>
        <Scale Group="CmdGroupEditing" Size="Medium"/>
      </ScalingPolicy.IdealSizes>
      <Scale Group="cmdGroupMain" Size="Medium"/>
      <Scale Group="cmdGroupMain" Size="Popup"/>
      <Scale Group="CmdGroupEditing" Size="Popup"/>
      <Scale Group="cmdGroupParagraph" Size="Popup"/>
      <Scale Group="cmdCheckHdr" Size="Small"/>
      <Scale Group="cmdCheckHdr" Size="Popup"/>
      <Scale Group="cmdGroupRichFont" Size="Medium"/>
      <Scale Group="cmdGroupRichFont" Size="Popup"/>
    </ScalingPolicy>
  </Tab.ScalingPolicy>
  ```
  
Not all commands support all sizes and their groups are restricted accordingly. If a control doesn't support a size option, it can even
override a custom size definition. Rather than go into exhaustive detail, I don't think it would improve on the clarity of just examining
the XML and seeing how it correlates with the video and reading [this documentation on MSDN](https://learn.microsoft.com/en-us/windows/win32/windowsribbon/windowsribbon-templates). 

### New Large Item gallery type
In Wordpad, the List dropdown is unlike any of the galleries in the Galleries Demo in appearance, though it's similar to the Border Type
dropdown in how it's created:

```
<SplitButtonGallery CommandName="CmdList" TextPosition="Hide">
  <SplitButtonGallery.MenuLayout>
    <FlowMenuLayout Gripper="None" Columns="3"/>
  </SplitButtonGallery.MenuLayout>
</SplitButtonGallery>
```
![image](https://github.com/user-attachments/assets/202742e1-9d7e-4f4a-bcb5-8ee9e33d1c6b)

This is another gallery type that's entirely dynamically populated at runtime, so the resource images and strings have to be inserted  manually, and the command IDs custom defined. The code to populate these will be nearly identical to the similar Border Type gallery.

### Automatically populated split button gallery
This type of gallery is similar to a popup menu; it can be defined just by giving the commands in the XML:

```
<Group CommandName="cmdGroupPic" SizeDefinition="OneButton">
  <SplitButton CommandName="CmdInsertPictureMore">
    <SplitButton.ButtonItem>
      <Button CommandName="CmdInsertPicture"/>
    </SplitButton.ButtonItem>
    <Button CommandName="CmdInsertPicture"/>
    <Button CommandName="CmdChangePicture"/>
    <Button CommandName="CmdResizePicture"/>
  </SplitButton>
</Group>
```
![image](https://github.com/user-attachments/assets/bb2a0b09-3f27-4a3d-b01f-496a83082366)

### Spinner Control
This was the last basic control type that hadn't yet been demonstrated. It's trivially simple to add in the XML, just\
`<Spinner CommandName="cmdZoom"/>`\
and it's very similar to the ComboBox control with some identical properties to set, but there's also some additional properties we have to populate in its update properties routine. Min/max value, the increment it changes when you click up/down, and a neat feature it has is the format option, which lets it display e.g. here as a value with %, while still having properties set and retrieved by numeric value.

![image](https://github.com/user-attachments/assets/903d09e3-33ad-44b8-95a4-2ea73c66cfb9)

### Dynamically enabling/disabling controls
In several places, the demo now shows enabling or disabling commands based on events, statuses, or other selected commands. 

-Cut/Copy/Delete are now only enabled when text is selected, set when EN_SELCHANGE is received for the first two. The Delete command presented a problem as you can't set the property until the control is created, and Delete is only created when specific Context Popups are shown. Fortunately, context popups are asynchronous so we can set enable/disable after the .ShowAtLocation call:

```vba
pCtxMenu.ShowAtLocation pt.x, pt.y
If IsTextSelected() Then
    If mEnableCPD = False Then
        pFramework.SetUICommandProperty(IDC_COPY, UI_PKEY_Enabled, CVar(True))
        pFramework.SetUICommandProperty(IDC_CUT, UI_PKEY_Enabled, CVar(True))
        pFramework.SetUICommandProperty(IDC_DELETE, UI_PKEY_Enabled, CVar(True))
        mEnableCPD = True
    End If
Else
    If mEnableCPD = True Then
        pFramework.SetUICommandProperty(IDC_COPY, UI_PKEY_Enabled, CVar(False))
        pFramework.SetUICommandProperty(IDC_CUT, UI_PKEY_Enabled, CVar(False))
        pFramework.SetUICommandProperty(IDC_DELETE, UI_PKEY_Enabled, CVar(False))
        mEnableCPD = False
    End If
End If
```

-The Undo and Redo commands are enabled based on invalidating their enabled property before the context popups they're on are shown, when they're clicked, or on the first edit in case they're in the QAT,  and then setting them according to EM_CANUNDO/EM_CANREDO. 

-Paste is enabled by checking EM_CANPASTE on startup, then using AddClipboardFormatListener to be notified when the clipboard contents change to check again via WM_CLIPBOARDUPDATE. 
          
-The Measurement Units dropdown shows how to create checked menu items, here in a radio fashion where the option you choose becomes the only one checked. (As there's no ruler it has no effect on the textbox however).

-The Activate Context buttons are highlighted by controlling their UI_PKEY_BooleanValue properties. We handle it in Update Properties so the initial context of 1 is  shown as enabled on startup, then explicitly set them with SetUiCommandProperty when one is clicked.

> [!NOTE]
> Some of these features use the EN_CHANGE notification, and it's worth noting serious problems with the documentation for this. First, to receive it at all, you need to enable it by including ENM_CHANGE in the EM_SETEVENTMASK message. But then, you don't receive it through WM_NOTIFY as documented, you receive it through WM_COMMAND, and this is *not* the documented RichEdit version of EN_CHANGE, it's the standard edit control version: lParam contains a handle to the control, not a pointer to a CHANGENOTIFY type.

### Quick Access Toolbar images
As was mentioned for the Scaling Policies, in some cases you may want to associate images with commands even when they're not shown by default. Adding groups and some  commands to the QAT is another one of those cases. In previous demos, you'd see a blank square in many cases. But now groups and commands have been updated to in all but a few cases have images associated with them in the QAT.

![image](https://github.com/user-attachments/assets/5b2a1194-678f-4def-a627-559b9790c2a3)


### Multiple DPI options
All new commands have images specified for different DPIs. This is done in the XML like this:

```
<Command.SmallImages>
  <Image Id="4416">Res/MeasurementUnits_S096.bmp</Image>
  <Image Id="4417" MinDPI="120">Res/MeasurementUnits_S120.bmp</Image>
  <Image Id="4418" MinDPI="144">Res/MeasurementUnits_S144.bmp</Image>
  <Image Id="4419" MinDPI="192">Res/MeasurementUnits_S192.bmp</Image>
</Command.SmallImages>
```

or for galleries populated during runtime, you just select the right image for the current DPI yourself as shown in the code for the galleries.

### New Ribbon events and helper functions 
The RibbonClasses.twin file's main helper class now raises events for `OnRibbonMinMax` and `OnRibbonShowHide`. This is inferred from the ViewChanged event, which is still raised itself. The helper module in the file adds corresponding functions `MinimizeRibbon`, `HideRibbon`, `IsRibbonMinimized` and `IsRibbonVisible`.\
**NOTE:** Min/max and show/hide appear to be non-functional from code on Windows 10; code to set/read them always returns S_OK following documented methods exactly, but the status never changes and retrieving them always gives False. The UI toggle and hide on resize to too small work.
          
It also adds `SetRibbonColors` to set the colors via the `UI_PKEY_GlobalBackgroundColor` / `UI_PKEY_GlobalTextColor` / `UI_PKEY_GlobalHighlightColor`, as well as functions for each indvidual option, and this is accessible via the UI  applying the colors chosen from the Color Pickers, but I think due to theming or a bad RGB to HSB algorithm, only the text color seems to work right.

> [!IMPORTANT]
> Because all the common events are handled by specific helpers, handling the OnViewChanged is now an optional choice between the automatic events and full user handling.\
To receive the OnViewChanged event, you must set the `pUIApp.HandleViewChange` to True, otherwise events are handled by helpers and raised as `OnRibbonMinMax` and `OnRibbonShowHide`, and saving/loading the ribbon state would have to be done fully manually instead of just setting a filename.

Meanwhile WinDevLib has implemented almost all of the remaining helpers and macros from the ribbon SDK headers, most of them being the `UIInitPropertyFrom`* functions and `UIPropertyTo`* functions. Note while these have the same arguments, the PKEY is ignored since there's no language support for validating it like in C++.

### Save / Load Ribbon State 
The Ribbon state (QAT items and position, min/max, etc) can now be saved and loaded automatically.\
To load Ribbon settings, set pUIApp.SettingsFileName immediately after creation.\
To save Ribbon settings, call pUIApp.SaveRibbonSettings from Form_Unload. If you don't specify a file name, the name originally set in pUIApp.SettingsFileName is used.

The settings are saved/loaded via an IStream, so in principle you could save it to the registry, but the demo just uses a file, ribbon.cfg:
```vba
Dim stream As IStream
Dim hr As Long = SHCreateStreamOnFileEx(StrPtr(sFile), STGM_WRITE Or STGM_CREATE, FILE_ATTRIBUTE_NORMAL, CTRUE, Nothing, stream)
If FAILED(hr) Then Return hr 
mRibbon.SaveSettingsToStream(stream)
hr = Err.LastHresult
If FAILED(hr) Then
    stream.Revert()
    Return hr
End If
stream.Commit(STGC_DEFAULT)
```

> [!NOTE]
> This currently isn't set up to handle multiple application instances. See https://learn.microsoft.com/en-us/windows/win32/windowsribbon/ribbon-statepersistence

### Real MRU List
The New/Open/Save/SaveAs buttons are now functional, so the MRU list has been reworked to also be functional with actual files (so note it will be blank at first).\
While a basic MRU is straightforward, pinning support was difficult to add. Whether a file has been pinned is first set when the MRU is loaded in the Update Properties call for `IDC_RECENTITEMS`, which occurs every time the File menu is opened. To determine whether it's changed, we have to look at the Command Execute event; IDC_RECENTITEMS gets an execute event with  UI_PKEY_RecentItems; this gives access to an array of IUISimplePropertySet implementing interfaces that contain the current values, not the ones set when initializing these items in the Update routine. That event is fired whenever the file menu is closed after a click, including pinning buttons, in Recent Items. There we can compare it to what we initially loaded to determine if the pinned status has changed for any item. If it has, we rewrite the MRU in the registry, so it reflects the new status the next time Update is called when the File menu is shown.

![image](https://github.com/user-attachments/assets/d2e05479-c99e-43c9-a764-8b4e8faf130c)
 
There's a lot of difficulty keeping this straight in the backend side. This demo stores MRU entries in the registry with SaveSetting, with pinned items being prefixed with an *. When the list is updated, it's sorted so that the pinned items all come first. This lets us slot unpinned items into the first available slot after that. Numerous situations have to be accounted for and the maximum MRU item limit enforced; such as not adding more items if the maximum is reached by pinned items, and handling an existing file that needs to be bumped up either to the very top if pinned, or to the top of the unpinned items if there's two separate groups. This demo was delayed by several days just handling all of this logic and making sure it generally worked; it  hasn't been exhaustively tested for absolutely every scenario. But the core need was met, showing how to persist a pinned file list that can update pinned status.

### Miscenalleous
- While I don't fully understand why, in the C++ SDK examples, the command to invalidate controls so Update Properties is called always works automatically. In tB, for some reason it doesn't for some controls. I've added workarounds that explicitly set the new states and/or call `FlushPendingInvalidations`, which isn't used in the C++ version but does seem to work to properly trigger the ones that don't fire automatically right away like C++.
         
- I've also verified a manual update technique: You can take the ribbon.bml file generated by uicc.exe and replace the APPLICATION_RIBBON resource with it, then manually add BITMAP files and string table entries with the correct IDs to update the ribbon without rebuilding a whole new project; even in the future when .res files can be imported independently this is useful because of the custom gallery resources and non-ribbon resources that need to be added to most projects.
  
  - While it's generally desirable to keep resource IDs under the signed Integer limit, which requires formally defining blocks like:
    ```
    <Command.LabelTitle>
      <String Id="2855">Zoom</String>
    </Command.LabelTitle>
    ```
    The XML markup now does show the alternative simpler way in one line where the LabelTitle's id will be chosen automatically:\
    `<Command Name="CmdCentimeters" Symbol="IDC_CMD_CENTIMETERS" Id="16" LabelTitle="Centimeters"/>` 
    
    - The customize QAT popup now includes items not enabled by default.
    
    - If you enabled resizing on previous demos or in your own application, you may have noticed visual glitching when the forms enlarges beyond its original size.\
    This demo provides multiple options for a fix:\
    1 (Preferred): Set the Form's `HasDC` property to False. This only affects your app if you do manual, API-based drawing directly to the Form DC and assume GDI objects selected into the DC to remain there after the DC is released and reacquired. It controls whether the form class has the `CS_OWNDC` style set.\
    2: Set the clip region to maximum: from `Form_Load` if `AutoRedraw` is False and the app supports Win8+ only, or from `Form_Resize` if it must be true or support Win7. This demo attempts to handle most scenarios, providing the code in `Form_Load` if it detects `HasDC` is True and AutoRedraw is False, and in `Form_Resize` if it detects `HasDC` and `AutoRedraw` are True.
       
## Bonus RichEdit features
- Color fonts support: This demo comes with a newer RichEdit version that's used in recent Microsoft Office versions and I believe the Windows 11 Notepad that supports color fonts, most commonly used for showing emojis in color. On Windows 10 and earlier, this is typically only available with Office installed, and then only in the same bitness of Office. But if you copy the Office riched20.dll and mtpls.dll, this can be used in your application by loading those DLLs then creating the RichEdit window with the 'RichEditD2D' window class. This demo includes both 32bit and 64bit versions and shows how to load them, or fall back to the regular system richedit options if they're not present. These DLLs are signed with verifiable Microsoft certificates.\
This feature is optional. It can be disabled with the dbg_usenewrichedit option at the top of the main form, or by the absence of the required dlls. If disabled, it falls back to the standard system RichEdit (msftedit.dll if available, or the old riched20.dll in System32).


- The New, Open, and Save/Save As buttons now work. In this demo only saving/loading Rich Text Format (.rtf) and plain text with Unicode encoding  (.txt) is supported,  so the Save As options for Office OpenXML/OpenDocument/Other Formats   don't have actions bound to them, you must pick the main button or the Rich Text or Plain Text buttons for Save As. This shows use of the `EM_STREAMIN` / `EM_STREAMOUT` / `EditStreamCallback` method of saving RTF.\
 Note: The demo doesn't currently track if the editor is 'dirty' i.e. changes have been made to an open document so a save prompt should be displayed.
         
 - The Find/Replace buttons are fully working, just like Notepad and Wordpad they show the system default Find/Replace dialog, using the `FindTextW`/`ReplaceTextW` APIs, which has been set to behave in a similar fashion with all buttons working, via a complex but easy to follow application of `EM_FINDTEXTEX`, `EM_EXSETSEL`, and `EM_REPLACESEL`. Also included is a hooking routine to enable use of  Tab and Enter in the dialog. These are much simpler implementations than in Krool's Common Dialog class/VBCCR, so are much easier to follow to learn how these APIs work, of course coming with a cost: not supporting custom templates, dialog customization callback, or multiple dialogs open at once. 
        
- The new Zoom spinner is functional (via `EM_SETZOOM`), as is the new List format gallery (via `PARAFORMAT2` fields) and Select All button. The Delete button from the older demo is now functional. 

## Loose ends
The pnly thing still not shown I believe are high contrast images. But these are trivially easy to implement if you've come this far; they're just additional image blocks in the XML of specially formatted bitmaps; [see here for details](https://learn.microsoft.com/en-us/windows/win32/windowsribbon/windowsribbon-imageformats#accessibility)

## That's all, folks!

My series on using the Windows UI Ribbon Framework in twinBASIC is now concluded. I'll continue to update this last demo with bug fixes and maybe some additional helper  class and RTF features, but I've now covered everything you need to know to add sophisticated ribbons to your own apps. All in all, the ribbon has definitely grown on me over the years, and while it's complicated, it's still easier than implementing something like it manually, and better than a plain toolbar IMO.

Feel free to create issues with questions or comments about using the ribbon in VB6/twinBASIC in addition to issues specific to this demo.

I hope you've found these demos useful!


