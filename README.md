# Lems64

![image](https://github.com/fafalone/Lems64/assets/7834493/c3a9cc6e-5c75-489a-97a8-70c10d4890c3)

### 64-bit compatible twinBASIC port of [Carles PV's Lems](https://github.com/Planet-Source-Code/carles-p-v-a-classic-one-and-sequel__1-61601)

**Current status: twinBASIC Beta 305 fixes the few remaining minor bugs, so Lems/Lems64 is now working near perfectly!**

**Update (08 July 2023):** Added Lems64VC.twinproj -- this is an XL size version with includes an audio volume control slider, since at least on my system it was obnoxiously loud compared to media players and I didn't want to have to adjust the system-wide volume.

**Update (06 June 2023, v1.5.13.7):** Added a volume control slider that sets the per-app volume, so you don't need to change the system volume to lower it. Also now allow arrow keys to scroll view, and A is now scroll left, Z or D for right. This is available only as a Beta source file in Lems64MS.twinproj; I have a few more things I want to do before another major release. 

**LemsEdit update (NEW!):** The repo now has a Release Candidate test version of the Lems level editor. The slider has been replaced with a tB slider, the statusbar with a textbox, and the tab control with a minimal direct API implementation. So no outside dependencies.

**Update (20 May 2023, v1.5.13.6):** Mousewheel fix for 64bit builds, updated URL in about screen since PSC is offline anyway.

**Update (19 May 2023, v1.5.13.5):** Permanent, and correctly implemented, fix for TreeView icons. Now uses my own 64bit port of the VBCCR ImageList control for this. **Requires [twinBASIC Beta 305](https://github.com/twinbasic/twinbasic/releases) or newer to build!**

**Update (19 May 2023, v1.5.13.4):** Temporary workaround for level select icons by setting them manually via API.


## Versions
There are three versions currently in the repository to work around this:

Lems_x64.twinproj - ImageList has been replaced with a Common Controls 6.0 ImageList Control. You will need 64bit Microsoft Office installed to get the 64bit version of this control, and will need to have copied it from the Office virtual file system to the regular system folder and registered it there. 

Lems_x64_N.twinproj - ImageList has been replaced with a 64bit-compatible port of Krool's VBCCR ImageList I made. It's been modified to stand alone as well; none of the other controls from Krool's VBCCR have 64 bit ports by anyone yet.

Lems_x64_N_XL.twinproj - Same as above except main playing screen modified to be larger by 1.5x (several points in code need to be changed on top of simply resizing the screen control).

There's also Lems_ImptWorking.twinproj, which is the immediate import of the working VB6 version, without further modification (32bit only)

Lems64 XL:

![image](https://github.com/fafalone/Lems64/assets/7834493/f5570f16-3412-4a25-a50a-fd9eec87845b)

## Requirements

**Requires [twinBASIC Beta 305](https://github.com/twinbasic/twinbasic/releases) or newer** to run from IDE and build without bugs.

IMPORTANT: If you've been using it in versions prior to 304, note that to fix the coloring issue with level previews, you'll need to delete the .bmp files in \LEVELS, which are cached versions.

## IMPORTANT: Game files required! (How to run)

The GameBase folder in the repository contains all the game files-- graphics, sounds, levels, etc. When you've picked a version and are ready to open/run it, it should be in the same folder as the contents of GameBase. It uses ini files and cache files, so it's not advisable to put multiple versions in the same folder, which is why the game files are stored separately here. 

To state it simply: The .twinproj and/or .exe must be in the same folder as the CONFIG/LEVELS/GFX etc folders. If you want multiple versions that don't share level progress, you can create multiple folders with copies of the items from GameBase with the other twinproj/exe.

Download the [current Release version](https://github.com/fafalone/Lems64/releases) for a ready-to-go directory setup; they include both the .twinproj source files and compiled versions of each.

## Level editor

![image](https://github.com/fafalone/Lems64/assets/7834493/bd772b92-cf68-40f8-a321-2fd2c78a2ea7)


The level editor is now complete and included in the repo and latest release. I replaced the comctllib tab control with a barebones pure-API version, and just used a TextBox as statusbar since it only displays the file path anyway. 

---

Questions, comments, bugs? Don't hesitate to create an issue!

**Enjoy!**


