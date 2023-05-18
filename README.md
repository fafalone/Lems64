# Lems64

![image](https://github.com/fafalone/Lems64/assets/7834493/c3a9cc6e-5c75-489a-97a8-70c10d4890c3)

### 64-bit compativle twinBASIC port of [Carles PV's Lems](https://github.com/Planet-Source-Code/carles-p-v-a-classic-one-and-sequel__1-61601)

**Current status: twinBASIC Beta 304 fixes the few remaining minor bugs, so Lems/Lems64 is now working perfectly!**

NOTE: While I've substituted the twinBASIC WinNativeCommonCtls TreeView for the comctl32.ocx TreeView, tB currently has no implementation of the ImageList control. 

## Versions
There are three versions currently in the repository to work around this:

Lems_x64.twinproj - ImageList has been replaced with a Common Controls 6.0 ImageList Control. You will need 64bit Microsoft Office installed to get the 64bit version of this control, and will need to have copied it from the Office virtual file system to the regular system folder and registered it there. 

Lems_x64_N.twinproj - ImageList has been replaced with a 64bit-compatible port of Krool's VBCCR ImageList I made. It's been modified to stand alone as well; none of the other controls from Krool's VBCCR have 64 bit ports by anyone yet.

Lems_x64_N_XL.twinproj - Same as above except main playing screen modified to be larger by 1.5x (several points in code need to be changed on top of simply resizing the screen control).

There's also Lems_ImptWorking.twinproj, which is the immediate import of the working VB6 version, without further modification (32bit only)

## Requirements

Requires twinBASIC Beta 304 or newer to run and build without bugs.

IMPORTANT: If you've been using it in previous versions, note that to fix the coloring issue with level previews, you'll need to delete the .bmp files in \LEVELS, which are cached versions.

## Game files required!

The GameBase folder in the repository contains all the game files-- graphics, sounds, levels, etc. When you've picked a version and are ready to open/run it, it should be in the same folder as the contents of GameBase. It uses ini files and cache files, so it's not advisable to put multiple versions in the same folder, which is why the game files are stored separately here. 

Download the Releases for a ready-to-go directory setup; they include both the .twinproj source files and compiled versions of each.

---

Questions, comments, bugs? Don't hesitate to create an issue!

**Enjoy!**
