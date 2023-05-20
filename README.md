# Lems64

![image](https://github.com/fafalone/Lems64/assets/7834493/c3a9cc6e-5c75-489a-97a8-70c10d4890c3)

### 64-bit compatible twinBASIC port of [Carles PV's Lems](https://github.com/Planet-Source-Code/carles-p-v-a-classic-one-and-sequel__1-61601)

**Current status: twinBASIC Beta 305 fixes the few remaining minor bugs, so Lems/Lems64 is now working near perfectly!**

**Update (19 May 2023, v1.5.13.6):** Mousewheel fix for 64bit builds, updated URL in about screen since PSC is offline anyway.

**Update (19 May 2023, v1.5.13.5):** Permanent, and correctly implemented, fix for TreeView icons. Now uses my own 64bit port of the VBCCR ImageList control for this. **Requires [twinBASIC Beta 305](https://github.com/twinbasic/twinbasic/releases) or newer to build!**

**Update (19 May 2023, v1.5.13.4):** Temporary workaround for level select icons by setting them manually via API.


## Versions
There are three versions currently in the repository to work around this:

Lems_x64.twinproj - ImageList has been replaced with a Common Controls 6.0 ImageList Control. You will need 64bit Microsoft Office installed to get the 64bit version of this control, and will need to have copied it from the Office virtual file system to the regular system folder and registered it there. 

Lems_x64_N.twinproj - ImageList has been replaced with a 64bit-compatible port of Krool's VBCCR ImageList I made. It's been modified to stand alone as well; none of the other controls from Krool's VBCCR have 64 bit ports by anyone yet.

Lems_x64_N_XL.twinproj - Same as above except main playing screen modified to be larger by 1.5x (several points in code need to be changed on top of simply resizing the screen control).

There's also Lems_ImptWorking.twinproj, which is the immediate import of the working VB6 version, without further modification (32bit only)

## Requirements

**Requires [twinBASIC Beta 305](https://github.com/twinbasic/twinbasic/releases) or newer** to run from IDE and build without bugs.

IMPORTANT: If you've been using it in versions prior to 304, note that to fix the coloring issue with level previews, you'll need to delete the .bmp files in \LEVELS, which are cached versions.

## IMPORTANT: Game files required! (How to run)

The GameBase folder in the repository contains all the game files-- graphics, sounds, levels, etc. When you've picked a version and are ready to open/run it, it should be in the same folder as the contents of GameBase. It uses ini files and cache files, so it's not advisable to put multiple versions in the same folder, which is why the game files are stored separately here. 

To state it simply: The .twinproj and/or .exe must be in the same folder as the CONFIG/LEVELS/GFX etc folders. If you want multiple versions that don't share level progress, you can create multiple folders with copies of the items from GameBase with the other twinproj/exe.

Download the [current Release version](https://github.com/fafalone/Lems64/releases) for a ready-to-go directory setup; they include both the .twinproj source files and compiled versions of each.

## Level editor

I have not yet completed a 64bit port of the level editor, as it uses more complex common controls without existing x64 ports besides the impractical option of installing MS Office 64bit and extracting the ocx from it's virtual file system. For the time being, you can use the original VB6 level editor. It's included in the VB6 folder of the repository, and an exe build is included in the Release versions.

---

Questions, comments, bugs? Don't hesitate to create an issue!

**Enjoy!**
