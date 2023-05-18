# Lems64

![image](https://github.com/fafalone/Lems64/assets/7834493/c3a9cc6e-5c75-489a-97a8-70c10d4890c3)

64-bit compativle twinBASIC port of [Carles PV's Lems](https://github.com/Planet-Source-Code/carles-p-v-a-classic-one-and-sequel__1-61601)

This is currently an minimally tested alpha version; mainly to share the source with a couple specific people, but anyone can play around with it.

NOTE: While I've substituted the twinBASIC WinNativeCommonCtls TreeView for the comctl32.ocx TreeView, tB currently has no implementation of the ImageList control. 

There are three versions currently in the repository to work around this:

Lems_x64.twinproj - ImageList has been replaced with a Common Controls 6.0 ImageList Control. You will need 64bit Microsoft Office installed to get the 64bit version of this control, and will need to have copied it from the Office virtual file system to the regular system folder and registered it there. 

Lems_x64_N.twinproj - ImageList has been replaced with a 64bit-compatible port of Krool's VBCCR ImageList I made. It's been modified to stand alone as well; none of the other controls from Krool's VBCCR have 64 bit ports by anyone yet.

Lems_x64_N_XL.twinproj - Same as above except main playing screen modified to be larger by 1.5x (several points in code need to be changed on top of simply resizing the screen control).

There's also Lems_ImptWorking.twinproj, which is the immediate import of the working VB6 version, without further modification (32bit only)

**Current status:** Game is playable, minor issues with level select form, menus, toolbar coords (use top half), and level preview colors. But the main game itself is fully playable. Level editor not yet ported; it should be working in 32bit though if you import it.
