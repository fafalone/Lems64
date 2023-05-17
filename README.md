# Lems64
64-bit compativle twinBASIC port of [Carles PV's Lems](https://github.com/Planet-Source-Code/carles-p-v-a-classic-one-and-sequel__1-61601)

This is currently an minimally tested alpha version; mainly to share the source with a couple specific people, but anyone can play around with it.

NOTE: While I've substituted the twinBASIC WinNativeCommonCtls TreeView for the comctl32.ocx TreeView, tB currently has no implementation of the ImageList control. 

There are two versions currently in the repository to work around this:

Lems_x64.twinproj - ImageList has been replaced with a Common Controls 6.0 ImageList Control. You will need 64bit Microsoft Office installed to get the 64bit version of this control, and will need to have copied it from the Office virtual file system to the regular system folder and registered it there. 

Lems_x64_N.twinproj - ImageList has been replaced with a 64bit-compatible port of Krool's VBCCR ImageList I made. It's been modified to stand alone as well; none of the other controls from Krool's VBCCR have 64 bit ports by anyone yet.

**Current status:** 64bit build works, but has same cursor issue as 32bit. Level editor not yet ported; it should be working in 32bit though if you import it.
