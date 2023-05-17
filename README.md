# Lems64
64-bit compativle twinBASIC port of [Carles PV's Lems](https://github.com/Planet-Source-Code/carles-p-v-a-classic-one-and-sequel__1-61601)

This is currently an minimally tested alpha version; mainly to share the source with a couple specific people, but anyone can play around with it.

NOTE: While I've substituted the twinBASIC WinNativeCommonCtls TreeView for the comctl32.ocx TreeView, tB currently has no implementation of the ImageList control. This has been replaced with a Common Controls 6.0 ImageList Control. You will need 64bit Microsoft Office installed to get the 64bit version of this control, and will need to have copied it from the Office virtual file system to the regular system folder and registered it there. I'll look into replacing it with pure API, or provide more detailed instructions, once this is working. Note that 32bit isn't working either at the moment; the game starts but you can't assign jobs to the Lemmings; I made this now in hopes the glitch is better on 64bit.

**Current status:** 64bit build works, but has same cursor issue as 32bit.
