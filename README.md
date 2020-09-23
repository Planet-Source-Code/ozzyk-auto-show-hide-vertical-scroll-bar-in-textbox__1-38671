<div align="center">

## Auto Show/Hide Vertical Scroll Bar in TextBox


</div>

### Description

Auto-show/hide vertical scroll bar in a multi-line TextBox control.

If you set the ScrollBars property of a multi-line TextBox control to 2-Vertical, the vertical scrollbar will be displayed whether or not there's any text in the TextBox, or the length of the text exceeds the visible area of the TextBox. So, if there's more text than the TextBox can display in its client area, then the vertical bar will be enabled, if not, it will still be displayed but disabled. This code subclasses a TextBox to intercept WM_KEYUP and WM_SIZE messages, and compares the height of the text to the height of the client area of the control. If the height of the text is greater than that of the client area, the vertical scroll bar will be displayed. Otherwise it will be hidden. So you don't have to see the disabled scroll bar when there's no need for it. I resorted to subclassing particularly to watch for the resizing of the TextBox. In this sample, the TextBox's size changes when the form is resized. Therefore subclassing may not be necessary as the Form_Resize event is perfectly suitable for doing whatever's needed. Yet the TextBox's size does not always change because of the form being resized. For example; there may be a splitter on the form, and the TextBox may be resized by that splitter. Or for some reason, the size of the TextBox may be programmatically changed. In such cases, I can't think of any way but subclassing to catch the TextBox's resizing.

The portions of the code pertaining to subclassing are, of course, nothing new as they are right out of Bruce McKinney's well-known "Hardcore Visual Basic", though a bit simplified. But for the rest of the code, I'd appreciate all feedback.
 
### More Info
 
Developed with MS Visual Basic 6 (SP5) on Windows XP


<span>             |<span>
---                |---
**Submitted On**   |2002-09-05 05:23:04
**By**             |[OzzyK](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ozzyk.md)
**Level**          |Intermediate
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Auto\_Show\_127006952002\.zip](https://github.com/Planet-Source-Code/ozzyk-auto-show-hide-vertical-scroll-bar-in-textbox__1-38671/archive/master.zip)








