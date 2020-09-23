<div align="center">

## Replace System Files After Rebooting


</div>

### Description

When you are creating a sort of Setup program, etc. sometimes you want to replace some system files or delete some other files, however if they are in use by Windows at the time, you can't. You need to update the files after rebooting, you know that message that says "Please wait while windows updates your configuration files", the utility that does this is called Wininit, and here's how to use it! I've included a small example that should work under Win9x/ME/NT/2K/XP.
 
### More Info
 
Windows Directory, path of the file to be replaced, path of the file that will replace it (this is not necessary of the file is to be deleted)

Just be careful when replacing system files, you should use a method for checking file versions to avoid replacing newer system files with older ones, this is the MAIN reason why windows is so unstable!


<span>             |<span>
---                |---
**Submitted On**   |2002-08-03 01:30:06
**By**             |[Luis Cantero](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/luis-cantero.md)
**Level**          |Intermediate
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Replace\_Sy113775822002\.zip](https://github.com/Planet-Source-Code/luis-cantero-replace-system-files-after-rebooting__1-1067/archive/master.zip)








