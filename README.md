http://sandsprite.com/vb-reversing/
**Note I have not virus scanned or tested all of the apps linked to in this page! user beware!**  
  
So today we are going to cover some resources on reversing VB6.  
  
VB6 can be compiled to either native or pcode. Even when native the program flow is not linear from the entry point but is based on form, module and class structures passed into the vb runtime at startup. IDA can have a hard time locating everything however the [vb.idc](./files/vb.idc) by Reginald Wong helps a lot. I am not aware of any vb6 pcode loaders for IDA yet so I would not recommend it for this case.  
  
While these structures are not documented by MS reverse engineers have stepped in and figured a lot of its inner operations.  
  
If you want to really get to the core of it take a look at the following resources:  
  

*   Visual Basic Image Internal Structure Format - Alex Ionescu ([local copy](./files/Alex_Ionescu_vb_structures.pdf))
*   VISUAL BASIC REVERSED - A decompiling approach - Andrea Geddon ([local copy](./files/VISUAL BASIC REVERSED.pdf))
*   VB P-code Information by Mr Silver ([local copy](./files/VB P-code Information by Mr Silver.html))
*   DISASSEMBLING VISUAL BASIC APPLICATIONS - Sanchit Karve ([link](https://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66081&lngWId=1), [local copy](./files/DISASSEMBLING VISUAL BASIC APPLICATIONS by Sanchit Karve.zip))
*   DISASSEMBLING VISUAL BASIC APPLICATIONS II - Sanchit Karve ([link](https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67074&lngWId=1), [local copy](./files/DISASSEMBLING VISUAL BASIC APPLICATIONS II.zip))
*   General Detection of Internet Worms - VB 2002 ([pdf](https://www.virusbulletin.com/uploads/pdf/magazine/2002/200201.pdf))
*   VB Wearing the Inside Out - Richard Marko Eset software ([pdf](https://www.virusbulletin.com/uploads/pdf/magazine/2002/200206.pdf))
*   VB6 runtime with debug symbols ([local copy](./files/msvbvm60.zip), [pcode mdb for offsets/disasm/etc](opcode_db.zip))
*   VB6 IDA script by Reginald Wong ([local copy](./files/vb.idc))
*   Open source Semi-VbDecompiler by vbgamer45 ([home page](https://www.theprogrammingzone.com/products/semivbdecompiler/), [source](https://github.com/VBGAMER45/Semi-VB-Decompiler), [installer](https://github.com/VBGAMER45/Semi-VB-Decompiler/raw/master/Semi%20VB%20Decompiler/SemiVBDecompilerInstall.exe))
*   [Microsoft P-Code Technology (msdn\_c7pcode2.asp)](./files/Microsoft P-Code Technology_msdn_c7pcode2.pdf)
*   VB6 Tracer presentation by Jurriaan Bremer and Marion Marschalek. ([pdf](https://github.com/jbremer/vb6tracer/raw/master/presentation/area41.pdf))
*   VB6 pcode opcode database by MrUnleaded, Moogman, and Napalm ([local copy](./files/Opcode Database.htm))
*   [Microsoft's P-Code Implementation](./johnChamberlain/) - [John Chamberlain (2001)](http://johnchamberlain.com/vb_vm.html)
*   [Eternal Bliss VB Crackme 10](Eternal Bliss VB Crackme 10.html)
*   [More resources](http://program-transformation.org/Transform/VisualBasicDecompilers)
*   [Vb compiler leaks info from host](https://sysenter-eip.github.io/VBParser) - sysenter-eip ([local copy](./VBParser/), [vb6.dbg, vba.dbg symbols](./VBParser/debug-symbols-vb.zip) )
*   [Python VB](https://github.com/williballenthin/python-vb/) - willi ballenthin
*   [Vb func for reversing](https://hvoidcode.wordpress.com/2016/02/06/vb-function-description-for-reversing/) - hvoid Enigma Group ([local copy](EnigmaGroup_VB function description for reversing.pdf))
*   [Obscure VB commands](https://www.aivosto.com/articles/obscure-vb.html) - Aivosto Oy ([local copy](./files/obscure_vb_commands.pdf))
*   [VB6 Data types](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60)?redirectedfrom=MSDN) ([local copy](VBA_Vartypes_Reference.pdf))
*   [Ruby + EB](http://bbs.vbstreets.ru/viewtopic.php?f=101&t=56551) - Hacker - vbstreets.ru ([EN translation](https://bbs-vbstreets-ru.translate.goog/viewtopic.php?f=101&t=56551&_x_tr_sl=auto&_x_tr_tl=en&_x_tr_hl=en&_x_tr_pto=wapp&_x_tr_sch=http), [local copy](./vbstreets/Ruby+EB.pdf))
*   Symbols for [vb6.dbg](./files/vb6.dbg) and [vba6.dbg](./files/vba6.dbg)
*   [dump of vb5/6 symbols/binaries/idbs](./vb_symbols.zip)
*   to open an old vb6 exe with debug symbols in IDA you must use the MSDIA parser batch file: **\[path\]\\ida.exe -Opdb:msdia %1**
*   (in at least one case, only my 2004 copy of IDA 4.8 would properly parse the dbg file i included my idb above)

Another great resource on the vb6 structures is the retired [vb-decompiler.theautomaters.com](./vb-decompiler/) message board. I was able to work with one of the admins to restore the board to flat html files so all that great info was not lost to the ages. The files are also available for download in both [zip](./vbdecompiler.zip) and [chm](./vbdecompiler.chm) format. Many thanks to MrUnleaded for helping resurrect this data!  
  
So first thing..we need something that parses these VB6 structures and can tell us the project structure and where each function is in the binary. This is valid for both native and pcode projects and is all defined in the vb structures. From this we can determine the number of forms, classes, user controls, modules etc. For forms we can recover the form definition and objects used, along with their properties. At this stage we can also recover offsets for the code in all of the functions and even names for things like event handlers or public classes (in ActiveX dlls)  
  
For this we can use a commercial app such as [vb-decompiler.org](https://www.vb-decompiler.org/) ($99 home/$230 business) or [VB-Reformer](https://qiil.io/VBReFormer.aspx) (free/$58 pro)  
  
Free alternatives also exist. The formerly commercial, now open source [semi-vbdecompiler](https://github.com/VBGAMER45/Semi-VB-Decompiler) by vbgamer45 looks very cool. ([installer here](https://github.com/VBGAMER45/Semi-VB-Decompiler/raw/master/Semi%20VB%20Decompiler/SemiVBDecompilerInstall.exe))  
  
Also be sure to check out [p32dasm](http://progress-tools.x10.mx/p32dasm.html) and [racevb](http://www.racevb6.com/). For the sake of history you can also check out [exdec by Josephco](./files/betaexdec.zip) (possible [newer version here](./files/ExDec.zip))  
  
There also exists some runtime analysis tools for VB6 apps. For native code [Numega SmartCheck](https://www.google.com/search?q=numega+smart+check+download&ie=utf-8&oe=utf-8&client=firefox-b-ab) is an interesting tool that may be of some help if you can find it. Another tool I havent had a chance to play with yet is [vb6tracer](https://github.com/jbremer/vb6tracer) by Jurriaan Bremer and Marion Marschalek.  
  
My favorite tool, and one that was like magic to me back in the day, is the WKTVBDE pcode [v1.4](./files/WKTVBDE1.4.zip) and [v4.3](./files/WKTVBDE4.3.rar) debugger by Mr. Silver & Mr. Snow. Also check out this tools [help file](./WKTVBDebugger_help_file/) for a listing of the vb opcodes, number of arguments etc. The [VB P-code Information](./files/VB P-code Information by Mr Silver.html) article by Mr Silver also goes into details on how the vb6 runtime processes pcode and how they implemented their debugger.  
  
If your working with older versions of VB different tools apply. VB3 had a decompiler called [DoDi's VB decompiler](./files/Vb3dec.zip). ([newer version](./files/VBDIS3.67e_Reloaded_Rev3_DoDi_s_VB3Decompiler.7z). There is also an untested port of this is the vb semi-decompiler repo as well as [on psc](https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=38543&lngWId=1) ([local copy](./files/DoDi's_vb31251478312002.zip))  
  
That should fill your noodle for a while.  
  
Research I still want to do:

*   pour over the open source semi-vbdecompiler
*   probe the inner operation of the vb6 runtime and WKTVBDE debugger.
*   experiment with runtime tracing as found in smartcheck and vb6tracer.
*   get a firm understanding of vb pcode and try writing a decompiler for it

The WKTVBDE debugger design is interesting. A loader is used to inject the debugger dll into the target process. The target isnt actually running using the debug api, but a series of hooks on all of the vbruntime pcode handlers which are held in a table. The entire UI it presents and which controls program flow is running in the target process! What they wrote is quite extensive, debugging that would have been painful! I tried contacting the authors to see if they would ever consider releasing it open source since its so old and hasnt been updated in a long time. I never heard back though. I might end up recreating it.  
  
Smartcheck uses a combination of techniques. It runs the target process in its own debugger, but also injects some substantial hook dlls. If you have time run [hook explorer](https://github.com/dzzie/HookExplorer) over a binary that is under analysis by smart check. Its interesting.  
  
If you are dealing with vba pcode in office thats a different thing. You might want to look at [pcodedmp](https://github.com/bontchev/pcodedmp)  
  

* * *

* * *

  
  
As kind of a catch all I am going to include some vb5 materials here as well.

*   [Vb5\_binaries\_w\_symbols\_sp3.zip](./files/Vb5_binaries_w_symbols_sp3.zip)
*   [vbsp3.exe (VB5 IDE SP3)](./files/vbsp3.exe)
*   [msvbvm50.dll with full symbols](./files/msvbvm50_w_symbols.zip) (including .Engine opcode handlers)
*   [full debug symbols for vb5 sp3 ide???](./files/symbols_Vb5sp3ds.exe) (listing below)
*   [full debug symbols for vb5 sp2 ide???](./files/VB5SP2DS.EXE)
*   [Q188588: VB5SP3DS.EXE Contains Visual Basic SP3 Debugging Symbols](https://jeffpar.github.io/kbarchive/kb/188/Q188588/)
*   [Q176547: Vb5sp2ds.exe Contains Visual Basic SP2 Debugging Symbols](https://jeffpar.github.io/kbarchive/kb/176/Q176547/)
*   [Super old mirrors ftp.microsoft.com/Softlib/MSLFILES/](http://ftpmirror.your.org/pub/misc/ftp.microsoft.com/Softlib/MSLFILES/)
*   [Matching debug information](http://www.debuginfo.com/articles/debuginfomatch.html) (handy and related)

Vb5sp3ds Contents:
	\\dll\\MSCDRun.dbg
	\\dll\\MSConDes.dbg
	\\dll\\MSRDO20.dbg
	\\dll\\MSVBVM50.dbg
	\\dll\\VB5DB.dbg
	\\dll\\vb5en.dbg
	\\dll\\vb5ide.dbg
	\\dll\\VBA5.dbg
	\\exe\\VB5.dbg
	\\ocx\\comct232.dbg
	\\ocx\\ComCtl32.dbg
	\\ocx\\ComDlg32.dbg
	\\ocx\\DBList32.dbg
	\\ocx\\mci32.dbg
	\\ocx\\MSComm32.dbg
	\\ocx\\msflxgrd.dbg
	\\ocx\\msinet.dbg
	\\ocx\\msmapi32.dbg
	\\ocx\\MSMask32.dbg
	\\ocx\\MSRDC20.dbg
	\\ocx\\mswinsck.dbg
	\\ocx\\PicClp32.dbg
	\\ocx\\RichTx32.dbg
	\\ocx\\SysInfo.dbg
	\\ocx\\tabctl32.dbg

Also I diffed the names in the pcode.engine between vb5/6. Some name changes but functions looks the same except:  

	table,index,vb6name,vb5name
	5 3     StAryRecMove        InValidExCode       
	5 4     StAryRecCopy        InValidExCode       
	5 1D    Bos                 InValidExCode                 
	5 32    CDargRefUdt         InValidExCode       
	5 33    CVarRefUdt          InValidExCode 

Additionally vb6 adds more entries to the last table (vb5 has 0x33, vb6 0x45)

	table,index,InstrSize,xx,name
	5, &H34, 5, 0, "CVarUdt"
	5, &H35, 3, 0, "StUdtVar"
	5, &H36, 3, 1, "StAryVar" 
	5, &H37, 3, 0, "CopyBytesZero"
	5, &H38, 5, 0, "FLdZeroAry"
	5, &H39, 3, 1, "FStVarZero"
	5, &H3A, 7, 0, "CVarAryUdt"
	5, &H3B, 7, 0, "RedimVarUdt"
	5, &H3C, 7, 0, "RedimPreserveVarUdt"
	5, &H3D, 5, 1, "VarLateMemLdRfVar"
	5, &H3E, 7, 1, "VarLateMemCallLdRfVar"  
	5, &H41, 5, 1, "VarLateMemLdVar"
	5, &H42, 7, 0, "VarLateMemCallLdVar"
	5, &H43, 3, 1, "VarLateMemSt"
	5, &H44, 5, 0, "VarLateMemCallSt"
	5, &H45, 5, 0, "VarLateMemStAd"

This is 21 new pcodes for vb6  
Also 33 new Exports added to vb6:

	\_\_vbaRedimVar2
	\_\_vbaVarLateMemCallLd
	\_\_vbaVarLateMemCallLdRf
	\_\_vbaVarLateMemCallSt
	\_\_vbaVarLateMemSt
	\_\_vbaVarLateMemStAd
	\_\_vbaVarZero
	\_\_vbaAryConstruct2
	\_\_vbaAryRecCopy
	\_\_vbaAryRecMove
	\_\_vbaCVarAryUdt
	\_\_vbaCopyBytesZero
	TipCreateInstanceProject2
	EbGetErrorInfo
	\_\_vbaLdZeroAry
	\_\_vbaRedimPreserveVar2
	\_\_vbaUdtVar
	rtcFormatNumber
	rtcFormatCurrency
	rtcFormatPercent
	rtcFormatDateTime
	rtcWeekdayName
	rtcMonthName
	rtcFilter
	rtcInStrRev
	rtcJoin
	rtcSplit
	rtcReplace
	rtcStrReverse
	rtcRound
	rtcCallByName
	rtcCreateObject2
	rtcStrConvVar2
