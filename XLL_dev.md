# Excel XLL and the C API

## Resource Trees


``Office Development``
```-- Office Clients``
``. . `-- Excel``
``. . . . `-- Excel XLL SDK`` [-->](https://msdn.microsoft.com/EN-US/library/bb687883.aspx)
``. . . . . . |-- Getting Started`` [-->](https://msdn.microsoft.com/EN-US/library/bb687894.aspx)
``. . . . . . | . |-- Programming with the C API``
``. . . . . . |-- Developing Excel XLLs`` [-->](https://msdn.microsoft.com/EN-US/library/bb687911.aspx)
``. . . . . . | . |-- Excel Programming Concepts``
``. . . . . . | . |-- Working with DLLs``
``. . . . . . | . |-- Accessing XLL Code in Excel``
``. . . . . . | . |-- Calling into Excel from the DLL or XLL``
``. . . . . . | . |-- Creating XLLs``
``. . . . . . | . |-- Multithreading and Memory Management``
``. . . . . . | . |-- Asynchronous User-Defined Functions``
``. . . . . . | . |-- Cluster Safe Functions``
``. . . . . . | . |-- Permitting User Breaks in Lengthy Operations``
``. . . . . . | . |-- Displaying Dialog Boxes from Within an XLL``
``. . . . . . | . |-- Access Excel Instance and Main Window Handles``
``. . . . . . | . `-- Backward Compatibility``
``. . . . . . `-- API Function Reference``[-->](https://msdn.microsoft.com/EN-US/library/bb687837.aspx)

``Windows desktop applications``
`` `-- Get Started``
``. . `-- What's New``
``. . . . `-- Earlier Versions``
``. . . . . . |-- Windows 7 App Quality Cookbook``
``. . . . . . | . `-- Tools, Best Practices, and Guidance``
``. . . . . . | . . . `-- Preventing Memory Leaks in Windows Applications``
``. . . . . . | . . . `-- Preventing Hangs in Windows Applications`` [-->](https://msdn.microsoft.com/en-us/library/windows/desktop/dd744765%28v=vs.85%29.aspx)
``. . . . . . |-- Windows 7 Developer Guide``
``. . . . . . `-- Hilo: Developing C++ Applications for Windows 7``

[A brief introduction to C++ and Interfacing with Excel](http://www.maths.manchester.ac.uk/~ahazel/EXCEL_C++.pdf) Is useful for learning the basics of C++. However, I think the described method of interfacing with Excel is via the Common Object Model, which we would like to avoid.

[Excel Add-in development in C/C++ - Applications in Finance](http://www-f9.ijs.si/~ilija/slike/cs/aaa.pdf)

[Financial Applications of Excel Add-in development in C++](http://ebooks.allfree-stuff.com/eBooks_down/Microsoft%20Office%202007/Financial%20Applications%20Using%20Excel%20Add-in%20Development%20in%20CC++.pdf) - Second Edition of the above book - for excel 2007

## Developing XLLs

It seems the most useful info to start off with is in *Creating XLLs* - In this article, the C api is described. In the words of the article:
>An XLL is a DLL that exports several procedures that are called by Excel... All of these DLL callbacks start with the prefix xlAuto. Only one of these, the command xlAutoOpen, is required. It is called when the add-in is activated, and it is typically used to register XLL functions and commands with Excel and to do other initialization tasks.

The first step is to install the C API SDK, which contains code samples and header files which can be linked to the projects in Visual Studio

[Advice on using VS](https://msdn.microsoft.com/en-us/library/cyz1h6zd%28v=vs.120%29.aspx)

[Blog on making a minimal xll!](http://blogs.msdn.com/b/andreww/archive/2007/12/09/building-an-excel-xll-in-c-c-with-vs-2008.aspx) - A great guide that gave me my minimally working example.

[Making XLLs with MinGW instead of in VS](http://www.mingw.org/wiki/MSVC_and_MinGW_DLLs)

[A tutorial on how to add menus!](http://xll.codeplex.com/SourceControl/changeset/view/10265#144179)

[Which is based upon this older one](https://support.microsoft.com/en-us/kb/178474) (which is similar/the same as the code in SAMPLE/GENERIC from the XLLSDK for registering a menu)

Financial applications also discusses adding menus (~p. 326) but this is only for Excel 2007 and before: it's not for the ribbon. It might still work, but I haven't spent enough time getting it to do so yet.

In fact, the official documentation doesn't describe `xlfAddMenu` - probably because this is only for the 2007 SKD, not the 2013 SDK and is semi-deprecated. In these cases, aparently, the best route is to look at similar code in Excel Macro Functions (XLM?) and translate directly into the C API. [here is the documentation for the XLM macro language](https://support.microsoft.com/en-us/kb/128185). [Here is an example](http://www.wilmott.com/messageview.cfm?catid=10&threadid=37468).

Another option would be to have the excel call some managed code that does the ribbon gui?

Maybe it would be reasonable to use the [Excell add in library](http://xll.codeplex.com/), which has a blog [here](http://xllblog.com/) and a high performance version here [here](). Interesting related discussion about calling the ribbon: [discussion](http://xll.codeplex.com/discussions/375056). [Link](http://blogs.msdn.com/b/jensenh/archive/2006/12/08/using-ribbonx-with-c-and-atl.aspx) [Link](https://msdn.microsoft.com/en-us/library/ee941475%28v=office.14%29.aspx)

[Also an interesting blog on Excel development in general](https://smurfonspreadsheets.wordpress.com/) (Now shutting down I think).

[Codematic has an overview of excel development](http://www.codematic.net/excel-development/excel-xll/excel-xll.htm) (Old, I think)

Apparently use some GUI to create an XML description of the ribbon which is then bundled with the xll file?? [Stack Exchange](http://stackoverflow.com/questions/21270017/excel-addin-with-ribbon-menu)

Interesting example of an XLL - [asynchronous UDFs](https://code.msdn.microsoft.com/office/Excel-2010-Writing-791e9222)
https://code.msdn.microsoft.com/Excel-2010-Writing-791e9222

https://blogs.office.com/2010/01/27/programmability-improvements-in-excel-2010/
http://stackoverflow.com/questions/18026991/excelasyncutil-observe-to-create-a-running-clock-in-excel
https://msdn.microsoft.com/en-us/library/ff796219.aspx
http://www.remkoweijnen.nl/blog/2012/06/08/excel-2010-multi-threaded-calculation/
http://www.maths.manchester.ac.uk/~ahazel/EXCEL_C++.pdf
http://www.drdobbs.com/parallel/improving-futures-and-callbacks-in-c-to/240004255
asynchronous I/O
http://origin.www.ms.akadns.net/downloads/en/details.aspx?FamilyID=607b1d3b-3006-4693-b82e-cb47429c63e8 ?
https://code.google.com/archive/p/c-api-tools/
http://stackoverflow.com/questions/10884086/microsoft-excel-c-api-and-visual-studio
https://smurfonspreadsheets.wordpress.com/2007/03/05/excel-xlls-and-the-c-api/
https://www.interactivebrokers.com/download/ExcelApiBeginners.pdf
Feel like I'm repeating myself here.
blogs.msdn.com/b/andreww/archive/2007/12/09/building-an-excel-xll-in-c-c-with-vs-2008.aspx
https://support.microsoft.com/en-us/kb/178474
https://blogs.msdn.microsoft.com/andreww/2007/12/09/building-an-excel-xll-in-cc-with-vs-2008/
blogs.msdn.com/b/andreww/archive/2007/12/13/invoking-native-excel-udfs-from-managed-code-pt1.aspx
http://xll.codeplex.com/discussions/375056
https://smurfonspreadsheets.wordpress.com/2010/02/16/raw-xlls/
http://xll.codeplex.com/discussions/533426
https://msdn.microsoft.com/en-us/library/ee941475%28v=office.14%29.aspx
https://blogs.msdn.microsoft.com/jensenh/2006/12/08/using-ribbonx-with-c-and-atl/
http://www.rondebruin.nl/win/s2/win003.htm
http://www.codeproject.com/Articles/76252/What-are-TCHAR-WCHAR-LPSTR-LPWSTR-LPCTSTR-etc
http://www.getcodesamples.com/src/EFD78B79/1FE3D2A9