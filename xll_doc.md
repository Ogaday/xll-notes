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

## Developing XLLs

It seems the most useful info to start off with is in *Creating XLLs* - In this article, the C api is described. In the words of the article:
>An XLL is a DLL that exports several procedures that are called by Excel... All of these DLL callbacks start with the prefix xlAuto. Only one of these, the command xlAutoOpen, is required. It is called when the add-in is activated, and it is typically used to register XLL functions and commands with Excel and to do other initialization tasks.

The first step is to install the C API SDK, which contains code samples and header files which can be linked to the projects in Visual Studio