# Win32 API Declarations for VBA
This project is tracking the Windows Win32 API declaration statements for the Visual Basic for Applications (VBA) programming language used in the Microsoft Office applications. These Win32 API declarations have changed over the years, with the introduction of 64-Bit versions of Windows and MS Office, and subsequent changes to the Windows API.  These Win32 API declaration statements are intended to be used with VBA within the Microsoft Office applications, such as MS Access, Excel, Outlook, and Word, and are dependent on the exact version of Office and Windows you are using.  For versions of MS Office 2007 and older, you should use the declaration statements in the "win32api.txt" file.  For versions of MS Office 2010 and later, you should use the declaration statements in the "win32api_ptrsafe.txt" file.

### *** WARNING *** 
__USING THE WIN32 API FROM MS OFFICE CAN BE DANGEROUS IF NOT USED CORRECTLY, AND CAN CAUSE YOUR PROGRAM TO CRASH, CAUSE YOUR DOCUMENTS TO BECOME CORRUPT, CAUSE LOSS OF DATA, AND/OR CAUSE LOSS OF WORK! Be sure to fully research and thoroughly test any usage of the Win32 API within your MS Office applications.__



## Where to View or Download:

- For versions of __Microsoft Office 2010 and later__, use these declarations: [Win32api_ptrsafe.txt](https://github.com/Access-Abraxas/Win32-API-Declarations-for-VBA/blob/main/win32api_ptrsafe.txt) 

- For versions of __Microsoft Office 2007 and earlier__, use these declarations: [Win32api.txt](https://github.com/Access-Abraxas/Win32-API-Declarations-for-VBA/blob/main/win32api.txt) 



## How to Use:
To use these Win32 API methods within your VBA code, complete the following steps:

1. In the correct version of Win32 API file you need to use (based on the version of MS Office you are developing VBA code for), find the declaration statement you need and copy it into the top of your VBA module, after any "Option" statements.

2. Depending on the specific method declaration statement, you may also need additional type definitions, enumerations, or constant statements for that particular Win32 method.  Be sure to copy any additional required code to the top of your VBA file as well (if any) after the Win32 method declaration statements.

3. Once you've copied all necessary statements out of the Win32 API file, you should be able to begin calling those Win32 APIs methods from within your VBA code.

4. For more information about any specific Win32 API method or type, see the [Microsoft Win32 API documentation](https://learn.microsoft.com/en-us/windows/win32/api/).



## Additional Notes:
Some things to be aware of when using these Win32 declaration statements within your VBA code:

1.  The declarations in the "Win32api.txt" file are extremely old now (from 1994).  You may want to check these declarations against the specific version of Windows you are using them on, if you run into problems with any of the "old" Win32 declarations.

2.  These files are not a complete list of all of the Win32 API methods or data types available for use with VBA, but rather the list of the most commonly used.  For a complete listing of all documented Win32 methods, see the [Microsoft Win32 API documentation here](https://learn.microsoft.com/en-us/windows/win32/api/).

3.  The declarations contained in either of these files are the suggested declarations for each specific Win32 API method, but in many cases, these are __NOT__ the only valid declaration statements of each specific Win32 API method listed in these files.  VBA declare statements, in some cases, can be written multiple ways, using different type parameters, all of which may be valid due to the differences between the data types in C++/Win32 and the data types in VBA.

4.  VBA __CANNOT__ handle exceptions thrown by the Win32 API.  Win32 exceptions will __NOT__ be captured and handled by the VBA Error Handler or Err object.  When an exception is thrown by the Win32 API, it is percolated up to your VBA code and will cause the application to close (crash) without warning.  When this happens, you can use the Windows Event Viewer to view the exception error code number thrown just before the application crashed.  Using that error code, you can then look up that exception number to find out more information about the exception that was thrown.

5.  In Win32 programming, a method call may succeed, but still have errors, which can usually be checked by calling the `GetLastError()` function of the Win32 API.  Depending on the specific method being called, you may want to refer to the [Microsoft Win32 API documentation](https://learn.microsoft.com/en-us/windows/win32/api/) to determine if your specific method requires a call to the `GetLastError()` function to check for errors after it succeeds.



## Project Contributors:
A great big **THANK YOU** to all the people and entities that have helped with this project:

1. [Microsoft](https://microsoft.com) - For providing these Win32 API declarations free of charge to the MS Office developer community.

2. [Geoffrey Griffith](https://geoffreygriffith.com) - For his work to create and maintain this repository for using the Win32 API with VBA.

