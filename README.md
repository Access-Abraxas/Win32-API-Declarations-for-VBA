# Win32 API Declarations for VBA
This project is tracking the Win32 API declaration statements for VBA provided by Microsoft for the MS Office applications.  These Win32 API declarations have changed over the years, with the introduction of 64-Bit versions of Windows and MS Office, and changes to the Windows API.  These Win32 API declaration statements are intended to be used with VBA within the Microsoft Office applications, such as MS Access, Excel, Outlook, and Word, and are dependent on the exact version of Office and Windows you are using.  For versions of MS Office 2007 and older, you should use the declaration statements in the "win32api.txt" file.  For versions of MS Office 2010 and later, you should use the declaration statements in the "win32api_ptrsafe.txt" file.

### *** WARNING *** 
__USING THE WIN32 API FROM MS OFFICE CAN BE DANGEROUS AND CAUSE YOUR PROGRAM TO CRASH, CAUSE YOU DOCUMENTS TO BECOME CORRUPT, AND/OR POTENTIALLY CAUSE LOSS OF DATA OR WORK, IF NOT USED CORRECTLY!__



## Where to View or Download:

- For versions of Microsoft Office 2010 and later, use these declarations: [Win32api_ptrsafe.txt](https://github.com/Access-Abraxas/Win32-API-Declarations-for-VBA/blob/main/win32api_ptrsafe.txt) 

- For versions of Microsoft Office 2007 and earlier, use these declarations: [Win32api.txt](https://github.com/Access-Abraxas/Win32-API-Declarations-for-VBA/blob/main/win32api.txt) 


## How to Use:
To use these Win32 API method in you VBA code, complete the following steps:

1. In the correct version of Win32 API file you need to use, find the declaration statement you need and copy it into the top of your VBA module, after an "Option" statements.

2. Also, depending on the declaration statement, you may need type definitions, enums, or constants needed for that particular Win32 method, so copy those to the top of your VBA file as well (if any) after the declaration statements.

3. Once you've copied all necessary statements out of the Win32 API file, you should be able to begin calling those Win32 APIs from methods in your VBA code.


## Additional Notes:
Some things to be aware of when using these Win32 declaration statements with VBA code:

1.  The declarations in the "Win32api.txt" file are extremely old now (from 1994).  You may want to check these declarations against the specific version of Windows you are using them on, if you run into problems with any of the "old" Win32 declarations.

2.  These files are not a complete list of all of the Win32 API methods or data types available for use with VBA, but rather the list of the most commonly used.  For a complete listing of all documented Win32 methods, see the [Microsoft Win32 API documentation here](https://learn.microsoft.com/en-us/windows/win32/api/).

3.  The declarations contained in either of these files are just suggestions about how to declare each specific Win32 method, but these are not the end all to writing these declare statements.  Declare statements, in some cases, can be written multiple ways, using different type paremeters, all of which may be valid due to the difference between data types in C++/Win32 and data types in VBA.


## Project Contributors:
A great big **THANK YOU** to all the people and entities that have helped with this project:

1. [Microsoft](https://microsoft.com) - For providing these Win32 API declarations free of charge to the MS Office developer community.

2. [Geoffrey Griffith](https://geoffreygriffith.com) - For his work to create and maintain this repository for using the Win32 API with VBA.

