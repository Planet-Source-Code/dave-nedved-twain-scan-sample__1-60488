Attribute VB_Name = "modTwain"
'// ==================================================================================================================================
'//
'// The Twain Module for Visual Basic was Coded by David Nedved...
'// DONT TAKE CREDIT FOR WHAT YOU DIDN'T CREATE!
'// Last Revised May 12 2005
'// The Twain DLL And DLL Source information as below...
'//
'// ==================================================================================================================================
'//
'// I Have included every single Call there is, even though you
'// Might not need them. This is incase you want to modify the code even more.
'//
'// EZTWAIN 1.x is not a product, and is not the work of any company involved
'// in promoting or using the TWAIN standard.  This code is sample code,
'// provided without charge, and you use it entirely at your own risk.
'// No rights or ownership is claimed by the author, or by any company
'// or organization.  There are no restrictions on use or (re)distribution.
'//
'// Download from:    www.dosadi.com
'//
'// Support contact:  support@dosadi.com
'//
'// ==================================================================================================================================
'//
'// The Origonal DLL was Coded in C++ this code is juat a 'Wrap Arround' to Call
'// The Functions From the C++ Dll.
'// Sorry VB Coders, but thats just the way it is...
'//
'// ==================================================================================================================================


Rem The EZTW32.DLL should be in the System directory or the same dir as the Program
Public Declare Function TWAIN_AcquireToFilename Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal bmpFileName As String) As Integer
Public Declare Function TWAIN_IsAvailable Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp As Long) As Long
Public Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, wPixTypes)
Public Declare Function TWAIN_State Lib "EZTW32.DLL" () As Long
Rem #define TWAIN_PRESESSION        1   // source manager not loaded
Rem #define TWAIN_SM_LOADED         2   // source manager loaded
Rem #define TWAIN_SM_OPEN           3   // source manager open
Rem #define TWAIN_SOURCE_OPEN       4   // source open but not enabled
Rem #define TWAIN_SOURCE_ENABLED    5   // source enabled to acquire
Rem #define TWAIN_TRANSFER_READY    6   // image ready to transfer
Rem #define TWAIN_TRANSFERRING      7   // image in transit



'//--------- DIB handling utilities ---------
Public Declare Function TWAIN_DibDepth Lib "EZTW32.DLL" (ByVal hdib) '// Depth of DIB, in bits i.e. bits per pixel.
Public Declare Function TWAIN_DibWidth Lib "EZTW32.DLL" (ByVal hdib) '// Width of DIB, in pixels (columns)
Public Declare Function TWAIN_DibHeight Lib "EZTW32.DLL" (ByVal hdib) '// Height of DIB, in lines (rows)
Public Declare Function TWAIN_DibNumColors Lib "EZTW32.DLL" (ByVal hdib) '// Number of colors in color table of DIB
Public Declare Function TWAIN_RowSize Lib "EZTW32.DLL" (ByVal hdib)
Public Declare Function TWAIN_ReadRow Lib "EZTW32.DLL" (ByVal hdib, nRow, prow)
Rem // Read row n of the given DIB into buffer at prow.
Rem // Caller is responsible for ensuring buffer is large enough.
Rem // Row 0 is the *top* row of the image, as it would be displayed.



'//--------- BMP file utilities ---------
Public Declare Function TWAIN_WriteNativeToFilename Lib "EZTW32.DLL" (ByVal hdib, pszFile)
Rem // Writes a DIB handle to a .BMP file
Rem //
Rem // hdib     = DIB handle, as returned by TWAIN_AcquireNative
Rem // pszFile  = far pointer to NUL-terminated filename
Rem // If pszFile is NULL or points to a null string, prompts the user
Rem // for the filename with a standard file-save dialog.
Rem //
Rem // Return values:
Rem //   0  success
Rem //  -1  user cancelled File Save dialog
Rem //  -2  file open error (invalid path or name, or access denied)
Rem //  -3  (weird) unable to lock DIB - probably an invalid handle.
Rem //  -4  writing BMP data failed, possibly output device is full
Public Declare Function TWAIN_WriteNativeToFile Lib "EZTW32.DLL" (ByVal hdib, FH)
Public Declare Function TWAIN_LoadNativeFromFilename Lib "EZTW32.DLL" (ByVal pszFile)
Public Declare Function TWAIN_LoadNativeFromFile Lib "EZTW32.DLL" (ByVal FH)

Rem //--------- Application Registration

Public Declare Function TWAIN_RegisterApp Lib "EZTW32.DLL" (ByVal nMajorNum, nMinorNum, nLanguage, nCountry, lpszVersion, lpszMfg, lpszFamily, lpszProduct)
Rem // Reg Info in Same Order as String(s)
Rem // major and incremental revision of application.
Rem // e.g. version 4.5: nMajorNum = 4, nMinorNum = 5
Rem // (human) language (use TWLG_xxx from TWAIN.H)
Rem // country (use TWCY_xxx from TWAIN.H)
Rem // version info string e.g. "1.0b3 Beta release"
Rem // name of mfg/developer e.g. "Crazbat Software"
Rem // product family e.g. "BitStomper"
Rem // specific product e.g. "BitStomper Deluxe Pro"
Rem //
Rem // TWAIN_RegisterApp can be called *AS THE FIRST CALL*, to register the
Rem // application. If this function is not called, the application is given a
Rem // 'generic' registration by EZTWAIN.
Rem // Registration only provides this information to the Source Manager and any
Rem // sources you may open - it is used for debugging, and (frankly) by some
Rem // sources to give special treatment to certain applications.

Rem //--------- Error Analysis and Reporting ------------------------------------

Public Declare Function TWAIN_GetResultCode Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_GetConditionCode Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_ErrorBox Lib "EZTW32.DLL" (ByVal pzMsg) As Long 'Posts a Error Mesage with an ! mark, and OK Button
Public Declare Function TWAIN_ReportLastError Lib "EZTW32.DLL" (ByVal pzMsg) As Long 'Like TWAIN_ErrorBox, but if some details are available from TWAIN about the last failure, they are included in the message box.

Rem //--------- TWAIN State Control ------------------------------------

Public Declare Function TWAIN_LoadSourceManager Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_OpenSourceManager Lib "EZTW32.DLL" (ByVal HWND) As Long
Public Declare Function TWAIN_OpenDefaultSource Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_EnableSource Lib "EZTW32.DLL" (ByVal HWND) As Long
Public Declare Function TWAIN_CloseSource Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_UnloadSourceManager Lib "EZTW32.DLL" (ByVal HWND) As Long
Public Declare Function TWAIN_MessageHook Lib "EZTW32.DLL" (ByVal lpmsg) As Long
Public Declare Function TWAIN_WaitForNativeXfer Lib "EZTW32.DLL" (ByVal HWND) As Long
Public Declare Function TWAIN_ModalEventLoop Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_EndXfer Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_AbortAllPendingXfers Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_WriteDibToFile Lib "EZTW32.DLL" (ByVal lpDIB, FH) As Long

Rem //--------- High-level Capability Negotiation Functions --------

Public Declare Function TWAIN_NegotiateXferCount Lib "EZTW32.DLL" (ByVal nXfers) As Long
Public Declare Function TWAIN_NegotiatePixelTypes Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_GetCurrentUnits Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_SetCurrentUnits Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_GetBitDepth Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_GetPixelType Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_SetBitDepth Lib "EZTW32.DLL" (ByVal nBits) As Long
Public Declare Function TWAIN_SetCurrentPixelType Lib "EZTW32.DLL" (ByVal nPixType) As Long
Public Declare Function TWAIN_GetCurrentResolution Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_GetYResolution Lib "EZTW32.DLL" () As Long
Public Declare Function TWAIN_SetCurrentResolution Lib "EZTW32.DLL" (ByVal dRes As Double)
Public Declare Function TWAIN_SetContrast Lib "EZTW32.DLL" (ByVal dCon As Double)
Public Declare Function TWAIN_SetBrightness Lib "EZTW32.DLL" (ByVal dBri As Double)
Public Declare Function TWAIN_SetXferMech Lib "EZTW32.DLL" (ByVal mech) As Long
Public Declare Function TWAIN_XferMech Lib "EZTW32.DLL" () As Long

Rem //--------- Low-level Capability Negotiation Functions --------

Public Declare Function TWAIN_ToFix32 Lib "EZTW32.DLL" (ByVal d As Double)
Public Declare Function TWAIN_Fix32ToFloat Lib "EZTW32.DLL" (ByVal nFIX As Long)

Rem //--------- Lowest-level functions for TWAIN protocol --------

Public Declare Function TWAIN_DS Lib "EZTW32.DLL" (ByVal DG As Long, DAT As Variant, MSG, FAR)
Public Declare Function TWAIN_Mgr Lib "EZTW32.DLL" (ByVal DG As Long, DAT As Variant, MSG, FAR)
