VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMFTSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***
'MFTSearch by Lucas Plumb @ 2023
'clsMFTSearch
'Utilizes DeviceIoControl and some advanced optimization techniques to quickly search drives for files
'Techniques include SAFEARRAY pointer swapping, BinarySearch, LSet, and other API functions tested for performance
'***

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Private Const FILE_ATTRIBUTE_ALL As Long = FILE_ATTRIBUTE_READONLY Or _
                                           FILE_ATTRIBUTE_HIDDEN Or _
                                           FILE_ATTRIBUTE_SYSTEM Or _
                                           FILE_ATTRIBUTE_ARCHIVE Or _
                                           FILE_ATTRIBUTE_NORMAL

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_EXECUTE As Long = &H20000000
Private Const GENERIC_ALL As Long = &H10000000

Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2

Private Const CREATE_NEW As Long = &H1
Private Const CREATE_ALWAYS As Long = &H2
Private Const OPEN_EXISTING As Long = &H3
Private Const OPEN_ALWAYS As Long = &H4
Private Const TRUNCATE_EXISTING As Long = &H5

Private Const NUM_0 As Long = 0&
Private Const NUM_1 As Long = 1&
Private Const NUM_2 As Long = 2&
Private Const NUM_4 As Long = 4&
Private Const NUM_8 As Long = 8&
Private Const NUM_16 As Long = 16&
Private Const NUM_60 As Long = 60&

Private Const INT_2 As Integer = 2


'winioctl.h
'#define FSCTL_QUERY_USN_JOURNAL         CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 61, METHOD_BUFFERED, FILE_ANY_ACCESS)
                                '#define CTL_CODE( DeviceType, Function, Method, Access ) (                 \ _
                                    ((DeviceType) << 16) | ((Access) << 14) | ((Function) << 2) | (Method) \ _
                                )
'FILE_DEVICE_FILE_SYSTEM = 0x00000009
'METHOD_BUFFERED = 0x0
'METHOD_NEITHER  = 0x3
'FILE_ANY_ACCESS = 0x0
'FSCTL_QUERY_USN_JOURNAL = 0x000900F4 (FILE_DEVICE_FILE_SYSTEM << 16 = 0x00090000, 61 << 2 = 0x000000F4
'#define FSCTL_ENUM_USN_DATA = 0x000900B3            CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 44,  METHOD_NEITHER, FILE_ANY_ACCESS) // MFT_ENUM_DATA, (44 << 2 = 0x000000B0, METHOD_NEITHER = 0x00000003)

'                                                                 44 << 2 = 0x000000B8, + 3
'#define FSCTL_READ_USN_JOURNAL = 0x000900BB          CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 46,  METHOD_NEITHER, FILE_ANY_ACCESS) // READ_USN_JOURNAL_DATA, USN

'                                                                       58 << 2 = 0x000000E8, + 3
'#define FSCTL_READ_FILE_USN_DATA        CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 58,  METHOD_NEITHER, FILE_ANY_ACCESS) // Read the Usn Record for a file

Private Const FSCTL_QUERY_USN_JOURNAL As Long = &H900F4
Private Const FSCTL_ENUM_USN_DATA As Long = &H900B3
Private Const FSCTL_READ_USN_JOURNAL As Long = &H900BB
Private Const FSCTL_READ_FILE_USN_DATA As Long = &H900EB

Private Type LARGE_INTEGER
    highpart As Long
    lowpart As Long
End Type

Private Type TFRNID
    ID(5) As Byte
End Type

Private Type TFRN
    ID As TFRNID
    Sequence As Integer
End Type

Private Type TDouble
    Value As Double
End Type

Private Type TFileRef
    FileRef As LARGE_INTEGER
    ParentFileRef As LARGE_INTEGER
End Type

Private Type TFolderRef
    FileRef As TDouble
    ParentFileRef As TDouble
    fileName As String
End Type

Private Type UINT64
    Value(7) As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0) As SAFEARRAYBOUND
End Type

'Private Type FILETIME
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type
'
'Private Type SYSTEMTIME
'  wYear As Integer
'  wMonth As Integer
'  wDayOfWeek As Integer
'  wDay As Integer
'  wHour As Integer
'  wMinute As Integer
'  wSecond As Integer
'  wMilliseconds As Integer
'End Type
'
'Private Type WIN32_FIND_DATA
'   dwFileAttributes As Long
'   ftCreationTime As FILETIME
'   ftLastAccessTime As FILETIME
'   ftLastWriteTime As FILETIME
'   nFileSizeHigh As Long
'   nFileSizeLow As Long
'   dwReserved0 As Long
'   dwReserved1 As Long
'   cFileName As String * MAX_PATH
'   cAlternate As String * 14
'End Type

'Private Type BootSector
'    jump(3) As Byte
'    nName(8) As Byte
'    bytesPerSector As Integer '        // The number of bytes in a sector. This should be 512.
'    sectorsPerCluster As Byte '      // The number of sectors in a cluster. Clusters are used for less-granular access to the partition. They're usually 4KB.
'    reservedSectors As Integer
'    unused0(3) As Byte
'    unused1 As Integer
'    media As Byte
'    unused2 As Integer
'    sectorsPerTrack As Integer
'    headsPerCylinder As Integer
'    hiddenSectors As Long
'    unused3 As Long
'    unused4 As Long
'    totalSectors1 As Long ' (part of a 64bit ulong)
'    totalSectors2 As Long ' (part of a 64bit ulong)
'    mftStart1 As Long '           // The start of the MFT, given as a cluster index. (part of a 64bit ulong)
'    mftStart2 As Long '           // The start of the MFT, given as a cluster index. (part of a 64bit ulong)
'    mftMirrorStart1 As Long ' (part of a 64bit ulong)
'    mftMirrorStart2 As Long ' (part of a 64bit ulong)
'    clustersPerFileRecord As Long
'    clustersPerIndexBlock As Long
'    serialNumber1 As Long ' (part of a 64bit ulong)
'    serialNumber2 As Long ' (part of a 64bit ulong)
'    checksum As Long
'    bootloader(426) As Byte
'    bootSignature As Integer
'End Type

Private Type USN_JOURNAL_DATA_V2
    UsnJournalID As UINT64
    FirstUsn As UINT64
    NextUsn As UINT64
    LowestValidUsn As UINT64
    MaxUsn As UINT64
    MaximumSize As UINT64
    AllocationDelta As UINT64
    MinSupportedMajorVersion As Integer
    MaxSupportedMajorVersion As Integer
    Flags As Long
    RangeTrackChunkSize As UINT64
    RangeTrackFileSizeThreshold As UINT64
End Type

Private Type MFT_ENUM_DATA_V0
    StartFileReferenceNumber As UINT64
    lowUSN As UINT64
    highUSN As UINT64
End Type

Private Type USN_RECORD_V2
    recordLength As Long '4 '0
    MajorVersion As Integer '2 '4
    MinorVersion As Integer '2 '6
    FileReferenceNumber As TFRN '8 '8
    ParentFileReferenceNumber As TFRN '8 '16
    Usn As UINT64 '8 '24
    TimeStamp As UINT64 '8
    Reason As Long '4
    SourceInfo As Long '4
    SecurityId As Long '4
    FileAttributes As Long '4 '52
    fileNameLength As Integer '2 '56
    FileNameOffset As Integer '2 '58
    fileName(1) As Byte '4 '60
End Type

Private Declare Function DeviceIoControl Lib "kernel32" _
    (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByVal lpInBuffer As Long, _
    ByVal nInBufferSize As Long, ByVal lpOutBuffer As Long, ByVal nOutBufferSize As Long, _
    lpBytesReturned As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
    
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
    
Private Declare Function CloseHandle& Lib "kernel32" (ByVal hObject&)


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Sub CopyMemoryArray Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination() As Any, Source As Any, ByVal length As Long)
    
'Private Declare Sub MoveArray Lib "msvbvm60" Alias "__vbaAryMove" (Destination() As Any, Source As Any)
'Private Declare Function vbaCopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal length As Long, dst As Any, src As Any) As Long
'Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Long, ByVal length As Long)

Private Declare Function CompareMemory Lib "ntdll" Alias "RtlCompareMemory" (ptrSrc1 As Any, ptrSrc2 As Any, ByVal cLen As Long) As Long

Private Declare Function StrCmpNIW Lib "shlwapi" (ByVal pSrc1 As Long, _
    ByVal pSrc2 As Long, ByVal CB As Long) As Long
    
Private Declare Function CharLowerBuff Lib "user32" Alias "CharLowerBuffW" ( _
    ByVal pStrBuff As Long, ByVal buffLen As Long) As Long
    
'Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As Long, ByVal psz As Long) As Long
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal length As Long) As Long


    

Public Function Find(findFileName As String, Optional driveLetter As String = "C", Optional exactMatch As Boolean = True, Optional caseMatch As Boolean = True) As String()
    'https://www.autohotkey.com/board/topic/79420-ahk-l-listmftfiles-ntfs-instant-file-search/
    'https://handmade.network/forums/articles/t/7002-tutorial_parsing_the_mft
    Dim returnStr() As String
    
    Dim timer As CTimer
    Set timer = New CTimer
    timer.StartCounter
    
    Dim drive As Long, bytesAccessed As Long, res As Long, journalData As USN_JOURNAL_DATA_V2
    Dim record As USN_RECORD_V2, i_plus_FileNameOffset As Long, pFileNameOffset As Long, fileNameBytes() As Byte, fileNameLen As Integer, fileNameLenLong As Long, fileNameStrLen As Long
    Dim i As Long, j As Long, y As Long
    Dim pFileNameBytes As Long
    Dim lowUSN1 As Double, lowUSN2 As Double
    Dim tempFileRef As TFileRef, tempLargeInt As LARGE_INTEGER, tempLowDbl As Double
    Dim quitSearch As Boolean
    Dim tmpFileNameStr As String, pTmpFileNameStr As Long, tmpFileNameChars As Long
    
    If caseMatch = False Then findFileName = LCase$(findFileName)
    fileNameBytes = findFileName
    fileNameLen = UBound(fileNameBytes) + 1
    fileNameLenLong = UBound(fileNameBytes) + NUM_1
    fileNameStrLen = (UBound(fileNameBytes) + NUM_2) \ NUM_2
    pFileNameBytes = VarPtr(fileNameBytes(NUM_0))
    pTmpFileNameStr = VarPtr(tmpFileNameStr)
    
    Dim folderIndex As Long, maxFolders As Long
    Dim folders(&H1FFFFF) As TFolderRef '(8^7), or 2097152 (1FFFFF = 2097151, plus 1 for 0th index)
    '4,294,967,295 is maxmium number of folders on NTFS, so 2,097,152 doesn't quite add up
    maxFolders = UBound(folders)
    
    Dim fileIndex As Long, maxResults As Long
    Dim files(99) As TFolderRef 'maximum of 100 results to find before we stop searching
    Dim fileParentRef As TFolderRef
    maxResults = UBound(files)
    
    Dim outBuff(&HFFFFFFF) As Byte '(8^8), or 16777216 (FFFFFFF = 16777215, plus 1 for 0th index)
    Dim outBuffPtr As Long, outBuffSize As Long
    outBuffPtr = VarPtr(outBuff(NUM_0))
    outBuffSize = UBound(outBuff) + NUM_1
    
    Dim enumData As MFT_ENUM_DATA_V0, enumDataPtr As Long, enumDataSize As Long
    enumDataPtr = VarPtr(enumData)
    enumDataSize = LenB(enumData)
    
    Dim totalMs As Double, curMs As Double
    
    Dim recordLen As Long
    recordLen = LenB(record)
    
    Dim m_SA As SAFEARRAY1D
    Dim recordSA() As USN_RECORD_V2
    Dim lastPtr As Long
    
    drive = CreateFile("\\.\" & driveLetter & ":", GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal NUM_0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NUM_0)
    
    If drive <= 0 Then
        If Err.LastDllError = 5 Then
            MsgBox "Access denied. FastFileSearch requires administrator privileges. CreateFile result: " & drive & vbCrLf & _
            "Error Code: " & Err.LastDllError & " - Source: " & Err.Source & ", Description: " & Err.Description, , "Error"
        Else
            MsgBox "Opening volume failed. CreateFile result: " & drive & vbCrLf & _
            "Error Code: " & Err.LastDllError & " - Source: " & Err.Source & ", Description: " & Err.Description, , "Error"
        End If
        Exit Function
    End If
    
    res = DeviceIoControl(drive, FSCTL_QUERY_USN_JOURNAL, NUM_0, NUM_0, VarPtr(journalData), LenB(journalData), bytesAccessed, NUM_0)
    
    If res <> NUM_0 Then
    
        If journalData.MaxSupportedMajorVersion < INT_2 Then
            MsgBox "Error: Journal on filesystem does not support USN_RECORD version 2." & vbCrLf & _
            "Search function currently only supports USN_RECORD_V2." _
            , , "Error"
            '"Error Code: " & Err.LastDllError & " - Source: " & Err.Source & ", Description: " & Err.Description, , "Error"
            Exit Function
        End If

        'check lowestvalidusn and firstusn - per the MS DOCs on USN_JOURNAL_DATA_V2:
        'LowestValidUsn - "Enumerating the files or directories on a volume can return a USN lower than this value (in other words, _
        a FirstUsn member value less than the LowestValidUsn member value). If it does, the journal has been stamped with a new identifier _
        since the last USN was written. In this case, LowestValidUsn may indicate a discontinuity in the journal, in which changes to some _
        or all files or directories on the volume may have occurred that are not recorded in the change journal."
        CopyMemory tempLargeInt, journalData.LowestValidUsn.Value(NUM_0), NUM_8
        lowUSN1 = LI2Double(tempLargeInt.lowpart, tempLargeInt.highpart)
        CopyMemory tempLargeInt, journalData.FirstUsn.Value(NUM_0), NUM_8
        lowUSN2 = LI2Double(tempLargeInt.lowpart, tempLargeInt.highpart)
        
        If lowUSN1 < lowUSN2 Then
            'enumData.lowUSN = journalData.LowestValidUsn
        Else
            'enumData.lowUSN = journalData.FirstUsn
            '<TODO> how to appropriately handle discontinuities in change journal records?
            MsgBox "Error: Journal FirstUsn is lower than LowestValidUsn." & vbCrLf & _
            "The journal has been stamped with a new identifier, which may indicate a discontinuity." & vbCrLf & _
            "Some files/folders may not be available." _
            , , "Error"
        End If
        enumData.lowUSN = journalData.LowestValidUsn 'for now, lets just search using LowestValidUsn
        enumData.highUSN = journalData.MaxUsn
        
        res = DeviceIoControl(drive, FSCTL_ENUM_USN_DATA, enumDataPtr, enumDataSize, outBuffPtr, outBuffSize, bytesAccessed)
        
        If res <> NUM_0 Then 'success!
        
            'initialize a safearray struct which we can directly modify the pointer to, marginally faster than CopyMemory to coerce a byte array into UDT
            With m_SA
                .cDims = 1
                .fFeatures = &H1 '&H1 = FADF_AUTO '&H11 = FADF_AUTO Or FADF_FIXEDSIZE
                .cLocks = NUM_1
                .cbElements = NUM_1
                '.pvData = VarPtr(outBuff(i)) 'set to the outbuff pointer - don't need to do this now, we do it a few lines later anyway
                .Bounds(NUM_0).cElements = NUM_60 '64 bytes in the record, but we only need 61 bytes
            End With
            CopyMemoryArray recordSA(), VarPtr(m_SA), NUM_4 'replace recordSA with m_SA
                           
            Do While res <> NUM_0
                'curMs = timer.TimeElapsed
                m_SA.pvData = outBuffPtr + NUM_8 'VarPtr(outBuff(i+8)), outBuff(0 to 7) is what we need to provide to the NEXT call to FSCTL_ENUM_USN_DATA - not part of the record...
                lastPtr = outBuffPtr + bytesAccessed 'end of data = VarPtr(outBuff(0)) + bytesAccessed
                Do While m_SA.pvData < lastPtr
                    With recordSA(NUM_0)
                        If .MajorVersion > INT_2 Then
                            MsgBox "Error: Encountered USN_RECORD version > 2 on filesystem." & vbCrLf & _
                            "Search function currently only supports USN_RECORD_V2." _
                            , , "Error"
                            '"Error Code: " & Err.LastDllError & " - Source: " & Err.Source & ", Description: " & Err.Description, , "Error"
                            Exit Function
                        End If
                        
                        If (.FileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then  '<> 0 then file IS a directory
                            'LSet to get the low 6-bytes of FileReferenceNumber into a double, much faster than CopyMemory
                            LSet folders(folderIndex).FileRef = .FileReferenceNumber.ID
                            LSet folders(folderIndex).ParentFileRef = .ParentFileReferenceNumber.ID
                            'SysReAllocStringLen is faster than copying the filename bytes and converting into a string with VBA
                            SysReAllocStringLen VarPtr(folders(folderIndex).fileName), m_SA.pvData + .FileNameOffset, .fileNameLength \ NUM_2
                            folderIndex = folderIndex + NUM_1
                            If folderIndex > maxFolders Then
                                MsgBox "Error: Number of directories on filesystem exceeds buffer size of " & UBound(folders) + 1 & "." & vbCrLf & _
                                "Search function is unable to resolve file paths without increasing memory allocation." _
                                , , "Error"
                                '"Error Code: " & Err.LastDllError & " - Source: " & Err.Source & ", Description: " & Err.Description, , "Error"
                                Exit Function
                            End If
                        Else
                            If (.FileAttributes And FILE_ATTRIBUTE_ALL) Then '<> 0 then file IS a real file
                                If exactMatch Then
                                    If .fileNameLength = fileNameLen Then 'for an exact match, lengths of the filenames must be equal
                                        pFileNameOffset = m_SA.pvData + .FileNameOffset 'VarPtr(outBuff(i_plus_FileNameOffset))
                                        If caseMatch Then
                                            'comparememory returns the number of bytes that match until the first difference
                                            If CompareMemory(ByVal pFileNameOffset, fileNameBytes(NUM_0), fileNameLenLong) = fileNameLenLong Then
                                                LSet files(fileIndex).FileRef = .FileReferenceNumber.ID
                                                LSet files(fileIndex).ParentFileRef = .ParentFileReferenceNumber.ID
                                                SysReAllocStringLen VarPtr(files(fileIndex).fileName), pFileNameOffset, .fileNameLength \ NUM_2
                                                fileIndex = fileIndex + NUM_1 '<TODO> add error handling for too many files found
                                                If fileIndex > maxResults Then quitSearch = True: Exit Do 'for now, just quit searching
                                            End If
                                        Else
                                            'do a match for the exact text, but case insensitive
                                            'StrCmpNIW should be the fastest method to compare case-insensitive strings
                                            If StrCmpNIW(pFileNameBytes, pFileNameOffset, fileNameStrLen) = NUM_0 Then
                                                LSet files(fileIndex).FileRef = .FileReferenceNumber.ID
                                                LSet files(fileIndex).ParentFileRef = .ParentFileReferenceNumber.ID
                                                SysReAllocStringLen VarPtr(files(fileIndex).fileName), pFileNameOffset, .fileNameLength \ NUM_2
                                                fileIndex = fileIndex + NUM_1 '<TODO> add error handling for too many files found
                                                If fileIndex > maxResults Then quitSearch = True: Exit Do 'for now, just quit searching
                                            End If
                                        End If
                                    End If
                                Else
                                    If .fileNameLength >= fileNameLen Then
                                        'do a partial filename match
                                        pFileNameOffset = m_SA.pvData + .FileNameOffset 'VarPtr(outBuff(i_plus_FileNameOffset))
                                        'get filename into a temp string to convert to lowercase if needed for case insensitive search
                                        tmpFileNameChars = .fileNameLength \ NUM_2
                                        SysReAllocStringLen pTmpFileNameStr, pFileNameOffset, tmpFileNameChars
                                        'CharLowerBuff is faster than LCase$
                                        If caseMatch = False Then CharLowerBuff StrPtr(tmpFileNameStr), tmpFileNameChars
                                        
                                        'InStrB should be faster than InStr
                                        'InStrB performs better when comparing strings instead of byte arrays for some reason
                                        If InStrB(tmpFileNameStr, findFileName) > NUM_0 Then
                                            LSet files(fileIndex).FileRef = .FileReferenceNumber.ID
                                            LSet files(fileIndex).ParentFileRef = .ParentFileReferenceNumber.ID
                                            SysReAllocStringLen VarPtr(files(fileIndex).fileName), pFileNameOffset, tmpFileNameChars
                                            fileIndex = fileIndex + NUM_1 '<TODO> add error handling for too many files found
                                            If fileIndex > maxResults Then quitSearch = True: Exit Do 'for now, just quit searching
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        m_SA.pvData = m_SA.pvData + .recordLength 'set safearray pointer to the next record
                    End With
                Loop
                'totalMs = totalMs + (timer.TimeElapsed - curMs)
                
                If quitSearch Then Exit Do 'exit the loop if we want to quit searching
                'first 8 bytes in the output buffer contain the next StartReferenceFileNumber to iterate with
                CopyMemory enumData.StartFileReferenceNumber.Value(NUM_0), outBuff(NUM_0), NUM_8
                
                res = DeviceIoControl(drive, FSCTL_ENUM_USN_DATA, enumDataPtr, enumDataSize, outBuffPtr, outBuffSize, bytesAccessed)
            Loop
            
'            If quitSearch Then Debug.Print "---EARLY QUIT"
'            Debug.Print "bytesAccessed " & bytesAccessed
'            Debug.Print "folder count - " & folderIndex
            'Debug.Print "totalMs - " & totalMs
            
            'curMs = timer.TimeElapsed
            'sort the array of folders so we can do binary search instead of hashing
            QuickSort folders, 0, folderIndex - 1
            'Debug.Print "sort time - " & (timer.TimeElapsed - curMs)
            
            If fileIndex > 0 Then
                ReDim returnStr(fileIndex - 1)
            End If
            
            'curMs = timer.TimeElapsed
            'loop through file matches found, get their folder paths
            For y = 0 To fileIndex - 1
                returnStr(y) = driveLetter & ":" & BuildFolderPath(files(y), folders, folderIndex - 1)
            Next y
            'Debug.Print "build folder time - " & (timer.TimeElapsed - curMs)
            
            Debug.Print "total elapsed time - " & timer.TimeElapsed
            
        End If
    End If
    
    CloseHandle drive
    
    Find = returnStr
End Function

'fast sorting method which takes vArray as an input/output, sorting in order of low to high FileRef numbers so that we can use BinarySearch
Private Sub QuickSort(vArray() As TFolderRef, inLow As Long, inHi As Long)
  Dim middle  As TFolderRef
  Dim tmpSwap As TFolderRef
  Dim tmpLow  As Long
  Dim tmpHi   As Long
  
  Dim pivotDbl      As Double
  Dim vArrayDblLow  As Double
  Dim vArrayDblHigh As Double

  tmpLow = inLow
  tmpHi = inHi

  middle = vArray((inLow + inHi) \ NUM_2) 'note the slightly unusual "\" operator - this performs integer division, faster than "normal" floating point division and leaves out the remainder which we dont want anyway

  pivotDbl = middle.FileRef.Value
  vArrayDblLow = vArray(tmpLow).FileRef.Value
  vArrayDblHigh = vArray(tmpHi).FileRef.Value
  
  While (tmpLow <= tmpHi)
     While (vArrayDblLow < pivotDbl And tmpLow < inHi)
        tmpLow = tmpLow + NUM_1
        vArrayDblLow = vArray(tmpLow).FileRef.Value
     Wend
     
     While (pivotDbl < vArrayDblHigh And tmpHi > inLow)
        tmpHi = tmpHi - NUM_1
        vArrayDblHigh = vArray(tmpHi).FileRef.Value
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + NUM_1
        tmpHi = tmpHi - NUM_1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub

'extremely fast method for searching a sorted array
'search vArray for searchFileRef (we want to find a FileRefNum, or ParentFileRefNum), the final parameter (inHi) is the max index of vArray to search
'binary searches work by examining a sorted array and checking the value in the middle index of the array (UBound / 2)
'if the middle value is less than the desired value, calculate a new middle where the lowest array index is middle + 1
'if the middle value is greater than the desired value, calculate a new middle where the highest array index is middle - 1
'thus, for every iteration, we eliminate half of the array elements from our search, greatly reducing the total number of elements evaluated
Private Function BinarySearch(vArray() As TFolderRef, searchFileRef As Double, ByVal inHi As Long) As TFolderRef
    Dim middle As Long, lo As Long
    
    lo = LBound(vArray)
    middle = (lo + inHi) \ NUM_2
    
    While vArray(middle).FileRef.Value <> searchFileRef
        If vArray(middle).FileRef.Value > searchFileRef Then
            inHi = middle - NUM_1
        Else
            lo = middle + NUM_1
        End If
        middle = (lo + inHi) \ NUM_2
        If lo > inHi Then Exit Function 'not found
    Wend
    BinarySearch = vArray(middle)
End Function
'build the complete path to "file", by matching its ParentFileRef (folder where file is located) to vArray.FileRef (array of folders)
Private Function BuildFolderPath(file As TFolderRef, vArray() As TFolderRef, ByVal inHi As Long) As String
    Dim fileParentRef As TFolderRef, directoryNameStr As String, pathStr As String
    Dim fileName As String
    
    fileName = Replace$(file.fileName, Chr$(NUM_0), vbNullString) 'remove null characters from filename
    fileParentRef = BinarySearch(vArray, file.ParentFileRef.Value, inHi) 'get folder where file is located
    directoryNameStr = Replace$(fileParentRef.fileName, Chr$(NUM_0), vbNullString) 'remove null characters from folder name
    pathStr = directoryNameStr 'start building the path string
    
    'continue using BinarySearch to find the next parent folder, root directory is reached when there is no directory name found
    While Len(directoryNameStr) > NUM_0
        fileParentRef = BinarySearch(vArray, fileParentRef.ParentFileRef.Value, inHi)
        directoryNameStr = Replace$(fileParentRef.fileName, Chr$(NUM_0), vbNullString)
        pathStr = directoryNameStr & "\" & pathStr
    Wend
    BuildFolderPath = pathStr & "\" & fileName
End Function

Private Function LI2Double(ByVal lilo As Long, ByVal lihi As Long) As Double
    Dim Low As Double
    Low = lilo
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = lihi * TWO_32 + Low
End Function
