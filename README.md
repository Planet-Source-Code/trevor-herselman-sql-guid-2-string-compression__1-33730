<div align="center">

## SQL GUID 2 String \(Compression\)


</div>

### Description

2 functions to convert a 38 character GUID eg. {2E1EFAD8-4AD9-48A6-B9A9-75505F2B9A51} to a compact/compressed 16 character VB String eg. @#$%<|{.#&(@*^&! and back again. I use these 2 functions to efficiently convert literally millions of my SQL uniqueidentifier variables to their 16 character VB String equivalents for storage in arrays. This code is from my SQL & VB based game which stores millions of these GUID's, I needed a more efficient way to store them as my memory requirements are already running into gigs. The code has been slightly modified for PSC (no input/debug checks). If anyone finds it useful or has any kind of speed improvement or has questions, please let me know !!! If you find this code/concept good/useful, or don't understand, or have any performance suggestions (besides using API calls), just let me know. NB. With the help of Ian Webling's AWESOME advice and re-coding, this is the latest version using concepts from both of us, and has been heavily modified from the first release. MAJOR credit and thank you's must again be paid to Ian for his optimization skill !!! Ian collectively improved their speed by an estimated 70%, and I provided an additional 30% increase on his code, both calculated on 100,000 conversions. If U want MORE power/speed, check my API versions.
 
### More Info
 
String concatenation, using &, reduced my first version's speed by over 20% for 100,000 conversions. The iif statement has been replaced by a MUCH faster Format$ and Replace$ combination. Uses the Space$(16) VB Function to create a Buffer of spaces in GUID2ID function, 1.5% faster than using String$(16, 32) function. 99% of this code is attributed to the enhancements and suggestions made by Ian Webling.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Trevor Herselman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/trevor-herselman.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/trevor-herselman-sql-guid-2-string-compression__1-33730/archive/master.zip)





### Source Code

```
Public Function GUID2ID(ByRef InGUID As String) As String
 Dim Hex As String * 4
 Let GUID2ID = Space$(16)
 Let Hex = "&H00"
 Mid$(Hex, 3, 2) = Mid$(InGUID, 2, 2)
 Mid$(GUID2ID, 1, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 4, 2)
 Mid$(GUID2ID, 2, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 6, 2)
 Mid$(GUID2ID, 3, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 8, 2)
 Mid$(GUID2ID, 4, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 11, 2)
 Mid$(GUID2ID, 5, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 13, 2)
 Mid$(GUID2ID, 6, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 16, 2)
 Mid$(GUID2ID, 7, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 18, 2)
 Mid$(GUID2ID, 8, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 21, 2)
 Mid$(GUID2ID, 9, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 23, 2)
 Mid$(GUID2ID, 10, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 26, 2)
 Mid$(GUID2ID, 11, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 28, 2)
 Mid$(GUID2ID, 12, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 30, 2)
 Mid$(GUID2ID, 13, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 32, 2)
 Mid$(GUID2ID, 14, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 34, 2)
 Mid$(GUID2ID, 15, 1) = Chr$(CLng(Hex))
 Mid$(Hex, 3, 2) = Mid$(InGUID, 36, 2)
 Mid$(GUID2ID, 16, 1) = Chr$(CLng(Hex))
End Function
Public Function ID2GUID(ByRef InID As String) As String
 Let ID2GUID = "{12345678-1234-1234-1234-123456789012}"
 Mid$(ID2GUID, 2, 2) = Format$(Hex$(Asc(InID)), "@@")
 Mid$(ID2GUID, 4, 2) = Format$(Hex$(Asc(Right$(InID, 15))), "@@")
 Mid$(ID2GUID, 6, 2) = Format$(Hex$(Asc(Right$(InID, 14))), "@@")
 Mid$(ID2GUID, 8, 2) = Format$(Hex$(Asc(Right$(InID, 13))), "@@")
 Mid$(ID2GUID, 11, 2) = Format$(Hex$(Asc(Right$(InID, 12))), "@@")
 Mid$(ID2GUID, 13, 2) = Format$(Hex$(Asc(Right$(InID, 11))), "@@")
 Mid$(ID2GUID, 16, 2) = Format$(Hex$(Asc(Right$(InID, 10))), "@@")
 Mid$(ID2GUID, 18, 2) = Format$(Hex$(Asc(Right$(InID, 9))), "@@")
 Mid$(ID2GUID, 21, 2) = Format$(Hex$(Asc(Right$(InID, 8))), "@@")
 Mid$(ID2GUID, 23, 2) = Format$(Hex$(Asc(Right$(InID, 7))), "@@")
 Mid$(ID2GUID, 26, 2) = Format$(Hex$(Asc(Right$(InID, 6))), "@@")
 Mid$(ID2GUID, 28, 2) = Format$(Hex$(Asc(Right$(InID, 5))), "@@")
 Mid$(ID2GUID, 30, 2) = Format$(Hex$(Asc(Right$(InID, 4))), "@@")
 Mid$(ID2GUID, 32, 2) = Format$(Hex$(Asc(Right$(InID, 3))), "@@")
 Mid$(ID2GUID, 34, 2) = Format$(Hex$(Asc(Right$(InID, 2))), "@@")
 Mid$(ID2GUID, 36, 2) = Format$(Hex$(Asc(Right$(InID, 1))), "@@")
 Let ID2GUID = Replace$(ID2GUID, " ", "0")
End Function
```

