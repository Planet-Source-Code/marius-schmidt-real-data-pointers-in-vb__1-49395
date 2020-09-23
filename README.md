<div align="center">

## REAL data pointers in VB


</div>

### Description

This code shows you how to use Real pointers with VB. Everyone who jelled that VB has no pointers is simply wrong! VB arrays are just pointers to memory and it's quite simple to change the address of these pointers. It's pretty usefull for fast memory access without using a crappy CopyMemory call (the CopyMemory way would a) create memory garbage and b) create overhead because data moving is slow [though cpymemory is faster then manually moving data]), though we still need 1 small CopyMemory call but it's just a helper.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-10-22 16:07:02
**By**             |[Marius Schmidt](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marius-schmidt.md)
**Level**          |Intermediate
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[REAL\_data\_16620910222003\.zip](https://github.com/Planet-Source-Code/marius-schmidt-real-data-pointers-in-vb__1-49395/archive/master.zip)

### API Declarations

```
CopyMemory (well, it's just used for a little data moving)
VarPtrArray
```





