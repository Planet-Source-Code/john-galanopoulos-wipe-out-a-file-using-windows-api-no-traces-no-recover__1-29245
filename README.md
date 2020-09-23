<div align="center">

## Wipe out a file using Windows API\. No traces, no recover\.


</div>

### Description

Completely destroy a file with no chance of recovery or trace. Use of CreateFile,

FILE_FLAG_NO_BUFFERING

(Open the file with no intermediate buffering or caching)

FILE_FLAG_WRITE_THROUGH

(Write through any intermediate cache and go directly to disk)

and also FlushFileBuffers function to ensure that file buffers will be flushed!

A must test and see.
 
### More Info
 
The File you wish to delete

The results for your action

Handle with care. No way to restore the file after deletion.


<span>             |<span>
---                |---
**Submitted On**   |2001-11-27 02:45:58
**By**             |[John Galanopoulos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-galanopoulos.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Wipe\_out\_a3765311262001\.zip](https://github.com/Planet-Source-Code/john-galanopoulos-wipe-out-a-file-using-windows-api-no-traces-no-recover__1-29245/archive/master.zip)

### API Declarations

Please use this version of my program.





