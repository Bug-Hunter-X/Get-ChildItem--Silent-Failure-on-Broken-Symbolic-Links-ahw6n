# VBScript Get-ChildItem Broken Symbolic Link Issue

This repository demonstrates a subtle issue with VBScript's `Get-ChildItem` function when dealing with broken symbolic links. The script `brokenSymLink.vbs` attempts to access file information through a broken symbolic link, resulting in silent failure without raising errors. This can be problematic for error handling and validation within VBScript applications.

The `fixedSymLink.vbs` provides a solution that checks for the existence of the target file before attempting to retrieve information, thus ensuring that any errors are caught and handled appropriately.