# VB.Net Native Copy

A simple way to use native windows copy functionality to copy multiple files and/or folders in VB.Net.

```vb
Dim copy_items_paths As List(Of String)
Dim target_path      As String

NativeCopy.Copy(copy_items_paths, target_path)
```
