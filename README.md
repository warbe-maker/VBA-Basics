# Common VBA Basic Services
# Services
| Service | Kind | Description |
|---------|------|-------------|
|_AppErr_ | Sub  | Ensures that the number of a programmed _Application-Error_ never conflicts with a _VB-Runtime-Error_ by adding the constant _vbObjectError_ which turns it into a negative value. In return, it translates the negative _Application-Error_ number back into its original positive number. When used in an error handling in conjunction with the unique identification of the error source (service ErrSrc in the mErH module) programmed error numbers can start from 1 to n in any procedure.|
|_AppInstalled_| Function | Returns TRUE when a named program/application, usually an .exe file is installed.|
|_ArrayCompare_| Function | Returns a Dictionary with the provided number of lines/items (defaults to all) which differ between two one dimensional arrays. When no differnece is encountered the returned Dictionary is empty (Count = 0).|
|_ArrayDiffers_| |Returns TRUE when two arrays are different (stops with the first difference).|
|_ArrayIsAlocated_|||
|ArrayNumberOfDims|||
|ArrayRemoveItems|||
|ArrayToRange|||
|ArrayTrim|||
|BaseName||Returns the base name of a File or Workbook object or a string|
|Center|||
|CleanTrim|||
|ElementOfIndex|||
|ErrMsg|||
|Max|||
|Min|||
|PointsPerPixel|||
|SelectFolder|||
|Spaced||Returns a provided string 'letterspaced' with non-braking spaces. Spaces already in the provided string are doubled.|


