# excel_vba_funcWorksheetExists
Excel vba procedure to see if a worksheet exists without going through the native procedure.

funcWorksheetExists
Returns a boolean value.
=================================================================

This is for Excel 2010+, though may work in earlier versions.

I needed a simple way, in my VBA code, to tell if a worksheet was present in the workbook.
Testing the worksheet itself would cause problems if it wasn't there.

So, being an old Perl guy, I wanted a snippet that would allow what I was writing to handle
that exception elegantly instead of being forced to deal with the Microsoft way.

This code looks at all the names of the worksheets in the Worksheet object and if the name
matches, it returns a boolean TRUE. If it does not match, then the worksheet name you
submitted does not exist so it returns a boolean FALSE.

The *.bas module contains a subroutine to test the procedure.

To use this code, copy and paste it into your project OR import the file worksheet_exists.bas
