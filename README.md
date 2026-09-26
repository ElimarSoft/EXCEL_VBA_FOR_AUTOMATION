# EXCEL_VBA_FOR_AUTOMATION

Excel VBA with WIN32 API access may be a superb tool for programing automation of windows 32 programs.
Here I leave some routines that ca be helpfull.

Use Mirosoft Spy++ to find the Windows Tree

Find again your objects handle everytime you open the corresponding window.


<p>h1 = FindWindowMul(h1, "TPanel", 1)</p>
<p>h1 = FindWindowMul(h1, "TPageControl", 3)</p>
<p>h1 = FindWindowMul(h1, "TTabSheet", 2)</p>

Then use the helper functions to read and write text, activate buttons or checkboxes.

# DISCLAIMER

This VBA code is provided "AS IS" without any warranty.
The author assumes no responsibility for any errors, data loss,
file corruption, business interruption, or any other damages
resulting from the use of this code.
Users are responsible for testing and validating the code before
using it in production environments.
Use at your own risk.

Copyright © 2026 elimar.com
