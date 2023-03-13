# EXCEL_VBA_FOR_AUTOMATION

Excel VBA with WIN32 API access may be a superb tool for programing automation of windows 32 programs.
Here I leave some routines that ca be helpfull.

Use Mirosoft Spy++ to find the Windows Tree

Find again your objects handle everytime you open the corresponding window.


<p>h1 = FindWindowMul(h1, "TPanel", 1)</p>
<p>h1 = FindWindowMul(h1, "TPageControl", 3)</p>
<p>h1 = FindWindowMul(h1, "TTabSheet", 2)</p>

Then use the helper functions to read and write text, activate buttons or checkboxes.

