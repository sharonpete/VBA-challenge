# VBA-challenge
The VBA of Wall Street


wrote out pseudo code
FUNCRES.XLAM blew up and I had to start over... don't use FUNCRES.XLAM, unless you are a masochist.  Or unless you know how to use it.

Got stuck on trying to find the rowcount... really stuck, for way too long.  It turns out that when you use
number_rows = Cells(Rows.Count, 1).End(xlUp).Row
that little section with "End(xlUp)" is using a lower-case "L" and not the number "1" 
