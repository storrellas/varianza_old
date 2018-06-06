# varianza

What we need is a Python function that receives an openpyxl worksheet and modifies its formatting, according to additional parameters such as:

* Column width adjust to content (Default True)
* Font for all cells (Default None, meaning no change)
* Header row (Default first row)
* Header background color (Default None)
* Odd rows background color (Default None)
* Even rows background color (Default None)
* Add thousands separator to number cells (Default True)

Given the number of parameters, maybe it's simpler if we use a dictionary to hold them, but it's something we can look into later.

Attached to this message you can find a reference input and output, where you can see that the function would change the background color of the header and even rows, change the font to Calibri Light, adjust column width to contents and add thousands separator.
