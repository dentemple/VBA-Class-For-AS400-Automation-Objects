# VBA Wrappers For AS400 Automation Objects

A self-contained class file for use in your Access or Excel projects.  This file contains late-binded class methods that allows you&#8212;the power user&#8212;to control an iSeries (AS400) client-side interface.

This file simply refactors IBM's Host Access Class Library (HACL) for readability and ease-of-use.

Not every single method has been included, just many of the ones I've tested and used live.  For the complete exhaustive list, visit the original documentation here (which is current as of October 2016):

https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/host_access08.htm

## Getting Started

Just download the zip, extract the files, and then upload the main one ("cAS400.cls") into your VBA project like any other module.  This VBA code can also be copy/paste'd directly into an existing VBA class file: just remove the header information first (i.e., the lines at the top from "VERSION..." to "ATTRIBUTE...")

More detailed instructions are listed in the Deployment section below.

An optional module has also been included that provides example usage and additional testing information.

### IMPORTANT NOTICE

This file assumes that your VBA code can access the iSeries directly, such as through a desktop installation or a remote-in option.

If an additional layer exists between you and the iSeries&#8212;for example, if you are on a thin-client, or if your company utilizes a web-interface to access the AS400&#8212;then additional steps may need to be taken to successfully utilize the HACL objects.

This may be as simple as early-binding the HACL or flipping over a single network setting; or, it can be as complicated as integrating this code with an automated web-scraper built from scratch with JavaScript.

(Because heavens forbid we actually change anything on the iSeries' backend to make everyone's lives easier).

Troubleshooting these situations, however, are beyond my current scope.  Still, I invite the community at large to copy or extend these files with their own solutions regarding these issues.

## Optional Pre-Download Test

## Example Usage

## Code-specific notes

### Regarding use of the Long data type

When handling client-side I/O, the iSeries automation objects prefers data type Long when handling numerical values.

To remain consistent, this class file does the same, even in instances where an integer value would be more intuitive.

### Regarding booleans in MS Access

IF YOU ARE USING MS ACCESS, please remember that  MS Access `True` may not always be the same `True` as found in other programs.

I've lost quite a few frustrating hours to Access's ability to inadvertently evaluate `True = True` situations to `False`.

Therefore, remember that these methods will pass "truthiness" directly from the AS400--and that any subsequent comparisons made in Access should be done defensively.

### Regarding duplicate HACL methods

## Deployment

## Author

Den Temple | dentemple.io

## License

This file is a refactor of IBM's Host Access Class Library (HACL). IBM provides documentation regarding the HACL here:

https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/host_access08.htm

Use of the HACL itself should be deferred to IBM's and Microsoft's usage guidelines.  

Any transformative code that does not necessarily need to fall under these aforementioned guidelines is licensed and freely available under the [License].
