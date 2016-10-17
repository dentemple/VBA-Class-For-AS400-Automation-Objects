# VBA Wrappers For AS400 Automation Objects

A self-contained class file for use in your Access or Excel projects.  This file contains late-binded class methods that allows you--the power user--to control an iSeries (AS400) client-side interface.

This file simply refactors IBM's Host Access Class Library (HACL) for readability and ease-of-use.

The original documentation can be found here (current as of October 2016):
https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/host_access08.htm

## Getting Started

Just download the zip, extract the files, and then upload the main one ("cAS400.cls") into your VBA project like any other module.  This VBA code can also be copy/paste'd directly into an existing VBA class file: just remove the header information first (i.e., the lines at the top from "VERSION..." to "ATTRIBUTE...")

More detailed instructions are listed in the Deployment section below.

An optional module has also been included that provides example usage and additional testing information

### IMPORTANT NOTICE

This file assumes both you and VBA can access the iSeries directly, such as through a desktop installation or a remote connection.

If an additional layer exists between you and the iSeries--for example, if you are on a thin-client, or if your company utilizes an iSeries web-interface--then additional steps may need to be taken to successfully utilize the HACL.

Troubleshooting these situations are beyond my current scope; however, I invite the community to copy or extend these files with their own solutions.

## Optional Pre-Download Test

## Example Usage

## Code-specific notes

### Regarding the Long data type

When handling client-side I/O, these iSeries automation objects prefer the Long data type when handling numerical values.

To remain consistent, this class file does the same, even in instances where an integer value would be more intuitive.

### Regarding booleans in MS Access

IF YOU ARE USING MS ACCESS, please remember that  MS Access `True` may not always be the same `True` as found in other programs.

I've lost quite a few frustrating hours to Access's ability to inadvertently evaluate `True = True` situations to `False`.

Therefore, remember that these methods will pass "truthiness" directly from the AS400--and that any subsequent  comparisons made in Access should be done defensively.

### Regarding duplicate HACL methods

## Deployment

## Author

Den Temple | dentemple.io

## License

This file is a refactor of IBM's Host Access Class Library (HACL). IBM provides documentation regarding the HACL here:
https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/host_access08.htm

Use of the HACL itself should be deferred to IBM's and Microsoft's usage guidelines.  

Any transformative code that does not necessarily need to fall under these aforementioned guidelines is licensed under the [License].
