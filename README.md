## Multilanguage file validator

This program is intended to make it easier to find faults
in customer provided translation excel files.

This is written in Python and makes use of the openpyxl
module, which can be installed via pip.  Rather than
have to count on users having Python and openpyxl
installed, it would be better after it's working to have
it made into a standalone executable.

## Functionality

Ideally, the program could be run on a command line,
accepting the names of a clean file pulled from backoffice
and the customer-provided file with errors as arguments.
It will create a separate output file that can then be
imported into backoffice.

Right now I have two separate parts, the first of which
just informs if there are a different number of tabs in
each file, and reports which one is extra.

The second part is supposed to correct faulty column
header/tab names HOWEVER this should only be run if the
tabs/columns are confirmed to be in the same order in both
files.

## Testing

I'm using an editor with almost no content to test,
camacho1@walkme.com

I tried opening a freshly exported file and saving it with
openpyxl without making changes and couldn't do it....when
I copy and paste each sheet from the exported file into the
output, then it works, so it seems that openpyxl is
changing something with the file, presumably formatting.