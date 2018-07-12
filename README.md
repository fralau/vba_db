# vba_db
Bringing higher-level concepts to VBA Access

# Introduction

There are much more advanced languages than VBA (e.g. Python), but here it is, you have to work with Ms-Access. Is everything lost? No!
The ecosystem is actually quite rich. It's only that it's not immediately accessible and not many people really bothered to work things
out in a way that they could be used by others.

There are actually things that you can do, with this set of modules:

1. **Databases**: Use rows from tables/queries as dictionaries (as everyone in their right mind would do today); easily establish connections and compose queries for Access, SQLServer,Oracle, and MySQL!
2. **RegExp2**: Using regexps in your application easily.
2. **Parameter**: How about a little parameter table for your app, with all whistles and bells?

## How to Install

Nothing fancy here:
1. Clone the repository or download the modules (`Database.bas`, `Parameter.bas`and `RegExp2.bas`)
2. In Ms-Access, open the Visual Basic Editor (Alt-F11)
3. From File menu, choose Import and select the file you want import
4. Set up the libraries you need (they are indicated in the header of your module): from the Tools menu, choose References to display the
   References dialog box (see more information on [Microsoft Dev Center](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/check-or-add-an-object-library-reference).
