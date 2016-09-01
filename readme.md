This is a suite of Visual Basic rules written for [Built by Newport](http://builtbynewport.com/) originally just to import parts from Autodesk Inventor into Epicor ERP.
It's since expanded from the original spec to include exporting part/assembly bills of materials from Inventor, along with population of spreadsheets with data for quoting.
I've tried to keep the business-specific logic separate from the Inventor API stuff for portability, but many of the details (especially the Inventor parameters used) are dictated by BBN operating procedure.

Just a warning: most of these modules probably won't compile under Visual Studio (etc.), since they use Inventor's [iLogic syntax and API](https://knowledge.autodesk.com/support/inventor-products/learn-explore/caas/CloudHelp/cloudhelp/2014/ENU/Inventor/files/GUID-B98DF82D-E489-4C19-8351-11C0ED63C349-htm.html).
They were tested and run on Inventor 2016.

This workflow supports part and assembly documents.
Others may work too if they access Parameters via `Document.ComponentDefinition.Parameters.UserParameters`.
BOM code paths are only defined for parts and assemblies though.

Configure Inventor to check the right directories to pull in all these external rules:
- Tools > Options > iLogic Configuration (may be in the dropdown rather than the ribbon)
- Add the `Epicor`, `shared`, `Species`, `Quoting`, and `BOM` directories separately under External Rule Directories (don't add the root git directory as a shortcut, or Inventor will unsuccessfully try to pull in everything, including the hundreds of objects in the `.git` folder!)

Place XML files in forms/ under Inventor's `Design Data\iLogic\UI` directory.
Inventor will generate `*.state.xml` files when the forms are opened for the first time.
These don't need to be tracked, since they store form dimensions/positioning and are thus prone to frequent change.
