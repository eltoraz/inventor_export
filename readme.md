This is a module of Visual Basic rules written for [Built by Newport](http://builtbynewport.com/) to import parts from Autodesk Inventor into Epicor ERP.
I tried to keep the business-specific logic separate from the Inventor API stuff for portability.

Tested under Inventor 2016. This workflow supports part and assembly documents.
Others may work too if they access Parameters via `Document.ComponentDefinition.Parameters.UserParameters`.

Configure Inventor to check the right directories to pull in all these external rules:
- Tools > Options > iLogic Configuration (may be in the dropdown rather than the ribbon)
- Add the Epicor, shared, and Species directories separately under External Rule Directories (adding the root git directory will make Inventor try to pull in everything, including the `.git` folder, which will cause the whole operation to fail)

Place XML files in forms/ under Inventor's `Design Data\iLogic\UI` directory
