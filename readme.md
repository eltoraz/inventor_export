Configure Inventor to check the right directories to pull in all these external rules:
- Tools > Options > iLogic Configuration (may be in the dropdown rather than the ribbon)
- Add the Epicor, shared, and Species directories separately under External Rule Directories (adding the root git directory will make Inventor try to pull in everything, including the .git folder, which will cause the whole operation to fail)

Place XML files in forms/ under Inventor's "Design Data\\iLogic\\UI" directory
