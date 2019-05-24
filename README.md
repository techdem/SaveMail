# SaveMail

This plugin allows users to select and save multiple e-mails using a pre-defined naming convention of the following format:
`yyyy/MM/dd sender subject.msg`

Installation is achieved through a Visual Studio Tools for Office Deployment Manifest. Any Microsoft .NET Framework dependency will also be installed if necessary.

Development has been approached in a test driven manner. The design follows a conventional model-view-controller pattern. All the methods in the controller class `SaveMail` are tested within the `UnitTestsForSaveMail` project.
