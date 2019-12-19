# SaveMail

This plugin allows users to select and save multiple e-mails using a pre-defined naming convention of the following format:
`yyyy/MM/dd sender subject.msg`

Once the save operation is successful the saved e-mails will also be moved into a new folder in the Inbox called 'Saved Mail'. This allows users to keep track of which e-mails have been exported and helps keep a clean inbox.

Installation is achieved through a MicroSoft Installer. All the dependencies are included within the setup package. The plugin is placed under `./Program Files/OPW` and the appropriate registry entries are also created. Note that for 32-bit versions of Outlook a modification is required for Outlook to pick up the plugin at runtime.

Development has been approached in a test driven manner. The design follows a conventional model-view-controller pattern. All the methods in the controller class `SaveMail` are tested within the `UnitTestsForSaveMail` project. The `SetupPackage` takes the output and builds a single installer file.
