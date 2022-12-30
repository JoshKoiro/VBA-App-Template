# VBA-App-Template
Template to create applications from VBA in Excel

## Purpose

This codebase is designed to be used as a starting point for creating mock stand-alone applications in Excel VBA. It bundles classes and modules to quickly develop data-driven applications in excel while also automatically importing and exporting code files for proper version control under src\VBAProjectFiles

## Contributing to this Repository

Feel free to make changes to the files under the src folder using the editor of your choice. If you are using Excel's built-in VBA code editor, then make sure to save the excel file once changes are made. This repository does not keep changes made to Excel Application Template.xlsm so any changes made to that file will not be accepted.

If changes are made to the ManageCode.bas file inside Excel Application Template.xlsm, then the .gitignore file must be removed and a manual merge must happen before updating the repository

## Structure and Classes

### FormGen Class

The first set of classes are associated with creating a form programatically. This is done by constructing a FormGen class and then using the FormGen.build() method to constuct a pre-defined, formatted UserForm object.

The GenForm class can either be constructed through property definition, or though the import of data from a .csv file using the FormGen.import() method.

### UserInput Class

The UserInput Class is a class that defines the Form components that make up the FormGen Class items collection. Instances of this class represent UserForm input components that are configured through the class properties.

This class extends the UserForm Component object.

