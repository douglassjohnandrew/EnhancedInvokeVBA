# EnhancedInvokeVBA
This library is an enhancement of the existing Invoke VBA activity, with the following improvements:
- Creates and invokes an enhanced version of your VBA code file that includes error handling and line numbers
- If a VBA error occurs, returns detailed error information to UiPath, including the exact line that failed

CreateWorkbookVariable arguments: (required arguments marked by :triangular_flag_on_post:)
- **CreateNew** - If the ExcelFilePath does not exist, then True - Create the Excel file, or False - Throw an error
- **ExcelFilePath** :triangular_flag_on_post: - Full path to the Excel file
- **OpenReadOnly** - If True, then open the Excel file in read-only mode. If False, then open the Excel file normally
- **Visible** - If True, then the Excel file will be visible. If False, then the Excel file is not shown (recommended)
- **WorkbookVariable** :triangular_flag_on_post: - The WorkbookApplication variable for later use

EnhancedInvokeVBA arguments: (required arguments marked by :triangular_flag_on_post:)
- **CloseWorkbook** - If True, then close the workbook after VBA finishes. If False, leave the workbook open
- **CodeFilePath** :triangular_flag_on_post: - Full path to the file containing VBA code
- **EntryMethodParameterDefs** - Comma-separated string of entry method parameter definitions. Example: "name As String, age As Integer"
- **EntryMethodParameterValues** - Collection of values to be passed as entry method parameters. Example: {"Paul Smith", 37}
- **KeepEnhancedCodeFile** - If True, then keep the enhanced code file that the bot creates. If False, then delete the enhanced code file after it has been used
- **OutputValue** - The entry method's return value (optional)
- **SaveChanges** - If True, then save changes after VBA finishes. If False, do not save changes
- **WorkbookVariable** :triangular_flag_on_post: - An existing WorkbookApplication variable where the VBA will run. This variable can be created using the CreateWorkbookVariable activity

Other notes about this library:
- CodeFilePath and ExcelFilePath need to be **full** paths (using Directory.GetCurrentDirectory)
- If a boolean argument is left blank, then it will default to False
- The Excel file **cannot** be password protected
- This library **cannot** be used inside an Excel Application Scope

How to write your VBA code files for this library:
- Make sure that your code file is in a plain text format (using **.vb** is recommended)
- Ensure that your "main" code is **not** inside a Function or Sub
- Create helper functions and subs if needed

As a simple example, suppose you want to pass two integers from UiPath to VBA and display their sum in a VBA message box.

You could create the following code file **(MessageBoxDemo.vb)**

```vb
' A helper sub that displays the sum of two numbers in a VBA message box
Sub DisplaySum(n1 As Integer, n2 As Integer)
Msgbox("The sum of " & n1 & " and " & n2 & " is " & (n1+n2))
End Sub

' The main code calls DisplaySum two times. Note that this main code is NOT inside a Function or Sub.
' num1 and num2 are Integer variables that store the two integers passed from UiPath
Call DisplaySum(num1, num2)
Call DisplaySum(num1 - 1, num2 - 1)
```

First set the EntryMethodParameterDefs argument to be **"num1 As Integer, num2 As Integer"**

Then set the EntryMethodParameterValues argument to be two integers, e.g. **{5, 3}**

Finally, the library will create an enhanced code file **(MessageBoxDemo-Enhanced.vb)** that looks like this:

```vb
' Helper functions and subs are moved to the top of this file, outside the Main function
Sub DisplaySum(n1 As Integer, n2 As Integer)
Msgbox("The sum of " & n1 & " and " & n2 & " is " & (n1+n2))
End Sub

' This Main function is created by the library, and then executed in UiPath
Function Main(num1 As Integer, num2 As Integer)

' If the main code throws an error, go to Handle
On Error GoTo Handle

' Main code is placed here, with line numbers automatically added
1 Call DisplaySum(num1, num2)
2 Call DisplaySum(num1 - 1, num2 - 1)

' If the main code succeeded, then stop here
Exit Function

' If the main code failed somewhere, handle the error here
Handle:

' Retrieve the VBA error information and return it to UiPath
Dim errorArr(4) As String
errorArr(0) = "Error Occurred"
errorArr(1) = Erl
errorArr(2) = Err.Description
errorArr(3) = Err.Number
errorArr(4) = Err.Source
Main = errorArr

End Function
```

After this enhanced code file is created, UiPath executes the enhanced Main function within the specified Excel file. If a VBA error occurs, UiPath receives the error information and throws a system exception.
