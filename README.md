# Immediate If Function for Classic ASP

Visual Basic and Visual Basic for Applications both include an Immediate If function (IIF) which returns the second or third parameter based on evaluation of the first parameter. This function is missing from VBScript, but we can add it back.

## Project Status

No further development is planned, as this function is considered complete.

## Installation

Place iif.function.asp in any location on your web server, or on another machine on the same network. For additional security, you may wish to place it in a location that isn't directly accessible by users.

The file iif.test.asp is not required in order to use the library and does not need to be placed on the web server unless you want to run unit tests.

## Usage

```vbscript
<!--#include file="iif.function.asp"-->
<%
' This will display "Wrong"
Response.Write IIf(2 + 2 = 5, "Correct", "Wrong")

' This will display "Correct"
Response.Write IIf(2 + 2 = 4, "Correct", "Wrong")
%>
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

See iif.test.asp for unit tests.

## Authors

Version 1.0 written April 2008 by Scott Vander Molen

## License
This work is licensed under a [Creative Commons Attribution 4.0 International License](https://creativecommons.org/licenses/by/4.0/).