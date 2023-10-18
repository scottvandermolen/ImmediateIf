<!--#include file="iif.function.asp"-->
<%
	' Immediate If Function Unit Tests
	'
	' Copyright (c) 2023, Scott Vander Molen; some rights reserved.
	' 
	' This work is licensed under a Creative Commons Attribution 4.0 International License.
	' To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/
	' 
	' @author  Scott Vander Molen
	' @version 1.0
	' @since   2008-04-26
	
	' Framework for running tests and building result strings.
	' 
	' @param input The specified value.
	' @param expected The expected value.
	' @param actual The actual value.
	' @return string The results of the test.
	function testFramework(input, expected, actual)
		dim result
		dim resultText
		
		if actual = expected then
			result = true
			resultText = "successful"
		else
			result = false
			resultText = "failed"
		end if
		
		dim returnString
		returnString = "Input:     " & input & vbCrLf & _
			"Expected:  " & expected & vbCrLf & _
			"Actual:    " & actual & vbCrLf & _
			"Result:    Test " & resultText &  "!" & vbCrLf & vbCrLf
		
		testFramework = returnString
	end function
	
	' Create an HTML container for our output.
	Response.Write "<!DOCTYPE html>" & vbCrLf
	Response.Write "<html lang=""en"">" & vbCrLf
	Response.Write "<meta http-equiv=""Content-Type"" content=""text/html;charset=UTF-8"" />" & vbCrLf
	Response.Write "<body>" & vbCrLf
	
	' Display code header
	Response.Write "<pre>"
	Response.Write "/***************************************************************************************\" & vbCrLf
	Response.Write "| ASP Immediate If Function Unit Tests                                                  |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "| Copyright (c) 2023, Scott Vander Molen; some rights reserved.                         |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "| This work is licensed under a Creative Commons Attribution 4.0 International License. |" & vbCrLf
	Response.Write "| To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/    |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "\***************************************************************************************/" & vbCrLf
	Response.Write "</pre>"
	
	' Run unit tests
	Response.Write "<pre>" & vbCrLf
	Response.Write "Unit Test: 2 + 2 = 5" & vbCrLf & testFramework(2 + 2 = 5, "Wrong", IIf(2 + 2 = 5, "Correct", "Wrong"))
	Response.Write "Unit Test: 2 + 2 = 4" & vbCrLf & testFramework(2 + 2 = 4, "Correct", IIf(2 + 2 = 4, "Correct", "Wrong"))
	Response.Write "</pre>" & vbCrLf
	
	' Close the HTML container.
	Response.Write "</body>" & vbCrLf
	Response.Write "</html>"
%>