<%
	' Immediate If Function
	'
	' Copyright (c) 2023, Scott Vander Molen; some rights reserved.
	' 
	' This work is licensed under a Creative Commons Attribution 4.0 International License.
	' To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/
	' 
	' @author  Scott Vander Molen
	' @version 1.0
	' @since   2008-04-26
	'
	' https://en.wikipedia.org/wiki/IIf

	' The iif function from Visual Basic that is missing from VBScript.
	function IIf(expression, truecondition, falsecondition)
		if cbool(expression) then
			IIf = truecondition
		else
			IIf = falsecondition
		end if
	end function
%>