<div align="center">

## Loan Amortization Calculator


</div>

### Description

Calculates montly payments for a loan depending on loan amount, interest and length of loan.

Perfect for mortgages or any loan!
 
### More Info
 
1) Loan amount

2) Interest amount (%)

3) Length of loan (in years)

1) Monthly payment

2) Total interest paid


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rob Swiger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rob-swiger.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rob-swiger-loan-amortization-calculator__4-8151/archive/master.zip)

### API Declarations

This code is free to use and distribute for non-profit or commercial use. I would appreciate an email just to let me know how you are using it!


### Source Code

```
<%@LANGUAGE="VBSCRIPT"%>
<!--
 ####################################################################
 # AUTHOR: Rob Swiger
 # EMAIL : rswiger@infinidev.com
 # URL : http://www.infinidev.com
 #
 # AMORTIZATION CALCULATOR
 # Calculates loan payments over a period of time with principle
 # and APR. Includes form.
 #
 # COPYRIGHT AND DISTRIBUTION INFORMATION
 # This code is free to use and distribute as needed for non-profit
 # and for-profit uses. Please leave the above information in tact
 # when using this code. An email letting me know would be nice also
 # just to let me know how many people are using it.
 #
 # WARRANTY INFORMATION
 # There are no warranties with this code. Please test it on your
 # server before deploying live. The author, nor infinidev, llc,
 # will be held accountable for any problems that may occur.
 #
 # SUPPORT INFORMATION
 # I do not offer support for this code, but if you send an email
 # to the above email address, I will do my best to answer any
 # questions you may have about the code free of charge.
 #
 # VERSION INFORMATION
 # Version 1.0
 # January 13, 2003
 #
 # NOTES AND RAMBLINGS
 # One addition that might be helpful is to write Javascript code
 # to check the values of the fields as the information is added.
 # I did not include this information because I did not want to
 # limit the code browser-wise because of spotty Javascript support.
 ####################################################################
-->
<%
 dim intAmount 'Amount of loan; priciple.
 dim intInterest 'Amount of interest in %.
 dim intPayment 'Monthly payment amount.
 dim intLength 'Length of loan.
 dim intTotInterest 'Total interest paid over length of loan.
%>
<html>
<head>
 <title>infinidev.com - Amortization Calculator</title>
</head>
<body>
<%
 'Check to see if form has been submitted.
 if len(request.form("isSub")) > 0 then
 'Check if all fields have information entered
 if len(request.form("amount")) > 0 and len(request.form("interest")) > 0 and len(request.form("length")) > 0 then
	 'Populate amount variable with amount field from form.
	 intAmount = request.form("amount")
	 'Populate interest variable by using interest form field divided by
	 '100 to make decimal (7% / 100 = 0.07), then divide by 12 for monthly interest.
	 intInterest = (request.form("interest") / 100) / 12
	 'Populate length of loan variable by taking length from form field (in years)
	 'and dividing by 12 to get total months.
	 intLength = request.form("length") * 12
	 'Calculate payments and round off to 2 decimal places. ($100.00)
	 intPayment = round(intAmount * intInterest/(1 - (1 + intInterest) ^ (-intLength)),2)
	 'Covert integer to US currency format.
	 intPayment = formatcurrency(intPayment,2)
	 'Calculate total interest paid over length of loan.
	 intTotInterest = round((intPayment * intLength) - intAmount,2)
	 'Covert integer to US currency format.
	 intTotInterest = formatcurrency(intTotInterest,2)
	else
	 'Alert user that all boxes must be filled in.
	 response.write("ALL BOXES must be filled in!")
	end if
 end if
%>
<h1>Loan Amortization Calculator</h1>
<hr color="#000000">
<br>
<table align="center" style="border:1px solid #000000;padding-left:3px;padding-right:3px;padding-top:3px;padding-bottom:3px;">
 <form id="loan" name="loan" action="a_calculator.asp" method="post">
 <input name="isSub" type="hidden" value="yes">
 <tr>
  <td align="right" >Amount of Loan (in US dollars):</td>
  <td> <input name="amount" type="text" id="amount" value="<%=request.form("amount")%>"></td>
 </tr>
 <tr>
  <td align="right">Annual Interest Rate (%):</td>
  <td><input name="interest" type="text" id="interest" value="<%=request.form("interest")%>"></td>
 </tr>
 <tr>
  <td align="right">Term of Loan (in years):</td>
  <td><input name="length" type="text" id="length" value="<%=request.form("length")%>"></td>
 </tr>
 <tr>
  <td colspan="2" align="center" >
<input type="submit" name="Submit" value="Find Monthly Payment!"></td>
 </tr>
 </form>
</table>
<%
 if len(intPayment) > 0 then
 response.write("<br>")
 response.write("<table border='0' width='400' align='center'><tr><td align='center'><p><b>")
	response.write("Your monthly loan payment will be:<br>")
	response.write("<font color='#CC0000' size='5'>")
	response.write(intPayment & "</font></b></p>")
	response.write("<p><b>Your total interest will be:<br>")
	response.write("<font color='#CC0000' size='5'>")
	response.write(intTotInterest & "</font></b></p>")
	response.write("</td></tr></table>")
 end if
%>
</table>
<br>
<hr color="#000000">
<p align="center">Calculator provided free by:<br>
<a href="http://www.infinidev.com">infinidev, llc</a></p>
</body>
</html>
```

