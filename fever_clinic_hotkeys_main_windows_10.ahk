
	; ************************ DIRECTIONS FOR USE *************************

		; TO BE USED SEQUENTIALLY FROM CTRL^1 THROUGH ^4 FOR PATIENT REGISTRATION WITHIN THE COVID-19 FEVER CLINIC
		
		; CTRL+SHIFT+Q TO BE USED TO COPY DETAILS FROM PAS WHEN EMAIL REGISTRATION NOT PRESENT
		
			;ONLY TO BE USED IN CONJUNCTION WITH OUTLOOK AND FEVER CLINIC XXX.xls
			
									; CREATED BY B.HOLYLAND, B.SCHUURMAN & D.TONG
									
	
	; ************************ KNOWN ERRORS ****************************
	
	; ^+q - mobile pattern recognition
	; MobilePattern recognition
	; Scroll Lock must be OFF 	
	
		
	; ************************ WORK IN PROGRESS ****************************
	
	; Clicks through HTML data 
	; Titles as variables: setting the SVHM UniCare and Microsoft Excel sheets as variables

	; Can we control the IE parameters? Ie. is it possible to set favourites bar off and menu bar on? 
																			
									
	; ************************ SHORTCUT INDEX ****************************
	
	
	; ^1 - parse clipboard, create variables and search patient in PAS
	; ^2 - new patient registration 
	; ^3 - writes patient information into spreadsheet and finds UR in PAS 
	; ^4 - updates NOK details for new patient registration
	; ^0 - sets zoom to 100%
	; ^l - registration: unknown demographics from mobile field 
	; ^+q - transfer patient demographic data from PAS to Excel 

	; ^+f - formats mobile phone to 0400 111 222
	; ^+g - formats home phone to 03 9231 2211
	; ^+r - reloads the script
	
	; ^{Escape} - kill switch 
	
	
	; clickCoord(userID) 					- adjusts click coordinates for SC, ED, ES and SU staff (dependent on User profile selection when script loads)
	; relationshipFunc(NOKrelationship) 	- relationship conditionals
	; urFunc(ur) 							- UR number to spreadsheet from PAS
	; patientTypeFunc(ptType) 				- patient type function 
	; mobilePattern(mobileall) 				- mobile function
	; aboriginalityFunc(response) 			- aboriginality 
	; religionFunc(religion) 				- Religion function
	; maritalFunc(maritalStatus) 			- marital status 
	; NOKmobilePattern(NOKmobileall) 		- NOK mobile phone pattern detect 
	; dobMonthFunc(QdobMonth) 				- DOB Month function 
	; cobFunc(cob) 							- Country of birth for PAS registration
	; genderFunc(gender) 					- sex selection
	; titleFunc(title)		 				- title selection
	; partialRego() 						- partial registration function (!careful! - ^' is used as find and replace in Notepad++)
	; QmobilePattern(mobileallQ)			- Q mobile function for ^+q function 
	; startingPosition(userID)				- sets the starting position for patient search in ^1 (dependent on User profile selection when script loads)
	; startingPositionNOK(userID)			- shift-tabs back in ^4 to functions to find "update patient details" tab (dependent on User profile selection when script loads)
	
	
	; WINDOWS IN-BUILT SHORTCUTS - specific to certain programs 
	
	; ^a - select all
	; ^b - bold highlight text 
	; ^c - copy (text, images, etc.)
	; ^f - find text (pattern in text)
	; ^o - open function 
	; ^p - print function 
	; ^s - save function
	; ^v - paste function
	; ^w - shut tab in browser 
	; ^x - cut function
	; ^z - undo function 
									

	; ************* CHECK UNINITIALISED VARIABLES ****************
	

	; NOTE: these are commented as they create more problems than they solve. 
	
;#Warn UseUnsetGlobal 
;#Warn UseUnsetLocal
;#Warn LocalSameAsGlobal
;#NoEnv
									
									
	; ************************ WELCOME MESSAGE *************************


		; **************** USER INPUT - SET USER ID ******************

if GetKeyState("Scrolllock", "T")
	{ 
	MsgBox, Scroll Lock is on, please turn off and re-run script.
	Return
	}
else

		
InputBox, userID, User Profile,
(LTrim
	Please enter your user profile (as a single number)`.
	 
	Enter 1 for Outpatient clerk`.
	Enter 2 for Emergency clerk`.
	Enter 3 for Elective Surgery clerk`.
	Enter 4 for SuperUser clerk`.
)
, , , 210

if userID is integer
{
	
	if userID=1
	{
		MsgBox, ,Script Successfully Loaded,
		(LTrim
			You will be using the script for an OUTPATIENT clerk profile`.
		)
	}
	
	else if userID=2 
	{
		MsgBox, ,Script Successfully Loaded,
		(LTrim
			You will be using the script for an EMERGENCY clerk profile`.
		)
	}
	
	else if userID=3
	{
		MsgBox, ,Script Successfully Loaded,
		(LTrim
			You will be using the script for an ELECTIVE SURGERY clerk profile`.
		)
	}
	
	else if userID=4
	{
		MsgBox, ,Script Successfully Loaded,
		(LTrim
			You will be using the script for a SUPERUSER clerk profile`.
		)
	}
	
	else
	{
		MsgBox, ,Oops`! Something is Wrong`.,
		(LTrim
			Please ensure you enter a valid number from 1-4`.
			Please reload the script and try again`. 
		)
	}
	
}

else 
{
	MsgBox, ,Script Failed to Load,
	(LTrim
		The script failed to load`.
		Please enter a number from 1-4`. 
	)
}




	; ****************** WELCOME MESSAGE ************************
	
	
Msgbox, , Welcome to Fever Clinic Hotkeys,
(
1. Make sure you have Excel, PAS, and Outlook open
2. Make sure you have a favorites bar in PAS

If anything goes wrong, press ctrl+Escape.
), 10


	; ******************************************************************
	
	; ******************* TITLES AS VARIABLES *************************
	
	
	;webBrowserTitle := "SVHM UniCare"
	;myExcelSheet := "Fever Clinic XXX"
	
	; ******************************************************************


	; ************************ PRIMARY HOTKEYS *************************

	
	; PATIENT SEARCH FROM EMAIL REGO - COPY ALL DETAILS (ctrl+a THEN ctrl+c)
	
^1::
	;FIRSTNAME
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, First Name: 		; Gets the position of the beginning of the word First
StringGetPos, length, Clipboard, Date of Birth: 	; Gets the position of the D
startPos := startPos + 14 							; Makes the starting pos at the end of the word First Name: in the string
length := (length+1) - startPos 					; the number of characters between the end of the word and the beginning of the D char
if startPos >= 0
  firstname := SubStr(Clipboard, startPos, length) 	; Grabs text between end of word First Name: and right before D
  firstname2char := SubStr(Clipboard, startPos, 2)
  StringLower, firstname, firstname, T
  
	;LASTNAME
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Last Name: 
StringGetPos, length, Clipboard, First Name:
startPos := startPos + 13
length := (length+1) - startPos 
if startPos >= 0
  lastname := SubStr(Clipboard, startPos, length) 
  lastname2char := SubStr(Clipboard, startPos, 2)
  StringLower, lastname, lastname, T
  
  	;DOB
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Date of Birth: 
StringGetPos, length, Clipboard, Gender:
startPos := startPos + 15 
length := (length+1) - startPos 
if startPos >= 0
  dob := SubStr(Clipboard, startPos, length) 
  
  	;ADDRESS LINE 1
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Address Line 1: 
StringGetPos, length, Clipboard, Address Line 2: 
startPos := startPos + 18 
length := (length+1) - startPos 
if startPos >= 0
  address1 := SubStr(Clipboard, startPos, length) 
  StringLower, address1, address1, T
  address1 := RegExReplace(address1, "[#]")

  	;ADDRESS LINE 2
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Address Line 2: 
StringGetPos, length, Clipboard, Suburb: 
startPos := startPos + 18 
length := (length+1) - startPos 
if startPos >= 0
  address2 := SubStr(Clipboard, startPos, length) 
  StringLower, address2, address2, T
  address2 := RegExReplace(address2, "[#]")
 
   	;SUBURB
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Suburb: 
StringGetPos, length, Clipboard, Postcode: 
startPos := startPos + 10 
length := (length+1) - startPos 
if startPos >= 0
  suburb := SubStr(Clipboard, startPos, length) 
  StringLower, suburb, suburb, T
  
     	;MOBILE
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Mobile:
StringGetPos, length, Clipboard, Telephone - Business:
startPos := startPos + 9 
length := (length+1) - startPos 
if startPos >= 0
  mobile1 := SubStr(Clipboard, startPos, length) 

     	;MOBILE2
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Home:
StringGetPos, length, Clipboard, Telephone � Mobile:
startPos := startPos + 6 
length := (length+1) - startPos 
if startPos >= 0
  mobile2 := SubStr(Clipboard, startPos, length) 

     	;MOBILE3
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Business: 
StringGetPos, length, Clipboard, Marital 
startPos := startPos + 10 
length := (length+1) - startPos 
if startPos >= 0
  mobile3 := SubStr(Clipboard, startPos, length) 
  
       	;TITLE
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Title: 
StringGetPos, length, Clipboard, Last Name: 
startPos := startPos + 7 
length := (length+1) - startPos 
if startPos >= 0
  title := SubStr(Clipboard, startPos, length) 
  
  
       	;GENDER
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Gender:
StringGetPos, length, Clipboard, Address Line 1:
startPos := startPos + 9 
length := (length+1) - startPos 
if startPos >= 0
  gender := SubStr(Clipboard, startPos, length) 
  
         	;NOK FIRST NAME
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, First Name NOK:
StringGetPos, length, Clipboard, Last Name NOK:
startPos := startPos + 18 
length := (length+1) - startPos 
if startPos >= 0
  NOKfirstname := SubStr(Clipboard, startPos, length) 
  
           	;NOK LAST NAME
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Last Name NOK:
StringGetPos, length, Clipboard, Relationship NOK:
startPos := startPos + 17 
length := (length+1) - startPos 
if startPos >= 0
  NOKlastname := SubStr(Clipboard, startPos, length) 
  
            ;NOK RELATIONSHIP
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Relationship NOK:
StringGetPos, length, Clipboard, Telephone - Home NOK:
startPos := startPos + 18 
length := (length+1) - startPos 
if startPos >= 0
  NOKrelationship := SubStr(Clipboard, startPos, length) 
  
            ;NOK MOBILE
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Mobile NOK:
StringGetPos, length, Clipboard, Telephone - Business NOK:
startPos := startPos + 12 
length := (length+1) - startPos 
if startPos >= 0
  NOKmobile1 := SubStr(Clipboard, startPos, length) 

            ;NOK MOBILE2
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Home NOK:
StringGetPos, length, Clipboard, Telephone - Mobile NOK:
startPos := startPos + 10 
length := (length+1) - startPos 
if startPos >= 0
  NOKmobile2 := SubStr(Clipboard, startPos, length)
  
            ;NOK MOBILE3
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Business NOK:
startPos := startPos + 14 
length :=  13
if startPos >= 0
  NOKmobile3 := SubStr(Clipboard, startPos, length) 

			;MOBILE ALL
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Telephone - Home:
StringGetPos, length, Clipboard, Marital
startPos := startPos + 1
length := (length+1) - startPos
if startPos >= 0
  mobileall := SubStr(Clipboard, startPos, length) 
  
			;NOK MOBILE ALL
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Telephone - Home NOK:
StringGetPos, length, Clipboard, Business NOK:
startPos := startPos + 1
length := (length+30) - startPos
if startPos >= 0
  NOKmobileall := SubStr(Clipboard, startPos, length)

  			; ABORIGINALITY
  
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Islander?
StringGetPos, length, Clipboard, Telephone - Home:
startPos := startPos + 10
length := (length+1) - startPos
if startPos >= 0
  response := SubStr(Clipboard, startPos, length) 
  
			; MARITAL STATUS
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Marital Status:
StringGetPos, length, Clipboard, Religion:
startPos := startPos + 16
length := (length+1) - startPos
if startPos >= 0
  maritalStatus := SubStr(Clipboard, startPos, length) 
  
			; RELIGION
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Religion:
StringGetPos, length, Clipboard, Country of Birth:
startPos := startPos + 12
length := (length+1) - startPos
if startPos >= 0
  religion := SubStr(Clipboard, startPos, length) 
 
			
			; COB
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Country of Birth:
StringGetPos, length, Clipboard, Occupation
startPos := startPos + 20
length := (length+1) - startPos
if startPos >= 0
  cob := SubStr(Clipboard, startPos, length)
  
  			; Message Consent
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, s via SMS?
StringGetPos, length, Clipboard, UR or
startPos := startPos + 11
length := (length+1) - startPos
if startPos >= 0
  msgConsent := SubStr(Clipboard, startPos, length)
msgConsent = %msgConsent%

  			; Employee Number
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Employee number:
StringGetPos, length, Clipboard, Location base:
startPos := startPos + 17
length := (length+1) - startPos
if startPos >= 0
  empNum := SubStr(Clipboard, startPos, length) 					 ; Get Employee Number 
empNum = %empNum% 													 ; clean whitespace
if empNum is integer
	employeeNumber := empNum 
else
	employeeNumber := ""
	

			; Patient Type
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Presenting as:
StringGetPos, length, Clipboard, SVHM Department:
startPos := startPos + 15
length := (length+1) - startPos
if startPos >= 0
  ptType := SubStr(Clipboard, startPos, length)
patientType = %ptType% 												  ; clean whitespace



		; Postcode
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Postcode:
StringGetPos, length, Clipboard, Do you identify
startPos := startPos + 10
length := (length+1) - startPos
if startPos >= 0
  pstCode := SubStr(Clipboard, startPos, length)
postCode = %pstCode% 													; clean whitespace

postCodeMatch := RegExMatch(postCode, "\d{4}", postCode)

if postCode is integer
{
	postCode := postCode
}

else
{
	postCode := 3
}

		; Department
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, SVHM Department:
StringGetPos, length, Clipboard, SVHM Employee number:
startPos := startPos + 17
length := (length+1) - startPos
if startPos >= 0
  department := SubStr(Clipboard, startPos, length)


 		; GP Consent 
StringReplace, Clipboard, Clipboard, `r`n, , All
StringGetPos, startPos, Clipboard, Consent for GP
StringGetPos, length, Clipboard, GP Name:
startPos := startPos + 41
length := (length+1) - startPos
if startPos >= 0
  gpConsent := SubStr(Clipboard, startPos, length)
gpConsent := LTrim(gpConsent)
  
Sleep 1000

;clickCoord(userID)														;Patient Search Function

	;Integrated clickCoord() function in here to return out if the PAS window isn't active

if WinExist("SVHM UniCare")
{
    WinActivate 
}
WinWaitActive, SVHM UniCare, , 2
 
if ErrorLevel
{
    MsgBox, Script timed out. Please Re-run.
    Return
}
else
Send ^0 								; Sets zoom to 100%
Sleep 400

CoordMode, Pixel, Window
PixelSearch, colourx, coloury, 5, 5, 300, 330, 0x00006D, 0, Fast RGB   ; was 0x292992
{					
	If ErrorLevel = 1 	; If not Found	
        {
		MsgBox, Click coordinates failed
	}
					
}

if userID=1
{
	coordX := colourx+200
	coordY := coloury+40						
}

else if userID=2
{
	coordX := colourx+300
	coordY := coloury+40
}

else if userID=3
{
	coordX := colourx+450
	coordY := coloury+40
}

else if userID=4
{
	coordX := colourx+600
	coordY := coloury+40
}


else
{
	MsgBox, ,Error,
	(LTrim, 
		There is an error with the cursor position.
		The program will exit. 
	)
	Return
}


 CoordMode, Mouse, Window
 Send {Click, %coordX%, %coordY%} 			
 Send {down}{enter}




Sleep 1000
startingPosition(userID)
Sleep 500
SendInput %lastname2char%
SendInput {tab}
SendInput %firstname2char%
Send {tab 3}
Send %dob%
Sleep 400
Send {enter}


MsgBox, , Script Complete, New patient? Click New Registration and Run Ctrl+2. Otherwise update patient demographics, 10
Return



; ******************************************************************

; ********************** NEW REGISTRATION **************************	


^2::
titleFunc(title)								;TITLE FUNCTION

Sleep 1500
SendInput {backspace} 							; deletes existing letters in Surname field 
Sleep 500
SendInput %lastname%
Sleep 250
SendInput {tab}
SendInput {backspace} 							; deletes existing letters in Firstname field 
Sleep 500
SendInput %firstname%
Sleep 250
SendInput {tab 2}

genderFunc(gender)								;GENDER FUNCTION

Sleep 1500	
Send {tab 3}
Send %address1% %address2%
Sleep 500
Send {tab 2}
;Send %suburb%
Send m
Sleep 500
Send {tab 2} 									; was 1 tab 
Sleep 800	 									; extra sleep time 
Send %postCode%
Send {tab}
Sleep 2000

mobilePattern(mobileall)					; MOBILE FUNCTION

Sleep 500
Send {tab 2}

mobilePattern(mobileall)					; MOBILE FUNCTION

Sleep 500	
Send {tab 2}

maritalFunc(maritalStatus)					; MARITAL STATUS FUNCTION

Sleep 1000
Send {tab}
Sleep 500

religionFunc(religion)						; RELIGION FUNCTION

Sleep 500
Send {tab}
Sleep 500

cobFunc(cob) 								; COB 	

Sleep 500
Send {tab}

aboriginalityFunc(response)					; ABORIGINALITY FUNCTION

Sleep 500

partialRego()								; PARTIAL REGO FUNCTION

MsgBox, , Script Complete, Return to Patient Demographics and Run Ctrl+3, 10
Return


; **************************************************************

; **************** TRANSFER DETAILS TO SPREADSHEET *************


^3::

if WinExist("SVHM UniCare")
{
    WinActivate 
}
WinWaitActive, SVHM UniCare, , 2
 
if ErrorLevel
{
    MsgBox, Script timed out. Please Re-run.
    Return
}
else

	Send ^0
	Sleep 400
	
	CoordMode, Pixel, Window
	; coordinates to search within work at all window sizes, use Window Spy for pixel colour and append '0x' to front
	PixelSearch, coordX, coordY, 400, 125, 3350, 500, 0x0000DD, 0, Fast RGB  
	{					
		If ErrorLevel = 1 	; If not Found	
		{
			MsgBox, Click coordinates failed
			Return
		}
	}
	
	BlockInput, On	
	CoordMode, Mouse, Window
	Send {Click, %coordX%, %coordY%}
	Sleep 500
	Send {Click, 2}
	Sleep 500
	Send {control down}
	Send c
	Send {control up}
	Sleep 200
	BlockInput, Off
	urNumber := Clipboard


if WinExist("Fever Clinic XXX")
{
    WinActivate 
}
WinWaitActive, Fever Clinic XXX, , 2
if ErrorLevel
{
    MsgBox, Script timed out. Please Re-run.
    return
}
else

Sleep 500
Send {escape}
Sleep 200
Send {F5}'FEVER CLINIC Tracking'+1C1000{tab 3}{enter}
Sleep 500
Send {control down}
Send {down}
Send {control up}
Sleep 200
Send {down}
Sleep 300
Send {F2}
Send %firstname% %lastname%
Send {tab}
Send {F2}
Send %urNumber%
Send {tab}
Send {F2}
Send %employeeNumber%
Send {tab}
Send {F2}
Send %dob%
;Send {tab}
;Send {F2}
;Send %postCode%
Send {tab}
Send {F2}
Send %address1% %address2%, %suburb%
Send {tab 2}
Send {F2}
Send %msgConsent%											  ; write message consent 
Send {tab}
Send {F2}
Sleep 200

gpConsentFunc(gpConsent)									; write GP consent

Send {tab}
Send {F2}
Sleep 200


patientTypeFunc(patientType)								  ; write Patient Type 

Sleep 500
Send {tab 3}
Send {F2}

mobilePattern(mobileall)                                	  ; write mobile 

Sleep 1500
Send {tab 1}
Send {F2}

Send %department%											  ; write staff department

Sleep 500
Send {enter}
Sleep 500

;MsgBox, , Script Complete, Return to Patient Demographics and Run Ctrl+4 to update NOK if required`., 10

Return 


	; ENTER NEXT OF KIN DETAILS TO PAS


^4::
if WinExist("SVHM UniCare") 	
{
    WinActivate 
}
WinWaitActive, SVHM UniCare, , 2
if ErrorLevel
{
    MsgBox, Script timed out. Please Re-run.
    return
}
else
  
CoordMode, Pixel, Window
	;coordinates to search within work at all window sizes, use Window Spy for pixel colour and append '0x' to front
	PixelSearch, colourx, coloury, 1, 1, 300, 330, 0xDDD05D, 0, Fast RGB 	
	{					
		If ErrorLevel = 1 	; If not Found	
		{
			MsgBox, Click coordinates failed
			Return
		}
	}
	
    coordx := colourx+460
    coordy:= coloury+90

	BlockInput, On	
	CoordMode, Mouse, Window
	Send {Click, %coordx%, %coordy%}
	Sleep 500
	BlockInput, Off

Sleep 1000	
Send e
Sleep 500
Send {enter}
Sleep 1000
SendInput {CtrlDown}a                                            ; deletes old name
Sleep 500
SendInput {CtrlUp} 
SendInput {BackSpace} 

/* 
 Sleep 500
 Send ^f
 Sleep 500
 Send address
 Sleep 500
 Send {escape}
 Sleep 500
 startingPositionNOK(userID)												; SC: {tab}, ED: {tab 4}, EL: {tab}
 Sleep 500
 Send u																	; UPDATE PATIENT DETAILS
 Sleep 500
 Send {enter}
 Sleep 500
 Send ^f
 Sleep 500
 Send U/R
 Sleep 500
 Send {escape}
 Sleep 1000
 Send {shift down}
 Send {tab 11}
 Send {shift up}
 Sleep 1000
 Send {enter}													; EMERGENCY CONTACT
 Sleep 1500
*/

Send %NOKfirstname% %NOKlastname%
Sleep 500
Send {tab 6}

relationshipFunc(NOKrelationship)								; NOK relationships 

Sleep 1500
Send {tab}

NOKmobilePattern(NOKmobileall)

Sleep 1500
Send {tab 2}

NOKmobilePattern(NOKmobileall)									; NOK mobile pattern and detect 

Sleep 1500
Send {tab}

MsgBox, , Script Complete, Check NOK details. If they are accurate then please update and complete any other patient information`., 10

Return



	; **************************************************************

	;************* FEVER CLINIC MANUAL REGISTRATION/AFTERHOURS ******************

	
^+q::

if WinExist("SVHM UniCare")
{
	WinMove, SVHM UniCare, , 0, 0,1291, 1000
}

if WinExist("SVHM UniCare")
{
    WinActivate 
}
WinWaitActive, SVHM UniCare, , 2
Sleep 500
if ErrorLevel
{
    MsgBox, Script timed out. Please Re-run.
    Return
}
else

Sleep 200
Send ^0 								; Sets zoom to Defualt
Sleep 200

CoordMode, Pixel, Window
PixelSearch, colourx, coloury, 1, 1, 300, 330, 0x00006D, 0, Fast RGB 
{					
	If ErrorLevel = 1 	; If not Found	
        {
		MsgBox, Click coordinates failed
	}
					
}
				;ADDRESS

    coordX := colourx+98
	coordY := coloury+150

	BlockInput, MouseMove
	CoordMode, Mouse, Window
	Send {Click, %coordX%, %coordY%}
	Sleep 500
	Send {Click, 3}
	Sleep 500
	Send ^c
	BlockInput, MouseMoveOff 

            address1 := Clipboard
			StringLower, address1, address1, T
			address1 := RegExReplace(address1, "[#]")


				;SUBURB

    coordX := colourx+98
	coordY := coloury+188

	BlockInput, MouseMove
	CoordMode, Mouse, Window
	Send {Click, %coordX%, %coordY%}
	Sleep 500
	Send {Click, 3}
	Sleep 500
	Send ^c
	BlockInput, MouseMoveOff 

            suburb := Clipboard
			StringLower, suburb, suburb, T


	; NEED TO FIX FROM HERE ON, TAKE ROUGHLY 15 OFF (x) AND 70 (y)

				;MOBILE

    coordX := colourx+555
	coordY := coloury+147

	BlockInput, MouseMove
	CoordMode, Mouse, Window
	Send {Click, %coordX%, %coordY%}
	Sleep 500
	Send {Click, 3}
	Sleep 500
	Send ^c
	BlockInput, MouseMoveOff 

            mobile := Clipboard




				;FIRSTNAME & LASTNAME

    coordX := colourx+20
	coordY := coloury+58

	BlockInput, MouseMove
	CoordMode, Mouse, Window
	Send {Click, %coordX%, %coordY%}
	Sleep 500
	Send {Click, 3}
	Sleep 500
	Send ^c
	BlockInput, MouseMoveOff 
    Sleep 200

        ;REMOVE ICONS FROM STRING

If InStr(Clipboard, "Medical Records")
    {
        StringGetPos, startPos, Clipboard, Medical Records 		 							        
        startPos := startPos + 1
        length := 200 					        
        if startPos >= 0
        
        Clipboard := StrReplace(Clipboard, (trimmed := SubStr(Clipboard, startPos, length)))
    }
    

If InStr(Clipboard, "Med Alerts")
{
    Clipboard := StrReplace(Clipboard, "Med Alerts")
}

else If InStr(Clipboard, "Micro Alerts")
{
    Clipboard := StrReplace(Clipboard, "Micro Alerts")
}

else If InStr(Clipboard, "Alerts")
{
    Clipboard := StrReplace(Clipboard, "Alerts")
}

            ;REMOVE TITLE

            If InStr(Clipboard, "Dr ", CaseSensitive)
            {
                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Patient Details 		; Gets the position of the beginning of the word First
                StringGetPos, length, Clipboard, Dr	                    ; Gets the position of the D
                startPos := startPos + 16 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := (length+1) - startPos 					    ; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                lastname := SubStr(Clipboard, startPos, length)
                StringLower, lastname, lastname, T 

                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Dr 		            ; Gets the position of the beginning of the word First
                startPos := startPos + 3 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := 50 					        				; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                firstname := SubStr(Clipboard, startPos, length)
                StringLower, firstname, firstname, T
            }
            else If InStr(Clipboard, "Mrs ", CaseSensitive)
            {
                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Patient Details 		; Gets the position of the beginning of the word First
                StringGetPos, length, Clipboard, Mrs	                ; Gets the position of the D
                startPos := startPos + 16 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := (length+1) - startPos 					    ; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                lastname := SubStr(Clipboard, startPos, length)
                StringLower, lastname, lastname, T

                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Mrs 		            ; Gets the position of the beginning of the word First
                startPos := startPos + 4 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := 50 					        				; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                firstname := SubStr(Clipboard, startPos, length)
                StringLower, firstname, firstname, T
            }
            else If InStr(Clipboard, "Miss ", CaseSensitive)
            {
                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Patient Details 		; Gets the position of the beginning of the word First
                StringGetPos, length, Clipboard, Miss	                ; Gets the position of the D
                startPos := startPos + 16 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := (length+1) - startPos 					    ; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                lastname := SubStr(Clipboard, startPos, length)
                StringLower, lastname, lastname, T

                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Miss 		        ; Gets the position of the beginning of the word First
                startPos := startPos + 5 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := 50 					        				; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                firstname := SubStr(Clipboard, startPos, length)
                StringLower, firstname, firstname, T
            }
            else If InStr(Clipboard, "Mr ", CaseSensitive)
            {
                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Patient Details 		; Gets the position of the beginning of the word First
                StringGetPos, length, Clipboard, Mr	                    ; Gets the position of the D
                startPos := startPos + 16 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := (length+1) - startPos 					    ; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                lastname := SubStr(Clipboard, startPos, length)
                StringLower, lastname, lastname, T 

                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Mr 		            ; Gets the position of the beginning of the word First
                startPos := startPos + 3 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := 50 					        				; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                firstname := SubStr(Clipboard, startPos, length)
                StringLower, firstname, firstname, T
            }
            else If InStr(Clipboard, "Ms ", CaseSensitive)
            {
                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Patient Details 		; Gets the position of the beginning of the word First
                StringGetPos, length, Clipboard, Ms	                    ; Gets the position of the D
                startPos := startPos + 16 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := (length+1) - startPos 					    ; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                lastname := SubStr(Clipboard, startPos, length)
                StringLower, lastname, lastname, T 

                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Ms 		            ; Gets the position of the beginning of the word First
                startPos := startPos + 3 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := 50 					        				; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                firstname := SubStr(Clipboard, startPos, length)
                StringLower, firstname, firstname, T
            }
            else If InStr(Clipboard, "Mx ", CaseSensitive)
            {
                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Patient Details 		; Gets the position of the beginning of the word First
                StringGetPos, length, Clipboard, Mx	                    ; Gets the position of the D
                startPos := startPos + 16 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := (length+1) - startPos 					    ; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                lastname := SubStr(Clipboard, startPos, length)
                StringLower, lastname, lastname, T 

                StringReplace, Clipboard, Clipboard, `r`n, , All
                StringGetPos, startPos, Clipboard, Mx 		            ; Gets the position of the beginning of the word First
                startPos := startPos + 3 							    ; Makes the starting pos at the end of the word First Name: in the string
                length := 50 					                        ; the number of characters between the end of the word and the beginning of the D char
                if startPos >= 0
                firstname := SubStr(Clipboard, startPos, length)
                StringLower, firstname, firstname, T
            }

			firstname := LTrim(firstname)
			lastname := LTrim(lastname)



        ;DOB

    coordX := colourx+50
	coordY := coloury+78

	BlockInput, MouseMove
	CoordMode, Mouse, Window
	Send {Click, %coordX%, %coordY%}
	Sleep 500
	Send {Click, 3}
	Sleep 500
	Send ^c
	BlockInput, MouseMoveOff 
    Sleep 200

        StringGetPos, startPos, Clipboard, Age		 							        
        startPos := startPos 
        length := 20				        
        if startPos >= 0
        
        dob := StrReplace(Clipboard, (trimmed := SubStr(Clipboard, startPos, length)))

        foundPosDay := RegExMatch(dob, "\d{2}", QdobDay)				; QdobDay
        foundPosYear := RegExMatch(dob, "\d{4}", QdobYear)				; QdobYear

        foundPosMonth := RegExMatch(dob, "[a-zA-Z]{3}", QdobMonth) 		; QdobMonth
        Sleep 300
        QdobMonth := dobMonthFunc(QdobMonth)
        Sleep 1000

        dob = %QdobDay%/%QdobMonth%/%QdobYear%

    ;UR NUMBER

CoordMode, Pixel, Window
	PixelSearch, coordX, coordY, 770, 180, 1277, 265, 0x0000DD, 0, Fast RGB 	;coordinates to search within work at all window sizes, use Window Spy for pixel colour and append '0x' to front
	{					
		If ErrorLevel = 1 	; If not Found	
		{
			MsgBox, Click coordinates failed
			Return
		}
	}
	BlockInput, MouseMove
	CoordMode, Mouse, Window
	Send {Click, %coordX%, %coordY%}
	Sleep 500
	Send {Click, 2}
	Sleep 500
	Send ^c
	BlockInput, MouseMoveOff
	urNumber := Clipboard

Sleep 200
MsgBox, 3,, Enter Patient details to Fever Clinic Sheet?
{
IfMsgBox, Yes
	{
		if WinExist("Fever Clinic XXX")
		{
				WinActivate 
			}
			WinWaitActive, Fever Clinic XXX, , 2
			if ErrorLevel
			{
				MsgBox, Script timed out. Please Re-run.
				return
			}
			else
			{
			
                Sleep 500
				Send {F5}'FEVER CLINIC Tracking'+1C1000{tab 3}{enter}
				Sleep 500
				Send {control down}
				Send {down}
				Send {control up}
				Send {down}
				Sleep 500
				Send {F2}
				Send %firstname% %lastname%
				Send {tab}
				Send %urNumber%
				Send {tab 2}
				Send %dob%
				Send {tab}
				Send %address1%, %suburb%
				Send {tab 7}				
				Sleep 500
				Send %mobile%
				Sleep 500
				Send {enter}
			
			Return
			}	
		
	}

IfMsgBox, No
	{			
		MsgBox, 3,, Enter Patient details to After Hours Sheet?
			{
				IfMsgBox, Yes
				{
					if WinExist("Fever Clinic XXX")
						{
							WinActivate 
						}
						WinWaitActive, Fever Clinic XXX, , 2
						if ErrorLevel
						{
							MsgBox, Script timed out. Please Re-run.
							return
						}
						else
						{
						
						Sleep 500
						Send {F5}'AFTER HOURS FEVER CLINIC'+1C900{tab 3}{enter}
						Sleep 500
						Send {control down}
						Send {down}
						Send {control up}
						Send {down}
						Send {F2}
						Send %firstname% %lastname%
						Send {tab}
						Send %urNumber%
						Send {tab}
						Send %dob%
						Send {tab}
						Send %address1%, %suburb%
						Send {tab 6}				
						Sleep 500
						Send %mobile%
						Sleep 500
						Send {enter}
						
						Return
						}	
				}
			}
	}
}
	
Return	





	;******************************* MOBILE FUNCTIONS *******************************



	; MOBILE PATTERN DETECT AND FORMAT 

mobilePattern(mobileall)
{

	if InStr(mobileall, "\+61") 															

		&& RegexMatch(mobileall, "4\d{2}\s?\d{3}\s?\d{3}")

		&& RegexMatch( phone := RegexReplace( phone, "\D" ), "^\d{10}$", phone )	;clean it

	{
		mobilepat1 := "0"SubStr(phone, 4, 3)" " SubStr(phone, 7, 3)" " SubStr(phone, 10, 3)	;and reformat to 04-- --- ---
		Sleep 500
		Send %mobilepat1%
		Return
	}

	else if RegexMatch(mobileall, "\+[0|2|3|4|5|7|8|9]{1,2}") 

	{ 
		Send N/A
		Return
	}


	else if RegexMatch(mobileall, "4\d{2}?\s?\d{3}\s?\d{3}", phone ) ; InStr(mobileall, "4\d{2}\s{0,1}\d{3}\s{0,1}\d{3}", false, 1, 1) 	;if 04 in string on mobileall

		;&& RegexMatch(mobileall, "4\d{2}?\s?\d{3}\s?\d{3}", phone )						;and the pattern 04-------- exists - was "[04\d]{8,8}"
				
		&& RegexMatch( phone := RegexReplace( phone, "\D" ), "^\d{9}$", phone )				;clean it
	{
		mobilepat1 := "0"SubStr( phone, 1, 3)" " SubStr( phone, 4, 3)" " SubStr( phone, 7, 3)	;and reformat to 04-- --- ---
		Sleep 500
		Send %mobilepat1%
		Return
	}


	else if RegexMatch(phone := RegexReplace(mobileall, "\D"), "0?[2|3|7|8]?\s?\d{4}\s?\d{4}", phoneOut)					; 0392313475

	{
		strLenMobile := StrLen(phone)
		if strLenMobile=10
		{
			mobilepat7 := SubStr(phone,1, 2)" " SubStr(phone,3,4)" " SubStr(phone,7,4)
			Sleep 500
			Send %mobilepat7%
			Return
		}

		else If strLenMobile=8

		{
			mobilepat7 := SubStr(phone, 1, 4)" "SubStr(phone, 5, 4)
			Sleep 500
			Send %mobilepat7%
			Return
		} 

		else
		{
			Sleep 500
			Send N/A
		}
	}



	else
	{
		Sleep 500
		Send N/A
	}

	Return

}




	;**************************************************************


	
	; Q-MOBILE PATTERN DETECT AND FORMAT 
	

QmobilePattern(mobileallQ)
{

	if InStr(mobileallQ, "\+61") 															

		&& RegexMatch(mobileallQ, "4\d{2}\s?\d{3}\s?\d{3}")

		&& RegexMatch( phone := RegexReplace( phone, "\D" ), "^\d{10}$", phone )	;clean it

	{
		mobilepat := "0"SubStr(phone, 4, 3)" " SubStr(phone, 7, 3)" " SubStr(phone, 10, 3)	;and reformat to 04-- --- ---
		Sleep 500
		;Send %mobilepat%
	}

	else if RegexMatch(mobileall, "\+[0|2|3|4|5|7|8|9]{1,2}") 

	{ 
		mobilepat := "N/A"
		
	}


	else if RegexMatch(mobileallQ, "4\d{2}?\s?\d{3}\s?\d{3}", phone ) ; InStr(mobileallQ, "4\d{2}\s{0,1}\d{3}\s{0,1}\d{3}", false, 1, 1) 	;if 04 in string on mobileall

		;&& RegexMatch(mobileallQ, "4\d{2}?\s?\d{3}\s?\d{3}", phone )						;and the pattern 04-------- exists - was "[04\d]{8,8}"
				
		&& RegexMatch( phone := RegexReplace( phone, "\D" ), "^\d{9}$", phone )				;clean it
	{
		mobilepat := "0"SubStr( phone, 1, 3)" " SubStr( phone, 4, 3)" " SubStr( phone, 7, 3)	;and reformat to 04-- --- ---
		Sleep 500
		;Send %mobilepat%
		
	}


	else if RegexMatch(phone := RegexReplace(mobileallQ, "\D"), "0?[2|3|7|8]?\s?\d{4}\s?\d{4}", phoneOut)					; 0392313475

	{
		strLenMobile := StrLen(phone)
		if strLenMobile=10
		{
			mobilepat := SubStr(phone,1, 2)" " SubStr(phone,3,4)" " SubStr(phone,7,4)
			Sleep 500
			;Send %mobilepat%
			
		}

		else If strLenMobile=8

		{
			mobilepat := SubStr(phone, 1, 4)" "SubStr(phone, 5, 4)
			Sleep 500
			;Send %mobilepat7%
			
		} 

		else
		{
			Sleep 500
			mobilepat := "N/A"
		}
	}



	else
	{
		Sleep 500
		mobilepat := "N/A"
	}

	Return mobilepat

}




	;**************************************************************


	; NOK MOBILE PATTERN DETECT AND FORMAT 


NOKmobilePattern(NOKmobileall)
{

	if InStr(NOKmobileall, "\+61") 															

		&& RegexMatch(NOKmobileall, "4\d{2}\s?\d{3}\s?\d{3}")

		&& RegexMatch( phone := RegexReplace( phone, "\D" ), "^\d{10}$", phone )	;clean it

	{
		mobilepat1 := "0"SubStr(phone, 4, 3)" " SubStr(phone, 7, 3)" " SubStr(phone, 10, 3)	;and reformat to 04-- --- ---
		Sleep 500
		Send %mobilepat1%
		Return
	}

	else if RegexMatch(NOKmobileall, "\+[0|2|3|4|5|7|8|9]{1,2}") 

	{ 
		Send N/A
		Return
	}


	else if RegexMatch(NOKmobileall, "4\d{2}?\s?\d{3}\s?\d{3}", phone ) ; InStr(NOKmobileall, "4\d{2}\s{0,1}\d{3}\s{0,1}\d{3}", false, 1, 1) 	;if 04 in string on mobileall

		;&& RegexMatch(NOKmobileall, "4\d{2}?\s?\d{3}\s?\d{3}", phone )						;and the pattern 04-------- exists - was "[04\d]{8,8}"
				
		&& RegexMatch( phone := RegexReplace( phone, "\D" ), "^\d{9}$", phone )				;clean it
	{
		mobilepat1 := "0"SubStr( phone, 1, 3)" " SubStr( phone, 4, 3)" " SubStr( phone, 7, 3)	;and reformat to 04-- --- ---
		Sleep 500
		Send %mobilepat1%
		Return
	}


	else if RegexMatch(phone := RegexReplace(NOKmobileall, "\D"), "0?[2|3|7|8]?\s?\d{4}\s?\d{4}", phoneOut)					; 0392313475

	{
		strLenMobile := StrLen(phone)
		if strLenMobile=10
		{
			mobilepat7 := SubStr(phone,1, 2)" " SubStr(phone,3,4)" " SubStr(phone,7,4)
			Sleep 500
			Send %mobilepat7%
			Return
		}

		else If strLenMobile=8

		{
			mobilepat7 := SubStr(phone, 1, 4)" "SubStr(phone, 5, 4)
			Sleep 500
			Send %mobilepat7%
			Return
		} 

		else
		{
			Sleep 500
			Send N/A
		}
	}



	else
	{
		Sleep 500
		Send N/A
	}

	Return

}





	;******************************* CONDITIONAL FUNCTIONS *******************************
	
	
	
	; TITLE FIELD SELECTION 
	
titleFunc(title)
{
if InStr(title, "Ms")
{
	{
		Sleep 500
		Send ms
		Sleep 500
		Send {tab}
	}
}

else if InStr(title, "rs", CaseSensitive := false, StartingPos := 2, Occurrence := 1)
{
	{
		Sleep 500
		Send a{up}
		Sleep 500
		Send ms{up}
		Sleep 500
		Send {tab}
	}
}

else if InStr(title, "Mr")
{
	{
		Sleep 500
		Send mr
		Sleep 500
		Send {tab}
	}
}

else if InStr(title, "doc") || InStr(title, "dr")
{
	{
		Sleep 500
		Send do
		Sleep 500
		Send {tab}
	}
}

else if InStr(title, "mis")
{
	{
		Sleep 500
		Send mis
		Sleep 500
		Send {tab}
	}
}

else if InStr(title, "mast")
{
	{
		Sleep 500
		Send mast
		Sleep 500
		Send {tab}
	}
}

else
{
	{
		Sleep 500
		Send a{up}
		Sleep 500
		Send {tab}
	}
}

Return 
}

	;**************************************************************

; SEX DETERMINATION

genderFunc(gender)
{
StringCaseSense, On

if InStr(gender, "fe")
{
	{
	Sleep 500
	Send fe
	Sleep 500
	}
}


else if InStr(gender, "ma")
{
	{
	Sleep 500
	Send ma
	Sleep 500
	}
}

else if InStr(gender, "ot")
{
	{
	Sleep 500
	Send ot
	Sleep 500
	}
}

else if InStr(gender, "ind")
{
	{
	Sleep 500
	Send in
	Sleep 500
	}
}

else if InStr(gender, "unk")
{
	{
	Sleep 500
	Send unk
	Sleep 500
	}
}

else
{
	{
	Sleep 500
	Send o{up}
	Sleep 500
	}
}

Return
}



	;**************************************************************

	; RELATIONSHIP CONDITIONALS

relationshipFunc(NOKrelationship)
{

	if InStr(NOKrelationship, "partner")
	{
		{
		Sleep 500
		Send par
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "wife")
	{
		{
		Sleep 500
		Send wi
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "husband")
	{
		{
		Sleep 500
		Send hus
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "uncle")
	{
		{
		Sleep 500
		Send uncle
		Sleep 500
		}
	}

	else if (InStr(NOKrelationship, "housemate") || InStr(NOKrelationship, "roommate") || InStr(NOKrelationship, "house mate") || InStr(NOKrelationship, "room mate") || InStr(NOKrelationship, "flat mate" || InStr(NOKrelationship, "flatmate")))
	{
		{
		Sleep 500
		Send friend
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "aunt")
	{
		{
		Sleep 500
		Send aunt
		Sleep 500
		}
	}



	else if InStr(NOKrelationship, "daught")
	{
		{
		Sleep 500
		Send daught
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "son")
	{
		{
		Sleep 500
		Send son
		Sleep 500
		}
	}

	else if (InStr(NOKrelationship, "defacto") || InStr(NOKrelationship, "de facto"))
	{
		{
		Sleep 500
		Send defact
		Sleep 500
		}
	}


	else if InStr(NOKrelationship, "boyfriend")
	{
		{
		Sleep 500
		Send boyfri
		Sleep 500
		}
	}


	else if InStr(NOKrelationship, "girlfriend")
	{
		{
		Sleep 500
		Send girlfri
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "father")
	{
		{
		Sleep 500
		Send father
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "mother")
	{
		{
		Sleep 500
		Send mother
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "dad")
	{
		{
		Sleep 500
		Send father
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "mum")
	{
		{
		Sleep 500
		Send mother
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "sister")
	{
		{
		Sleep 500
		Send sister
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "brother")
	{
		{
		Sleep 500
		Send brother
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "sibling")
	{
		{
		Sleep 500
		Send sibli
		Sleep 500
		}
	}


	else if InStr(NOKrelationship, "spouse")
	{
		{
		Sleep 500
		Send spous
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "cousin")
	{
		{
		Sleep 500
		Send cousin
		Sleep 500
		}
	}

	else if InStr(NOKrelationship, "friend")
	{
		{
		Sleep 500
		Send frie
		Sleep 500
		}
	}



	else
	{
		{
		Sleep 500
		}
	}

	Return

}

	;**************************************************************
	

	; MARITAL STATUS CONDITIONALS
	
	
	
maritalFunc(maritalStatus)
{
if InStr(maritalStatus, "never married")
{
	{
	Sleep 500
	Send never
	Sleep 500
	}
}

else if InStr(maritalStatus, "married")
{
	{
	Sleep 500
	Send marr
	Sleep 500
	}
}



else if InStr(maritalStatus, "de facto")
{
	{
	Sleep 500
	Send de
	Sleep 500
	}
}

else if InStr(maritalStatus, "divorced")
{
	{
	Sleep 500
	Send div
	Sleep 500
	}
}

else if (InStr(maritalStatus, "separated") || InStr(maritalStatus, "seperated"))
{
	{
	Sleep 500
	Send sep
	Sleep 500
	}
}

else
{
	{
	Sleep 500
	Send unkn
	Sleep 500
	}
}

Return
}

	;**************************************************************

	; RELIGION
	

religionFunc(religion)
{

if InStr(religion, "agnostic")
{
	{
	Sleep 500
	Send aus
	Send {down 2}
	Sleep 500
	}
}

else if InStr(religion, "anglican")
{
	{
	Sleep 500
	Send ang
	Sleep 500
	}
}

else if (InStr(religion, "atheist") || InStr(religion, "athiest"))
{
	{
	Sleep 500
	Send ath
	Sleep 500
	}
}

else if InStr(religion, "buddhis")
{
	{
	Sleep 500
	Send bud
	Sleep 500
	}
}


else if InStr(religion, "catholic")
{
	{
	Sleep 500
	Send cat
	Sleep 500
	}
}


else if InStr(religion, "christian")
{
	{
	Sleep 500
	Send christ
	Sleep 500
	}
}

else if (InStr(religion, "coe") || InStr(religion, "church of england"))
{
	{
	Sleep 500
	Send chur
	Send {down}
	Sleep 500
	}
}

else if InStr(religion, "orthodox") 		; Note: this is above greek to ensure no match against 'greek orthodox'
{
	{
	Sleep 500
	Send ort
	Sleep 500
	}
}


else if InStr(religion, "greek")
{
	{
	Sleep 500
	Send greek
	Sleep 500
	}
}

else if InStr(religion, "lutheran")
{
	{
	Sleep 500
	Send luth
	Sleep 500
	}
}

else if (InStr(religion, "hindu") || InStr(religion, "hindi"))
{
	{
	Sleep 500
	Send hindu
	Sleep 500
	}
}


else if InStr(religion, "jehovah")
{
	{
	Sleep 500
	Send jeh
	Sleep 500
	}
}


else if InStr(religion, "jewish")
{
	{
	Sleep 500
	Send jew
	Sleep 500
	}
}

else if (InStr(religion, "islam") || InStr(religion, "moslem") || InStr(religion, "muslim"))
{
	{
	Sleep 500
	Send mos
	Sleep 500
	}
}


else if (InStr(religion, "N/A") || InStr(religion, "none") || InStr(religion, "nil") || InStr(religion, "na"))
{
	{
	Sleep 500
	Send none
	Sleep 500
	}
}

else if InStr(religion, "jewish")
{
	{
	Sleep 500
	Send jew
	Sleep 500
	}
}

else if InStr(religion, "other")
{
	{
	Sleep 500
	Send oth
	Sleep 500
	}
}

else if InStr(religion, "presby")
{
	{
	Sleep 500
	Send pres
	Sleep 500
	}
}

else if InStr(religion, "protestant")
{
	{
	Sleep 500
	Send pro
	Sleep 500
	}
}

else if InStr(religion, "salvation")
{
	{
	Sleep 500
	Send sal
	Sleep 500
	}
}

else if InStr(religion, "sikh")
{
	{
	Sleep 500
	Send sikh
	Sleep 500
	}
}

else if InStr(religion, "uniting church")
{
	{
	Sleep 500
	Send uni
	Send {down}
	Sleep 500
	}
}


else
{
	{
	Sleep 500
	Send unkn
	Sleep 500
	}
}

Return	
}
	
	

	;**************************************************************


	
	; COUNTRY OF BIRTH 


cobFunc(cob)
{ 
if InStr(cob, "Austra")
{
	{
	Sleep 500
	Send aus
	Send {down 2}
	Sleep 500
	}
}

else if InStr(cob, "italy")
{
	{
	Sleep 500
	Send it
	Sleep 500
	}
}

else if (InStr(cob, "netherlands") || InStr(cob, "holland"))
{
	{
	Sleep 500
	Send netherla
	Sleep 500
	}
}

else if InStr(cob, "france")
{
	{
	Sleep 500
	Send fra
	Sleep 500
	}
}

else if InStr(cob, "germany")
{
	{
	Sleep 500
	Send germ
	Sleep 500
	}
}

else if (InStr(cob, "greece") || InStr(cob, "greek"))
{
	{
	Sleep 500
	Send gree
	Sleep 500
	}
}


else if InStr(cob, "sri lanka")
{
	{
	Sleep 500
	Send sri
	Sleep 500
	}
}

else if InStr(cob, "mauritius")
{
	{
	Sleep 500
	Send mauritius
	Sleep 500
	}
}


else if InStr(cob, "indonesia")
{
	{
	Sleep 500
	Send indones
	Sleep 500
	}
}

else if InStr(cob, "thailand")
{
	{
	Sleep 500
	Send thai
	Sleep 500
	}
}

else if InStr(cob, "nepal")
{
	{
	Sleep 500
	Send nep
	Sleep 500
	}
}

else if InStr(cob, "luxembourg")
{
	{
	Sleep 500
	Send lux
	Sleep 500
	}
}

else if InStr(cob, "norway")
{
	{
	Sleep 500
	Send norw
	Sleep 500
	}
}

else if InStr(cob, "sweden")
{
	{
	Sleep 500
	Send swed
	Sleep 500
	}
}

else if InStr(cob, "denmark")
{
	{
	Sleep 500
	Send den
	Sleep 500
	}
}

else if InStr(cob, "finland")
{
	{
	Sleep 500
	Send fin
	Sleep 500
	}
}

else if InStr(cob, "belgium")
{
	{
	Sleep 500
	Send belg
	Sleep 500
	}
}

else if InStr(cob, "switzerland")
{
	{
	Sleep 500
	Send swit
	Sleep 500
	}
}

else if InStr(cob, "china")
{
	{
	Sleep 500
	Send chin
	Sleep 500
	}
}

else if InStr(cob, "argentina")
{
	{
	Sleep 500
	Send arg
	Sleep 500
	}
}

else if InStr(cob, "colombia")
{
	{
	Sleep 500
	Send col
	Sleep 500
	}
}

else if InStr(cob, "brazil")
{
	{
	Sleep 500
	Send bra
	Sleep 500
	}
}

else if InStr(cob, "ireland")
{
	{
	Sleep 500
	Send ire
	Sleep 500
	}
}

else if InStr(cob, "england")
{
	{
	Sleep 500
	Send eng
	Sleep 500
	}
}

else if (InStr(cob, "united kingdom") || InStr(cob, "uk"))
{
	{
	Sleep 500
	Send uni
	Send {down}
	Sleep 500
	}
}

else if InStr(cob, "vietnam")
{
	{
	Sleep 500
	Send viet
	Sleep 500
	}
}

else if InStr(cob, "malaysia")
{
	{
	Sleep 500
	Send malay
	Sleep 500
	}
}

else if (InStr(cob, "USA") || InStr(cob, "United States"))
{
	{
	Sleep 500
	Send uni
	Send {down 2}
	Sleep 500
	}
}

else if InStr(cob, "canada")
{
	{
	Sleep 500
	Send cana
	Sleep 500
	}
}

else if InStr(cob, "taiwan")
{
	{
	Sleep 500
	Send tai
	Sleep 500
	}
}

else if InStr(cob, "south korea")
{
	{
	Sleep 500
	Send kor
	Send {down}
	Sleep 500
	}
}

else if InStr(cob, "scotland")
{
	{
	Sleep 500
	Send scot
	Sleep 500
	}
}

else if InStr(cob, "philippines")
{
	{
	Sleep 500
	Send phil
	Sleep 500
	}
}


else if InStr(cob, "samoa")
{
	{
	Sleep 500
	Send sam
	Sleep 500
	}
}

else if (InStr(cob, "new zealand") || InStr(cob, "nz"))
{
	{
	Sleep 500
	Send new
	Send {down 2}
	Sleep 500
	}
}

else if InStr(cob, "india")
{
	{
	Sleep 500
	Send ind
	Sleep 500
	}
}

else if InStr(cob, "singapore")
{
	{
	Sleep 500
	Send sing
	Sleep 500
	}
}

else if InStr(cob, "spain")
{
	{
	Sleep 500
	Send spa
	Sleep 500
	}
}

else if InStr(cob, "sudan")
{
	{
	Sleep 500
	Send sud
	Sleep 500
	}
}

else if InStr(cob, "iran")
{
	{
	Sleep 500
	Send iran
	Sleep 500
	}
}

else if InStr(cob, "iraq")
{
	{
	Sleep 500
	Send iraq
	Sleep 500
	}
}

else if InStr(cob, "hong")
{
	{
	Sleep 500
	Send hong
	Sleep 500
	}
}

else if InStr(cob, "ethiopia")
{
	{
	Sleep 500
	Send ethi
	Sleep 500
	}
}

else if InStr(cob, "kenya")
{
	{
	Sleep 500
	Send keny
	Sleep 500
	}
}

else
{
	{
	Sleep 500
	Send unkn
	Sleep 500
	}
}

Return 
}




	;**************************************************************
	

	; ABORIGINALITY CONDITIONALS

aboriginalityFunc(response)
{
if InStr(response, "Yes")
{
	{
		Sleep 500
		Send aborig
		Sleep 500
	}
}

else if InStr(response, "No")
{
	{
		Sleep 500
		Send no
		Sleep 500
	}
}

else
{
	{
		Sleep 500
		Send unable
		Sleep 500
	}
}

Return
}





	;**************************************************************

		
			; DOB MONTH CONDITIONALS

dobMonthFunc(QdobMonth)
{
	if Instr(QdobMonth, "Jan")
	{
		{
		Sleep 300		
		QdobMonth := "01"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Feb")
	{
		{
		Sleep 300		
		QdobMonth := "02"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Mar")
	{
		{
		Sleep 300		
		QdobMonth := "03"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Apr")
	{
		{
		Sleep 300		
		QdobMonth := "04"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "May")
	{
		{
		Sleep 300		
		QdobMonth := "05"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Jun")
	{
		{
		Sleep 300		
		QdobMonth := "06"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Jul")
	{
		{
		Sleep 300		
		QdobMonth := "07"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Aug")
	{
		{
		Sleep 300		
		QdobMonth := "08"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Sep")
	{
		{
		Sleep 300		
		QdobMonth := "09"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Oct")
	{
		{
		Sleep 300		
		QdobMonth := "10"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Nov")
	{
		{
		Sleep 300		
		QdobMonth := "11"
		Sleep 300	
		}
	}

	else if InStr(QdobMonth, "Dec")
	{
		{
		Sleep 300		
		QdobMonth := "12"
		Sleep 300	
		}
	}

	else
	{
		{
		Sleep 300		
		QdobMonth := "??"
		Sleep 300	
		}
	}
	
	Return QdobMonth
}





	;******************************* HOTKEY HELPER FUNCTIONS *******************************

	
	
 ; ************* GP CONSENT ***********************
 
 
 
gpConsentFunc(gpConsent)
{

if InStr(gpConsent, "yes")
{
	gpConsent := "Yes"
	Send %gpConsent%
}

else if InStr(gpConsent, "no")
{
	gpConsent := "No"
	Send %gpConsent%
}

else
{
	gpConsent := "N/A"
	Send %gpConsent%
}

Return
}

	
	; ******************* UR FUNCTION *************************

	
urFunc()			
{
if WinExist("SVHM UniCare")
{
    WinActivate 
}
WinWaitActive, SVHM UniCare, , 2
if ErrorLevel
{
    MsgBox, Script timed out. Please Re-run.
    return
}

else
{
	Sleep 500
	CoordMode, Mouse, Window
	Send {Click, 400, 500, right}
	Send v
	Sleep 500
	Send ^f
	Sleep 500
	Send PatientURN="{enter}
	Sleep 500
	Send {Esc}{right}
	Sleep 500
	Send {shift down}
	Send {right 8}
	Send {shift up}
	SendInput ^c
	SendInput ^w 
}

if WinExist("Fever Clinic XXX")
{
    WinActivate 
}
WinWaitActive, Fever Clinic XXX, , 2
if ErrorLevel
{
    MsgBox, Script timed out. Please Re-run.
    return
}
else
{
	Send {F5}'FEVER CLINIC Tracking'+1D5{enter}
	Sleep 500
	Send ^{down}{down}^v
}

	; omitted parameter below
if WinExist("XXX")
{
    WinClose
}
Sleep 500
Return
}

	;**************************************************************

	
	; PATIENT TYPE FUNCTION
	
patientTypeFunc(ptType)
{

if InStr(ptType, "staff")
	{
		pType1 := "Staff"
		Send %pType1%
		Return 
	}

else if InStr(ptType, "public")
	{ 	
		pType2 := "Public"
		Send %pType2%
		Return
	}	
	
else if InStr(ptType, "inpatient")
	{ 	
		pType3 := "Inpatient/Resident"
		Send %pType3%
		Return
	}	
	
else if InStr(ptType, "resident")
	{ 	
		pType4 := "Inpatient/Resident"
		Send %pType4%
		Return
	}	

else if InStr(ptType, "paramedic")
	{ 	
		pType5 := "Paramedic"
		Send %pType5%
		Return
	}	
	
else if InStr(ptType, "teacher")
	{
		pType6 := "Teacher"
		Send %pType6%
		Return 
	}
	
else if InStr(ptType, "health")
	{
		pType7 := "HCW"
		Send %pType7%
		Return 
	}

else
	{
		Sleep 200
		Return
	}
	
Return 
}
	

	;**************************************************************

	; STARTING POSITION 

startingPosition(userID)
{

if (userID=1)						; Outpatient or SuperUser Profile 
{
	Send {shift down}
	Send {tab 2}
	Send {shift up}
}

else if (userID=2 || userID=3 || userID=4)					; Elective Surgery or Emergency Profile
{
	Sleep 50
}


else
{
	MsgBox, ,Error,
	(LTrim, 
		There is an error with the user profile position.
		The program will exit. 
	)
	Return
}

Return 
}


	;**************************************************************

	; STARTING POSITION FOR NOK

startingPositionNOK(userID)
{

if (userID=1 || userID=3)						; Outpatient or Elective Surgery 
{
	Send {shift down}
	Send {tab}
	Send {shift up}
}

else if userID=2								; Emergency Profile
{
	Send {shift down}
	Send {tab 4}
	Send {shift up}
}


else
{
	MsgBox, ,Error,
	(LTrim, 
		There is an error with the user profile position.
		The program will exit. 
	)
	Return
}

Return 
}




	;**************************************************************


	; REGISTRATION FUNCTION

^l::
Send {tab 2}u{tab}unk{tab}unk{tab}u{tab}n{tab 3}rr{tab 4}{tab 2}n{tab 2}n{tab 12}u{tab}
Sleep 500
Return	


	;**************************************************************	
	
	
	; PARTIAL REGISTRATION FUNCTION
	
partialRego()
{
Send {tab}no
Sleep 100
Send {tab 3}
Sleep 100
Send rr{tab 4}{tab 2}n{tab 2}n{tab 12}u{tab}
Sleep 500
Return
}

	;**************************************************************

	; RELOAD

^+r::
Reload
Return

	;**************************************************************

	; EXIT SCRIPT

^Escape::
ExitApp
Return

	;**************************************************************

	; FORMAT MOBILE

^+f::
Send, {F2}{left 3}{space}{left 4}{space}{left 4}{backspace}0{left}{backspace 2}
Return

	;**************************************************************

	; FORMAT HOME PHONE
	
^+g::
Send {left 4}{space}{left 5}{backspace 3}03{space}
Return

	;**************************************************************

	; CONCATENATE STRING

Concatenate2(x, y) 
{
    Return, x y
}

	;**************************************************************
