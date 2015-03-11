Option Explicit   'Debugging function
Dim Beta_agency  'DIM for the rouine functions

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF



DIM DES_Workllist_Dialog, First_Line, Second_line, Additional_Information, Employer_Name, Verified, ButtonPressed, DES_Combobox, New_info, Combo1, message, DES_Worklist_Reviewed_Note, Case_Number   'Dim functions for dialog

BeginDialog DES_Workllist_Dialog, 0, 0, 386, 355, "Des Worklist Dialog"    'Inserting Dialog box 
  ButtonGroup ButtonPressed
    OkButton 260, 330, 50, 15
    CancelButton 325, 330, 50, 15
  ComboBox -350, 135, 60, 45, "", Combo1
  ComboBox 25, 25, 140, 35, "Select One"+chr(9)+"Old Information"+chr(9)+"No New Information"+chr(9)+"New Information", DES_Combobox
  Text 175, 30, 100, 10, "DES Worklist Reviewed Note"
  Text 30, 60, 60, 10, "New Information."
  EditBox 90, 60, 70, 15, New_info
  Text 30, 100, 60, 10, "Employer Name"
  EditBox 105, 95, 50, 15, Employer_Name
  DropListBox 105, 130, 60, 45, "Yes "+chr(9)+"No", Verified
  Text 35, 130, 50, 10, "Verified"
  Text 35, 165, 55, 10, "New Address"
  EditBox 105, 165, 120, 15, First_Line
  EditBox 105, 185, 120, 15, Second_line
  Text 30, 220, 85, 10, "Additional Information"
  EditBox 35, 240, 335, 15, Additional_Information
  Text 30, 265, 50, 10, "Case Number"
  EditBox 85, 260, 90, 20, Case_Number
EndDialog


EMConnect "" 'Connecting to bluezone
CALL Check_for_PRISM(True) 'checks to make sure we are in PRISM 


DO         'LOOPING to avoid not entering info 
DIALOG DES_Workllist_Dialog    
IF ButtonPressed = Cancel THEN StopScript
IF Employer_Name = "" THEN Msgbox "You must enter an employer!"  
LOOP UNTIL Employer_Name <> "" 


call navigate_to_PRISM_screen ("CAAD") 'Gets to CAAD
PF5 'creates a new note
EMWriteScreen "A", 3,29 'Sets to add note


'Writing the CAAD Note
EMWriteScreen "FREE", 4, 54  'sets as a free note

EMSetCursor 16, 4  'Set the cursor to the CAAD note area
CALL write_bullet_and_variable_in_CAAD ("DES Worklist Reviewed Note", DES_Combobox)   'bullets for each item entered on dialog box 
CALL write_bullet_and_variable_in_CAAD ("New Information", New_Info)
CALL write_bullet_and_variable_in_CAAD ("Employer Name", Employer_Name)
CALL write_bullet_and_variable_in_CAAD ("Verified", Verified)
CALL write_bullet_and_variable_in_CAAD ("First Line", First_Line)
CALL write_bullet_and_variable_in_CAAD ("Second Line", Second_Line)
CALL write_bullet_and_variable_in_CAAD ("Additional Information", Additional_Information) 


Transmit   'Enter at the end of CAAD Note

StopScript













