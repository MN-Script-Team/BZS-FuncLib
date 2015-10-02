'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'This script "library" contains functions and variables that the other BlueZone scripts use very commonly. The other BlueZone scripts contain a few lines of code that run 
'this script and get the functions. This saves time in writing and copy/pasting the same functions in many different places. Only add functions to this script if they've 
'been tested in other scripts first. This document is actively used by live scripts, so it needs to be functionally complete at all times. 
'
'============THAT MEANS THAT IF YOU BREAK THIS SCRIPT, ALL OTHER SCRIPTS ****STATEWIDE**** WILL NOT WORK! MODIFY WITH CARE!!!!!============
'
'
'Here's the code to add (remove comments before using):
'
''LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
'IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
'	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
'		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		Else																		'Everyone else should use the release branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		End if
'		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
'		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
'		req.send													'Sends request
'		IF req.Status = 200 THEN									'200 means great success
'			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
'			Execute req.responseText								'Executes the script code
'		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
'			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
'					vbCr & _
'					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
'					vbCr & _
'					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
'					vbTab & "- The name of the script you are running." & vbCr &_
'					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
'					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
'					vbTab & vbTab & "responsible for network issues." & vbCr &_
'					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
'					vbCr & _
'					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
'					vbCr &_
'					"URL: " & FuncLib_URL
'					script_end_procedure("Script ended due to error connecting to GitHub.")
'		END IF
'	ELSE
'		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
'		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
'		text_from_the_other_script = fso_command.ReadAll
'		fso_command.Close
'		Execute text_from_the_other_script
'	END IF
'END IF

'GLOBAL CONSTANTS----------------------------------------------------------------------------------------------------
Dim checked, unchecked, cancel, OK, blank		'Declares this for Option Explicit users

checked = 1			'Value for checked boxes
unchecked = 0		'Value for unchecked boxes
cancel = 0			'Value for cancel button in dialogs
OK = -1			'Value for OK button in dialogs
blank = ""

'Time arrays which can be used to fill an editbox with the convert_array_to_droplist_items function
time_array_15_min = array("7:00 AM", "7:15 AM", "7:30 AM", "7:45 AM", "8:00 AM", "8:15 AM", "8:30 AM", "8:45 AM", "9:00 AM", "9:15 AM", "9:30 AM", "9:45 AM", "10:00 AM", "10:15 AM", "10:30 AM", "10:45 AM", "11:00 AM", "11:15 AM", "11:30 AM", "11:45 AM", "12:00 PM", "12:15 PM", "12:30 PM", "12:45 PM", "1:00 PM", "1:15 PM", "1:30 PM", "1:45 PM", "2:00 PM", "2:15 PM", "2:30 PM", "2:45 PM", "3:00 PM", "3:15 PM", "3:30 PM", "3:45 PM", "4:00 PM", "4:15 PM", "4:30 PM", "4:45 PM", "5:00 PM", "5:15 PM", "5:30 PM", "5:45 PM", "6:00 PM")
time_array_30_min = array("7:00 AM", "7:30 AM", "8:00 AM", "8:30 AM", "9:00 AM", "9:30 AM", "10:00 AM", "10:30 AM", "11:00 AM", "11:30 AM", "12:00 PM", "12:30 PM", "1:00 PM", "1:30 PM", "2:00 PM", "2:30 PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM", "5:00 PM", "5:30 PM", "6:00 PM")

'BELOW ARE THE ACTUAL FUNCTIONS----------------------------------------------------------------------------------------------------

Function add_ACCI_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen ACCI_date, 8, 6, 73
  ACCI_date = replace(ACCI_date, " ", "/")
  If datediff("yyyy", ACCI_date, now) < 5 then
    EMReadScreen ACCI_type, 2, 6, 47
    If ACCI_type = "01" then ACCI_type = "Auto"
    If ACCI_type = "02" then ACCI_type = "Workers Comp"
    If ACCI_type = "03" then ACCI_type = "Homeowners"
    If ACCI_type = "04" then ACCI_type = "No Fault"
    If ACCI_type = "05" then ACCI_type = "Other Tort"
    If ACCI_type = "06" then ACCI_type = "Product Liab"
    If ACCI_type = "07" then ACCI_type = "Med Malprac"
    If ACCI_type = "08" then ACCI_type = "Legal Malprac"
    If ACCI_type = "09" then ACCI_type = "Diving Tort"
    If ACCI_type = "10" then ACCI_type = "Motorcycle"
    If ACCI_type = "11" then ACCI_type = "MTC or Other Bus Tort"
    If ACCI_type = "12" then ACCI_type = "Pedestrian"
    If ACCI_type = "13" then ACCI_type = "Other"
    x = x & ACCI_type & " on " & ACCI_date & ".; "
  End if
End function

Function add_ACCT_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen ACCT_amt, 8, 10, 46
  ACCT_amt = trim(ACCT_amt)
  ACCT_amt = "$" & ACCT_amt
  EMReadScreen ACCT_type, 2, 6, 44
  EMReadScreen ACCT_location, 20, 8, 44
  ACCT_location = replace(ACCT_location, "_", "")
  ACCT_location = split(ACCT_location)
  For each a in ACCT_location
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      If len(a) > 3 then
        new_ACCT_location = new_ACCT_location & b & c & " "
      Else
        new_ACCT_location = new_ACCT_location & a & " "
      End if
    End if
  Next
  EMReadScreen ACCT_ver, 1, 10, 63
  If ACCT_ver = "N" then 
    ACCT_ver = ", no proof provided"
  Else
    ACCT_ver = ""
  End if
  x = x & ACCT_type & " at " & new_ACCT_location & "(" & ACCT_amt & ")" & ACCT_ver & ".; "
  new_ACCT_location = ""
End function

Function add_BUSI_to_variable(variable_name_for_BUSI) 'x represents the name of the variable (example: assets vs. spousal_assets)
	'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
	EMReadScreen BUSI_footer_month, 5, 20, 55
	BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
	
	'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
	If datediff("d", "02/01/2015", BUSI_footer_month) < 0 then 
		EMReadScreen BUSI_type, 2, 5, 37
		If BUSI_type = "01" then BUSI_type = "Farming"
		If BUSI_type = "02" then BUSI_type = "Real Estate"
		If BUSI_type = "03" then BUSI_type = "Home Product Sales"
		If BUSI_type = "04" then BUSI_type = "Other Sales"
		If BUSI_type = "05" then BUSI_type = "Personal Services"
		If BUSI_type = "06" then BUSI_type = "Paper Route"
		If BUSI_type = "07" then BUSI_type = "InHome Daycare"
		If BUSI_type = "08" then BUSI_type = "Rental Income"
		If BUSI_type = "09" then BUSI_type = "Other"
		EMWriteScreen "x", 7, 26
		EMSendKey "<enter>"
		EMWaitReady 0, 0
		If cash_check = 1 then
			EMReadScreen BUSI_ver, 1, 9, 73
		ElseIf HC_check = 1 then 
			EMReadScreen BUSI_ver, 1, 12, 73
			If BUSI_ver = "_" then EMReadScreen BUSI_ver, 1, 13, 73
		ElseIf SNAP_check = 1 then
			EMReadScreen BUSI_ver, 1, 11, 73
		End if
		EMSendKey "<PF3>"
		EMWaitReady 0, 0
		If SNAP_check = 1 then
			EMReadScreen BUSI_amt, 8, 11, 68
			BUSI_amt = trim(BUSI_amt)
		ElseIf cash_check = 1 then 
			EMReadScreen BUSI_amt, 8, 9, 54
			BUSI_amt = trim(BUSI_amt)
		ElseIf HC_check = 1 then 
			EMWriteScreen "x", 17, 29
			EMSendKey "<enter>"
			EMWaitReady 0, 0
			EMReadScreen BUSI_amt, 8, 15, 54
			If BUSI_amt = "    0.00" then EMReadScreen BUSI_amt, 8, 16, 54
			BUSI_amt = trim(BUSI_amt)
			EMSendKey "<PF3>"
			EMWaitReady 0, 0
		End if
		variable_name_for_BUSI = variable_name_for_BUSI & trim(BUSI_type) & " BUSI"
		EMReadScreen BUSI_income_end_date, 8, 5, 71
		If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")
		If IsDate(BUSI_income_end_date) = True then
			variable_name_for_BUSI = variable_name_for_BUSI & " (ended " & BUSI_income_end_date & ")"
		Else
			If BUSI_amt <> "" then variable_name_for_BUSI = variable_name_for_BUSI & ", ($" & BUSI_amt & "/monthly)"
		End if
		If BUSI_ver = "N" or BUSI_ver = "?" then 
			variable_name_for_BUSI = variable_name_for_BUSI & ", no proof provided.; "
		Else
			variable_name_for_BUSI = variable_name_for_BUSI & ".; "
		End if
	Else		'------------This was updated 01/07/2015.
		'Checks the current footer month. If this is the future, it will know later on to read the HC pop-up
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		If datediff("d", date, BUSI_footer_month) > 0 then
			pull_future_HC = TRUE
		Else
			pull_future_HC = FALSE
		End if
	
		'Converting BUSI type code to a human-readable string
		EMReadScreen BUSI_type, 2, 5, 37
		If BUSI_type = "01" then BUSI_type = "Farming"
		If BUSI_type = "02" then BUSI_type = "Real Estate"
		If BUSI_type = "03" then BUSI_type = "Home Product Sales"
		If BUSI_type = "04" then BUSI_type = "Other Sales"
		If BUSI_type = "05" then BUSI_type = "Personal Services"
		If BUSI_type = "06" then BUSI_type = "Paper Route"
		If BUSI_type = "07" then BUSI_type = "InHome Daycare"
		If BUSI_type = "08" then BUSI_type = "Rental Income"
		If BUSI_type = "09" then BUSI_type = "Other"
		
		'Reading and converting BUSI Self employment method into human-readable 
		EMReadScreen BUSI_method, 2, 16, 53
		IF BUSI_method = "01" THEN BUSI_method = "50% Gross Income"
		IF BUSI_method = "02" THEN BUSI_method = "Tax Forms"
		
		'Going to the Gross Income Calculation pop-up
		EMWriteScreen "x", 6, 26
		transmit
		
		'Getting the verification codes for each type. Only does income, expenses are not included at this time.
		EMReadScreen BUSI_cash_ver, 1, 9, 73
		EMReadScreen BUSI_IVE_ver, 1, 10, 73
		EMReadScreen BUSI_SNAP_ver, 1, 11, 73
		EMReadScreen BUSI_HCA_ver, 1, 12, 73
		EMReadScreen BUSI_HCB_ver, 1, 13, 73
		
		'Converts each ver type to human readable
		If BUSI_cash_ver = "1" then BUSI_cash_ver = "tax returns provided"
		If BUSI_cash_ver = "2" then BUSI_cash_ver = "receipts provided"
		If BUSI_cash_ver = "3" then BUSI_cash_ver = "client ledger provided"
		If BUSI_cash_ver = "6" then BUSI_cash_ver = "other doc provided"
		If BUSI_cash_ver = "N" then BUSI_cash_ver = "no proof provided"
		If BUSI_cash_ver = "?" then BUSI_cash_ver = "no proof provided"
		If BUSI_IVE_ver = "1" then BUSI_IVE_ver = "tax returns provided"
		If BUSI_IVE_ver = "2" then BUSI_IVE_ver = "receipts provided"
		If BUSI_IVE_ver = "3" then BUSI_IVE_ver = "client ledger provided"
		If BUSI_IVE_ver = "6" then BUSI_IVE_ver = "other doc provided"
		If BUSI_IVE_ver = "N" then BUSI_IVE_ver = "no proof provided"
		If BUSI_IVE_ver = "?" then BUSI_IVE_ver = "no proof provided"
		If BUSI_SNAP_ver = "1" then BUSI_SNAP_ver = "tax returns provided"
		If BUSI_SNAP_ver = "2" then BUSI_SNAP_ver = "receipts provided"
		If BUSI_SNAP_ver = "3" then BUSI_SNAP_ver = "client ledger provided"
		If BUSI_SNAP_ver = "6" then BUSI_SNAP_ver = "other doc provided"
		If BUSI_SNAP_ver = "N" then BUSI_SNAP_ver = "no proof provided"
		If BUSI_SNAP_ver = "?" then BUSI_SNAP_ver = "no proof provided"
		If BUSI_HCA_ver = "1" then BUSI_HCA_ver = "tax returns provided"
		If BUSI_HCA_ver = "2" then BUSI_HCA_ver = "receipts provided"
		If BUSI_HCA_ver = "3" then BUSI_HCA_ver = "client ledger provided"
		If BUSI_HCA_ver = "6" then BUSI_HCA_ver = "other doc provided"
		If BUSI_HCA_ver = "N" then BUSI_HCA_ver = "no proof provided"
		If BUSI_HCA_ver = "?" then BUSI_HCA_ver = "no proof provided"
		If BUSI_HCB_ver = "1" then BUSI_HCB_ver = "tax returns provided"
		If BUSI_HCB_ver = "2" then BUSI_HCB_ver = "receipts provided"
		If BUSI_HCB_ver = "3" then BUSI_HCB_ver = "client ledger provided"
		If BUSI_HCB_ver = "6" then BUSI_HCB_ver = "other doc provided"
		If BUSI_HCB_ver = "N" then BUSI_HCB_ver = "no proof provided"
		If BUSI_HCB_ver = "?" then BUSI_HCB_ver = "no proof provided"
		
		'Back to the main screen
		PF3
		
		'Reading each income amount, trimming them to clean out unneeded spaces.
		EMReadScreen BUSI_cash_retro_amt, 8, 8, 55
		BUSI_cash_retro_amt = trim(BUSI_cash_retro_amt)
		EMReadScreen BUSI_cash_pro_amt, 8, 8, 69
		BUSI_cash_pro_amt = trim(BUSI_cash_pro_amt)
		EMReadScreen BUSI_IVE_amt, 8, 9, 69
		BUSI_IVE_amt = trim(BUSI_IVE_amt)
		EMReadScreen BUSI_SNAP_retro_amt, 8, 10, 55
		BUSI_SNAP_retro_amt = trim(BUSI_SNAP_retro_amt)
		EMReadScreen BUSI_SNAP_pro_amt, 8, 10, 69
		BUSI_SNAP_pro_amt = trim(BUSI_SNAP_pro_amt)
		
		'Pulls prospective amounts for HC, either from prosp side or from HC inc est.
		If pull_future_HC = False then
			EMReadScreen BUSI_HCA_amt, 8, 11, 69
			BUSI_HCA_amt = trim(BUSI_HCA_amt)
			EMReadScreen BUSI_HCB_amt, 8, 12, 69
			BUSI_HCB_amt = trim(BUSI_HCB_amt)
		Else
			EMWriteScreen "x", 17, 27
			transmit
			EMReadScreen BUSI_HCA_amt, 8, 15, 54
			BUSI_HCA_amt = trim(BUSI_HCA_amt)
			EMReadScreen BUSI_HCB_amt, 8, 16, 54
			BUSI_HCB_amt = trim(BUSI_HCB_amt)		
			PF3
		End if

		'Reads end date logic (in case it ended), converts to an actual date
		EMReadScreen BUSI_income_end_date, 8, 5, 72
		If BUSI_income_end_date <> "__ __ __" then BUSI_income_end_date = replace(BUSI_income_end_date, " ", "/")
		
		'Entering the variable details based on above
		variable_name_for_BUSI = variable_name_for_BUSI & trim(BUSI_type) & " BUSI:; "
		If IsDate(BUSI_income_end_date) = True then	variable_name_for_BUSI = variable_name_for_BUSI & "- Income ended " & BUSI_income_end_date & ".; "
		If BUSI_cash_retro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- Cash/GRH retro: $" & BUSI_cash_retro_amt & " budgeted, " & BUSI_cash_ver & "; "
		If BUSI_cash_pro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- Cash/GRH pro: $" & BUSI_cash_pro_amt & " budgeted, " & BUSI_cash_ver & "; "
		If BUSI_IVE_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- IV-E: $" & BUSI_IVE_amt & " budgeted, " & BUSI_IVE_ver & "; "
		If BUSI_SNAP_retro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- SNAP retro: $" & BUSI_SNAP_retro_amt & " budgeted, " & BUSI_SNAP_ver & "; "
		If BUSI_SNAP_pro_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- SNAP pro: $" & BUSI_SNAP_pro_amt & " budgeted, " & BUSI_SNAP_ver & "; "
		'Leaving out HC income estimator if footer month is not Current month + 1
		If BUSI_HCA_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- HC Method A: $" & BUSI_HCA_amt & " budgeted, " & BUSI_HCA_ver & "; "
		If BUSI_HCB_amt <> "0.00" then variable_name_for_BUSI = variable_name_for_BUSI & "- HC Method B: $" & BUSI_HCB_amt & " budgeted, " & BUSI_HCB_ver & "; "
		'Checks to see if pre 01/15 or post 02/15 then decides what to put in case note based on what was found/needed on the self employment method.
		If IsDate(BUSI_income_end_date) = false then
			IF BUSI_method <> "__" or BUSI_method = "" THEN 
				variable_name_for_BUSI = variable_name_for_BUSI & "- Self employment method: " & BUSI_method & "; "
			Else
				variable_name_for_BUSI = variable_name_for_BUSI & "- Self employment method: None; "
			END IF
		End if
	End if
End function

Function add_CARS_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen CARS_year, 4, 8, 31
  EMReadScreen CARS_make, 15, 8, 43
  CARS_make = replace(CARS_make, "_", "")
  EMReadScreen CARS_model, 15, 8, 66
  CARS_model = replace(CARS_model, "_", "")
  CARS_type = CARS_year & " " & CARS_make & " " & CARS_model
  CARS_type = split(CARS_type)
  For each a in CARS_type
    If len(a) > 1 then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      new_CARS_type = new_CARS_type & b & c & " "
    End if
  Next
  EMReadScreen CARS_amt, 8, 9, 45
  CARS_amt = trim(CARS_amt)
  CARS_amt = "$" & CARS_amt
  x = x & trim(new_CARS_type) & ", (" & CARS_amt & "); "
  new_CARS_type = ""
End function

Function add_JOBS_to_variable(variable_name_for_JOBS) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen JOBS_month, 5, 20, 55
  JOBS_month = replace(JOBS_month, " ", "/")
  EMReadScreen JOBS_type, 30, 7, 42
  JOBS_type = replace(JOBS_type, "_", ""	)
  JOBS_type = trim(JOBS_type)
  JOBS_type = split(JOBS_type)
  For each a in JOBS_type
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      new_JOBS_type = new_JOBS_type & b & c & " "
    End if
  Next
' Navigates to the FS PIC
    EMWriteScreen "x", 19, 38
    transmit
    EMReadScreen SNAP_JOBS_amt, 8, 17, 56
    SNAP_JOBS_amt = trim(SNAP_JOBS_amt)
    EMReadScreen snap_pay_frequency, 1, 5, 64
	EMReadScreen date_of_pic_calc, 8, 5, 34
	date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
    transmit
'  Reads the information on the retro side of JOBS
    EMReadScreen retro_JOBS_amt, 8, 17, 38
    retro_JOBS_amt = trim(retro_JOBS_amt)
'  Reads the information on the prospective side of JOBS
	EMReadScreen prospective_JOBS_amt, 8, 17, 67
	prospective_JOBS_amt = trim(prospective_JOBS_amt)
'  Reads the information about health care off of HC Income Estimator 
    EMReadScreen pay_frequency, 1, 18, 35
    EMWriteScreen "x", 19, 54
    transmit
    EMReadScreen HC_JOBS_amt, 8, 11, 63
    HC_JOBS_amt = trim(HC_JOBS_amt)
    transmit
  
  EMReadScreen JOBS_ver, 1, 6, 38
  EMReadScreen JOBS_income_end_date, 8, 9, 49
  If JOBS_income_end_date <> "__ __ __" then JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
  If IsDate(JOBS_income_end_date) = True then
    variable_name_for_JOBS = variable_name_for_JOBS & new_JOBS_type & "(ended " & JOBS_income_end_date & "); "
  Else
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" or pay_frequency = "5" then pay_frequency = "non-monthly"
    IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
    IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
    IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
    IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
    IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
    variable_name_for_JOBS = variable_name_for_JOBS & "EI from " & trim(new_JOBS_type) & ", " & JOBS_month  & " amts:; "
    If SNAP_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- PIC: $" & SNAP_JOBS_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
    If retro_JOBS_amt <> "" then variable_name_for_JOBS = variable_name_for_JOBS & "- Retrospective: $" & retro_JOBS_amt & " total; "
    IF prospective_JOBS_amt <> "" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- Prospective: $" & prospective_JOBS_amt & " total; "
    'Leaving out HC income estimator if footer month is not Current month + 1
    current_month_for_hc_est = dateadd("m", "1", date)
    current_month_for_hc_est = datepart("m", current_month_for_hc_est)
    IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
    IF footer_month = current_month_for_hc_est THEN 
	IF HC_JOBS_amt <> "________" THEN variable_name_for_JOBS = variable_name_for_JOBS & "- HC Inc Est: $" & HC_JOBS_amt & "/" & pay_frequency & "; "
    END IF
	If JOBS_ver = "N" or JOBS_ver = "?" then variable_name_for_JOBS = variable_name_for_JOBS & "- No proof provided for this panel; "
  End if
End function

Function add_OTHR_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen OTHR_type, 16, 6, 43
  OTHR_type = trim(OTHR_type)
  EMReadScreen OTHR_amt, 10, 8, 40
  OTHR_amt = trim(OTHR_amt)
  OTHR_amt = "$" & OTHR_amt
  x = x & trim(OTHR_type) & ", (" & OTHR_amt & ").; "
  new_OTHR_type = ""
End function

Function add_RBIC_to_variable(variable_name_for_RBIC) 'x represents the name of the variable (example: assets vs. spousal_assets)
	EMReadScreen RBIC_month, 5, 20, 55
	RBIC_month = replace(RBIC_month, " ", "/")
	EMReadScreen RBIC_type, 14, 5, 48
	RBIC_type = trim(RBIC_type)
	EMReadScreen RBIC01_pro_amt, 8, 10, 62
	RBIC01_pro_amt = trim(RBIC01_pro_amt)
	EMReadScreen RBIC02_pro_amt, 8, 11, 62
	RBIC02_pro_amt = trim(RBIC02_pro_amt)
	EMReadScreen RBIC03_pro_amt, 8, 12, 62
	RBIC03_pro_amt = trim(RBIC03_pro_amt)
	EMReadScreen RBIC01_retro_amt, 8, 10, 47
	IF RBIC01_retro_amt <> "________" THEN RBIC01_retro_amt = trim(RBIC01_retro_amt)
	EMReadScreen RBIC02_retro_amt, 8, 11, 47
	IF RBIC02_retro_amt <> "________" THEN RBIC02_retro_amt = trim(RBIC02_retro_amt)
	EMReadScreen RBIC03_retro_amt, 8, 12, 47
	IF RBIC03_retro_amt <> "________" THEN RBIC03_retro_amt = trim(RBIC03_retro_amt)
	EMReadScreen RBIC_group_01, 17, 10, 25
		RBIC_group_01 = replace(RBIC_group_01, " __", "")
		RBIC_group_01 = replace(RBIC_group_01, " ", ", ")
	EMReadScreen RBIC_group_02, 17, 11, 25
		RBIC_group_02 = replace(RBIC_group_02, " __", "")
		RBIC_group_02 = replace(RBIC_group_02, " ", ", ")
	EMReadScreen RBIC_group_03, 17, 12, 25
		RBIC_group_03 = replace(RBIC_group_03, " __", "")
		RBIC_group_03 = replace(RBIC_group_03, " ", ", ")
	
	EMReadScreen RBIC_01_verif, 1, 10, 76
	IF RBIC_01_verif = "N" THEN
		RBIC01_pro_amt = RBIC01_pro_amt & ", not verified"
		RBIC01_retro_amt = RBIC01_retro_amt & ", not verified"
	END IF
	
	EMReadScreen RBIC_02_verif, 1, 11, 76
	IF RBIC_02_verif = "N" THEN
		RBIC02_pro_amt = RBIC02_pro_amt & ", not verified"
		RBIC02_retro_amt = RBIC02_retro_amt & ", not verified"
	END IF
	
	EMReadScreen RBIC_03_verif, 1, 12, 76
	IF RBIC_03_verif = "N" THEN
		RBIC03_pro_amt = RBIC03_pro_amt & ", not verified"
		RBIC03_retro_amt = RBIC03_retro_amt & ", not verified"
	END IF
	
	RBIC_expense_row = 15
	DO
		EMReadScreen RBIC_expense_type, 13, RBIC_expense_row, 28
		RBIC_expense_type = trim(RBIC_expense_type)
		EMReadScreen RBIC_expense_amt, 8, RBIC_expense_row, 62
		RBIC_expense_amt = trim(RBIC_expense_amt)
		EMReadScreen RBIC_expense_verif, 1, RBIC_expense_row, 76
		IF RBIC_expense_type <> "" THEN
			total_RBIC_expenses = total_RBIC_expenses & "- " & RBIC_expense_type & ", $" & RBIC_expense_amt
			IF RBIC_expense_verif <> "N" THEN
				total_RBIC_expenses = total_RBIC_expenses & "; "
			ELSE
				total_RBIC_expenses = total_RBIC_expenses & ", not verified; "
			END IF
			RBIC_expense_row = RBIC_expense_row + 1
			IF RBIC_expense_row = 19 THEN
				PF20
				EMReadScreen RBIC_last_page, 21, 24, 2
				RBIC_expense_row = 15
			END IF
		END IF
	LOOP UNTIL RBIC_expense_type = "" OR RBIC_last_page = "THIS IS THE LAST PAGE"
	EMReadScreen RBIC_ver, 1, 10, 76
	If RBIC_ver = "N" then RBIC_ver = ", no proof provided"
	EMReadScreen RBIC_end_date, 8, 6, 68
	RBIC_end_date = replace(RBIC_end_date, " ", "/")
	If isdate(RBIC_end_date) = True then
		variable_name_for_RBIC = variable_name_for_RBIC & trim(RBIC_type) & " RBIC, ended " & RBIC_end_date & RBIC_ver & "; "
	Else
		IF left(RBIC01_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_01 & ", Prospective, ($" & RBIC01_pro_amt & "); "
		IF left(RBIC01_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_01 & ", Retrospective, ($" & RBIC01_retro_amt & "); "
		IF left(RBIC02_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_02 & ", Prospective, ($" & RBIC02_pro_amt & "); "
		IF left(RBIC02_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_02 & ", Retrospective, ($" & RBIC02_retro_amt & "); "
		IF left(RBIC03_pro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_03 & ", Prospective, ($" & RBIC03_pro_amt & "); "
		IF left(RBIC03_retro_amt, 1) <> "_" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC: " & trim(RBIC_type) & " from MEMB(s) " & RBIC_group_03 & ", Retrospective, ($" & RBIC03_retro_amt & "); "
		IF total_RBIC_expenses <> "" THEN variable_name_for_RBIC = variable_name_for_RBIC & "RBIC Expenses:; " & total_RBIC_expenses
	End if
End function

Function add_REST_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen REST_type, 16, 6, 41
  REST_type = trim(REST_type)
  EMReadScreen REST_amt, 10, 8, 41
  REST_amt = trim(REST_amt)
  REST_amt = "$" & REST_amt
  x = x & trim(REST_type) & ", (" & REST_amt & ").; "
  new_REST_type = ""
End function


Function add_SECU_to_variable(x) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen SECU_amt, 8, 10, 52
  SECU_amt = trim(SECU_amt)
  SECU_amt = "$" & SECU_amt
  EMReadScreen SECU_type, 2, 6, 50
  EMReadScreen SECU_location, 20, 8, 50
  SECU_location = replace(SECU_location, "_", "")
  SECU_location = split(SECU_location)
  For each a in SECU_location
    If a <> "" then
      b = ucase(left(a, 1))
      c = LCase(right(a, len(a) -1))
      If len(a) > 3 then
        new_SECU_location = new_SECU_location & b & c & " "
      Else
        new_SECU_location = new_SECU_location & a & " "
      End if
    End if
  Next
  EMReadScreen SECU_ver, 1, 11, 50
  If SECU_ver = "1" then SECU_ver = "agency form provided"
  If SECU_ver = "2" then SECU_ver = "source doc provided"
  If SECU_ver = "3" then SECU_ver = "verified via phone"
  If SECU_ver = "5" then SECU_ver = "other doc verified"
  If SECU_ver = "N" then SECU_ver = "no proof provided"
  x = x & SECU_type & " at " & new_SECU_location & " (" & SECU_amt & "), " & SECU_ver & ".; "
  new_SECU_location = ""
End function

Function add_UNEA_to_variable(variable_name_for_UNEA) 'x represents the name of the variable (example: assets vs. spousal_assets)
  EMReadScreen UNEA_month, 5, 20, 55
  UNEA_month = replace(UNEA_month, " ", "/")
  EMReadScreen UNEA_type, 16, 5, 40
  If UNEA_type = "Unemployment Ins" then UNEA_type = "UC"
  If UNEA_type = "Disbursed Child " then UNEA_type = "CS"
  If UNEA_type = "Disbursed CS Arr" then UNEA_type = "CS arrears"
  UNEA_type = trim(UNEA_type)
  EMReadScreen UNEA_ver, 1, 5, 65
  EMReadScreen UNEA_income_end_date, 8, 7, 68
  If UNEA_income_end_date <> "__ __ __" then UNEA_income_end_date = replace(UNEA_income_end_date, " ", "/")
  If IsDate(UNEA_income_end_date) = True then
    variable_name_for_UNEA = variable_name_for_UNEA & UNEA_type & " (ended " & UNEA_income_end_date & "); "
  Else
    EMReadScreen UNEA_amt, 8, 18, 68
    UNEA_amt = trim(UNEA_amt)
      EMWriteScreen "x", 10, 26
      transmit
      EMReadScreen SNAP_UNEA_amt, 8, 17, 56
      SNAP_UNEA_amt = trim(SNAP_UNEA_amt)
      EMReadScreen snap_pay_frequency, 1, 5, 64
	EMReadScreen date_of_pic_calc, 8, 5, 34
	date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
      transmit
      EMReadScreen retro_UNEA_amt, 8, 18, 39
      retro_UNEA_amt = trim(retro_UNEA_amt)
	EMReadScreen prosp_UNEA_amt, 8, 18, 68
	prosp_UNEA_amt = trim(prosp_UNEA_amt)
      EMWriteScreen "x", 6, 56
      transmit
      EMReadScreen HC_UNEA_amt, 8, 9, 65
      HC_UNEA_amt = trim(HC_UNEA_amt)
      EMReadScreen pay_frequency, 1, 10, 63
      transmit
      If HC_UNEA_amt = "________" then
        EMReadScreen HC_UNEA_amt, 8, 18, 68
        HC_UNEA_amt = trim(HC_UNEA_amt)
        pay_frequency = "mo budgeted prospectively"
    End If
    If pay_frequency = "1" then pay_frequency = "monthly"
    If pay_frequency = "2" then pay_frequency = "semimonthly"
    If pay_frequency = "3" then pay_frequency = "biweekly"
    If pay_frequency = "4" then pay_frequency = "weekly"
    If pay_frequency = "_" then pay_frequency = "non-monthly"
    IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
    IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
    IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
    IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
    IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"
    variable_name_for_UNEA = variable_name_for_UNEA & "UNEA from " & trim(UNEA_type) & ", " & UNEA_month  & " amts:; "
    If SNAP_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
    If retro_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Retrospective: $" & retro_UNEA_amt & " total; "
    If prosp_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Prospective: $" & prosp_UNEA_amt & " total; "
    'Leaving out HC income estimator if footer month is not Current month + 1
    current_month_for_hc_est = dateadd("m", "1", date)
    current_month_for_hc_est = datepart("m", current_month_for_hc_est)
    IF len(current_month_for_hc_est) = 1 THEN current_month_for_hc_est = "0" & current_month_for_hc_est
    IF footer_month = current_month_for_hc_est THEN
    	If HC_UNEA_amt <> "" THEN variable_name_for_UNEA = variable_name_for_UNEA & "- HC Inc Est: $" & HC_UNEA_amt & "/" & pay_frequency & "; "
    END IF
    If UNEA_ver = "N" or UNEA_ver = "?" then variable_name_for_UNEA = variable_name_for_UNEA & "- No proof provided for this panel; "
  End if
End function

'This function will assign an address to a variable selected from the interview_location variable in the Appt Letter script.
Function assign_county_address_variables(address_line_01, address_line_02)		
	For each office in county_office_array				'Splits the county_office_array, which is set by the config program and declared earlier in this file
		If instr(office, interview_location) <> 0 then		'If the name of the office is found in the "interview_location" variable, which is contained in the MEMO - appt letter script.
			new_office_array = split(office, "|")		'Split the office into its own array
			address_line_01 = new_office_array(0)		'Line 1 of the address is the first part of this array
			address_line_02 = new_office_array(1)		'Line 2 of the address is the second part of this array
		End if
	Next
End function

Function attn
  EMSendKey "<attn>"
  EMWaitReady -1, 0
End function

Function autofill_editbox_from_MAXIS(HH_member_array, panel_read_from, variable_written_to)
 'First it navigates to the screen. Only does the first four characters because we use separate handling for HCRE-retro. This is something that should be fixed someday!!!!!!!!!
  call navigate_to_MAXIS_screen("stat", left(panel_read_from, 4))
  
  'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
  EMReadScreen panel_total_check, 6, 2, 73
  IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info
  
  If variable_written_to <> "" then variable_written_to = variable_written_to & "; "
  If panel_read_from = "ABPS" then '--------------------------------------------------------------------------------------------------------ABPS
    EMReadScreen ABPS_total_pages, 1, 2, 78
    If ABPS_total_pages <> 0 then 
      Do
        'First it checks the support coop. If it's "N" it'll add a blurb about it to the support_coop variable
        EMReadScreen support_coop_code, 1, 4, 73
        If support_coop_code = "N" then
          EMReadScreen caregiver_ref_nbr, 2, 4, 47
          If instr(support_coop, "Memb " & caregiver_ref_nbr & " not cooperating with child support; ") = 0 then support_coop = support_coop & "Memb " & caregiver_ref_nbr & " not cooperating with child support; "'the if...then statement makes sure the info isn't duplicated. 
        End if
        'Then it gets info on the ABPS themself.
        EMReadScreen ABPS_current, 45, 10, 30
        If ABPS_current = "________________________  First: ____________" then ABPS_current = "Parent unknown"
        ABPS_current = replace(ABPS_current, "  First:", ",")
        ABPS_current = replace(ABPS_current, "_", "")
        ABPS_current = split(ABPS_current)
        For each a in ABPS_current
          b = ucase(left(a, 1))
          c = LCase(right(a, len(a) -1))
          If len(a) > 1 then
            new_ABPS_current = new_ABPS_current & b & c & " "
          Else
            new_ABPS_current = new_ABPS_current & a & " "
          End if
        Next
        ABPS_row = 15 'Setting variable for do...loop
        Do
          Do 'Using a do...loop to determine which MEMB numbers are with this parent
            EMReadScreen child_ref_nbr, 2, ABPS_row, 35
            If child_ref_nbr <> "__" then
              amt_of_children_for_ABPS = amt_of_children_for_ABPS + 1
              children_for_ABPS = children_for_ABPS & child_ref_nbr & ", "
            End if
            ABPS_row = ABPS_row + 1
          Loop until ABPS_row > 17		'End of the row
          EMReadScreen more_check, 7, 19, 66
          If more_check = "More: +" then
            EMSendKey "<PF20>"
            EMWaitReady 0, 0
            ABPS_row = 15
          End if
        Loop until more_check <> "More: +"
        'Cleaning up the "children_for_ABPS" variable to be more readable
        children_for_ABPS = left(children_for_ABPS, len(children_for_ABPS) - 2) 'cleaning up the end of the variable (removing the comma for single kids)
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it around to change the last comma to an "and"
        children_for_ABPS = replace(children_for_ABPS, ",", "dna ", 1, 1)        'it's backwards, replaces just one comma with an "and"
        children_for_ABPS = strreverse(children_for_ABPS)                       'flipping it back around 
        if amt_of_children_for_ABPS > 1 then HH_memb_title = " for membs "
        if amt_of_children_for_ABPS <= 1 then HH_memb_title = " for memb "
        variable_written_to = variable_written_to & trim(new_ABPS_current) & HH_memb_title & children_for_ABPS & "; "
        'Resetting variables for the do...loop in case this function runs again
        new_ABPS_current = "" 
        amt_of_children_for_ABPS = 0
        children_for_ABPS = ""
        'Checking to see if it needs to run again, if it does it transmits or else the loop stops
        EMReadScreen ABPS_current_page, 1, 2, 73
        If ABPS_current_page <> ABPS_total_pages then transmit
      Loop until ABPS_current_page = ABPS_total_pages
      'Combining the two variables (support coop and the variable written to)
      variable_written_to = support_coop & variable_written_to
    End if
  Elseif panel_read_from = "ACCI" then '----------------------------------------------------------------------------------------------------ACCI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCI_total, 1, 2, 78
      If ACCI_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCI_to_variable(variable_written_to)
          EMReadScreen ACCI_panel_current, 1, 2, 73
          If cint(ACCI_panel_current) < cint(ACCI_total) then transmit
        Loop until cint(ACCI_panel_current) = cint(ACCI_total)
      End if
    Next
  Elseif panel_read_from = "ACCT" then '----------------------------------------------------------------------------------------------------ACCT
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen ACCT_total, 1, 2, 78
      If ACCT_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_ACCT_to_variable(variable_written_to)
          EMReadScreen ACCT_panel_current, 1, 2, 73
          If cint(ACCT_panel_current) < cint(ACCT_total) then transmit
        Loop until cint(ACCT_panel_current) = cint(ACCT_total)
      End if
    Next
  Elseif panel_read_from = "ADDR" then '----------------------------------------------------------------------------------------------------ADDR
    EMReadScreen addr_line_01, 22, 6, 43
    EMReadScreen addr_line_02, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 12, 9, 43
    variable_written_to = replace(addr_line_01, "_", "") & "; " & replace(addr_line_02, "_", "") & "; " & replace(city_line, "_", "") & ", " & state_line & " " & replace(zip_line, "__ ", "-")
    variable_written_to = replace(variable_written_to, "; ; ", "; ") 'in case there's only one line on ADDR
  Elseif panel_read_from = "AREP" then '----------------------------------------------------------------------------------------------------AREP
    EMReadScreen AREP_name, 37, 4, 32
    AREP_name = replace(AREP_name, "_", "")
    AREP_name = split(AREP_name)
    For each word in AREP_name
      If word <> "" then
        first_letter_of_word = ucase(left(word, 1))
        rest_of_word = LCase(right(word, len(word) -1))
        If len(word) > 2 then
          variable_written_to = variable_written_to & first_letter_of_word & rest_of_word & " "
        Else
          variable_written_to = variable_written_to & word & " "
        End if
      End if
    Next
  Elseif panel_read_from = "BILS" then '----------------------------------------------------------------------------------------------------BILS
    EMReadScreen BILS_amt, 1, 2, 78
    If BILS_amt <> 0 then variable_written_to = "BILS known to MAXIS."
  Elseif panel_read_from = "BUSI" then '----------------------------------------------------------------------------------------------------BUSI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen BUSI_total, 1, 2, 78
      If BUSI_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_BUSI_to_variable(variable_written_to)
          EMReadScreen BUSI_panel_current, 1, 2, 73
          If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
        Loop until cint(BUSI_panel_current) = cint(BUSI_total)
      End if
    Next
  Elseif panel_read_from = "CARS" then '----------------------------------------------------------------------------------------------------CARS
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen CARS_total, 1, 2, 78
      If CARS_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_CARS_to_variable(variable_written_to)
          EMReadScreen CARS_panel_current, 1, 2, 73
          If cint(CARS_panel_current) < cint(CARS_total) then transmit
        Loop until cint(CARS_panel_current) = cint(CARS_total)
      End if
    Next
  Elseif panel_read_from = "CASH" then '----------------------------------------------------------------------------------------------------CASH
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen cash_amt, 8, 8, 39
      cash_amt = trim(cash_amt)
      If cash_amt <> "________" then
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Cash ($" & cash_amt & "); "
      End if
    Next
  Elseif panel_read_from = "COEX" then '----------------------------------------------------------------------------------------------------COEX
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen support_amt, 8, 10, 63
      support_amt = trim(support_amt)
      If support_amt <> "________" then
        EMReadScreen support_ver, 1, 10, 36
        If support_ver = "?" or support_ver = "N" then
          support_ver = ", no proof provided"
        Else
          support_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Support ($" & support_amt & "/mo" & support_ver & "); "
      End if
      EMReadScreen alimony_amt, 8, 11, 63
      alimony_amt = trim(alimony_amt)
      If alimony_amt <> "________" then
        EMReadScreen alimony_ver, 1, 11, 36
        If alimony_ver = "?" or alimony_ver = "N" then
          alimony_ver = ", no proof provided"
        Else
          alimony_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Alimony ($" & alimony_amt & "/mo" & alimony_ver & "); "
      End if
      EMReadScreen tax_dep_amt, 8, 12, 63
      tax_dep_amt = trim(tax_dep_amt)
      If tax_dep_amt <> "________" then
        EMReadScreen tax_dep_ver, 1, 12, 36
        If tax_dep_ver = "?" or tax_dep_ver = "N" then
          tax_dep_ver = ", no proof provided"
        Else
          tax_dep_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Tax dep ($" & tax_dep_amt & "/mo" & tax_dep_ver & "); "
      End if
      EMReadScreen other_COEX_amt, 8, 13, 63
      other_COEX_amt = trim(other_COEX_amt)
      If other_COEX_amt <> "________" then
        EMReadScreen other_COEX_ver, 1, 13, 36
        If other_COEX_ver = "?" or other_COEX_ver = "N" then
          other_COEX_ver = ", no proof provided"
        Else
          other_COEX_ver = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & "Other ($" & other_COEX_amt & "/mo" & other_COEX_ver & "); "
      End if
    Next
  Elseif panel_read_from = "DCEX" then '----------------------------------------------------------------------------------------------------DCEX
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
	  EMReadScreen DCEX_total, 1, 2, 78
      If DCEX_total <> 0 then
		variable_written_to = variable_written_to & "Member " & HH_member & "- "
		Do
			DCEX_row = 11
			Do
				EMReadScreen expense_amt, 8, DCEX_row, 63
				expense_amt = trim(expense_amt)
				If expense_amt <> "________" then
					EMReadScreen child_ref_nbr, 2, DCEX_row, 29
					EMReadScreen expense_ver, 1, DCEX_row, 41
					If expense_ver = "?" or expense_ver = "N" or expense_ver = "_" then
						expense_ver = ", no proof provided"
					Else
						expense_ver = ""
					End if
					variable_written_to = variable_written_to & "Child " & child_ref_nbr & " ($" & expense_amt & "/mo DCEX" & expense_ver & "); "
				End if
				DCEX_row = DCEX_row + 1
			Loop until DCEX_row = 17
			EMReadScreen DCEX_panel_current, 1, 2, 73
			If cint(DCEX_panel_current) < cint(DCEX_total) then transmit
		Loop until cint(DCEX_panel_current) = cint(DCEX_total)
	  End if
    Next
  Elseif panel_read_from = "DIET" then '----------------------------------------------------------------------------------------------------DIET
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      DIET_row = 8 'Setting this variable for the next do...loop
      EMReadScreen DIET_total, 1, 2, 78
      If DIET_total <> 0 then 
        DIET = DIET & "Member " & HH_member & "- "
        Do
          EMReadScreen diet_type, 2, DIET_row, 40
          EMReadScreen diet_proof, 1, DIET_row, 51
          If diet_proof = "_" or diet_proof = "?" or diet_proof = "N" then 
            diet_proof = ", no proof provided"
          Else
            diet_proof = ""
          End if
          If diet_type = "01" then diet_type = "High Protein"
          If diet_type = "02" then diet_type = "Cntrl Protein (40-60 g/day)"
          If diet_type = "03" then diet_type = "Cntrl Protein (<40 g/day)"
          If diet_type = "04" then diet_type = "Lo Cholesterol"
          If diet_type = "05" then diet_type = "High Residue"
          If diet_type = "06" then diet_type = "Preg/Lactation"
          If diet_type = "07" then diet_type = "Gluten Free"
          If diet_type = "08" then diet_type = "Lactose Free"
          If diet_type = "09" then diet_type = "Anti-Dumping"
          If diet_type = "10" then diet_type = "Hypoglycemic"
          If diet_type = "11" then diet_type = "Ketogenic"
          If diet_type <> "__" and diet_type <> "  " then variable_written_to = variable_written_to & diet_type & diet_proof & "; "
          DIET_row = DIET_row + 1
        Loop until DIET_row = 19
      End if
    Next
  Elseif panel_read_from = "DISA" then '----------------------------------------------------------------------------------------------------DISA
    For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
	  EMReadscreen DISA_total, 1, 2, 78
	  IF DISA_total <> 0 THEN
		'Reads and formats CASH/GRH disa status
		EMReadScreen CASH_DISA_status, 2, 11, 59
		EMReadScreen CASH_DISA_verif, 1, 11, 69
		IF CASH_DISA_status = "01" or CASH_DISA_status = "02" or CASH_DISA_status = "03" OR CASH_DISA_status = "04" THEN CASH_DISA_status = "RSDI/SSI certified"
		IF CASH_DISA_status = "06" THEN CASH_DISA_status = "SMRT/SSA pends"
		IF CASH_DISA_status = "08" THEN CASH_DISA_status = "Certified Blind"
		IF CASH_DISA_status = "09" THEN CASH_DISA_status = "Ill/Incap"
		IF CASH_DISA_status = "10" THEN CASH_DISA_status = "Certified disabled"
		IF CASH_DISA_verif = "?" OR CASH_DISA_verif = "N" THEN
			CASH_DISA_verif = ", no proof provided"
		ELSE
			CASH_DISA_verif = ""
		END IF
		
		'Reads and formats SNAP disa status
		EmreadScreen SNAP_DISA_status, 2, 12, 59
		EMReadScreen SNAP_DISA_verif, 1, 12, 69
		IF SNAP_DISA_status = "01" or SNAP_DISA_status = "02" or SNAP_DISA_status = "03" OR SNAP_DISA_status = "04" THEN SNAP_DISA_status = "RSDI/SSI certified"
		IF SNAP_DISA_status = "08" THEN SNAP_DISA_status = "Certified Blind"
		IF SNAP_DISA_status = "09" THEN SNAP_DISA_status = "Ill/Incap"
		IF SNAP_DISA_status = "10" THEN SNAP_DISA_status = "Certified disabled"
		IF SNAP_DISA_status = "11" THEN SNAP_DISA_status = "VA determined PD disa"
		IF SNAP_DISA_status = "12" THEN SNAP_DISA_status = "VA (other accept disa)"
		IF SNAP_DISA_status = "13" THEN SNAP_DISA_status = "Cert RR Ret Disa & on MEDI"
		IF SNAP_DISA_status = "14" THEN SNAP_DISA_status = "Other Govt Perm Disa Ret Bnft"
		IF SNAP_DISA_status = "15" THEN SNAP_DISA_status = "Disability from MINE list"
		IF SNAP_DISA_status = "16" THEN SNAP_DISA_status = "Unable to p&p own meal"
		IF SNAP_DISA_verif = "?" OR SNAP_DISA_verif = "N" THEN
			SNAP_DISA_verif = ", no proof provided"
		ELSE
			SNAP_DISA_verif = ""
		END IF
		
		'Reads and formats HC disa status/verif
		EMReadScreen HC_DISA_status, 2, 13, 59
		EMReadScreen HC_DISA_verif, 1, 13, 69
		If HC_DISA_status = "01" or HC_DISA_status = "02" or DISA_status = "03" or DISA_status = "04" then DISA_status = "RSDI/SSI certified"
		If HC_DISA_status = "06" then HC_DISA_status = "SMRT/SSA pends"
		If HC_DISA_status = "08" then HC_DISA_status = "Certified blind"
		If HC_DISA_status = "10" then HC_DISA_status = "Certified disabled"
		If HC_DISA_status = "11" then HC_DISA_status = "Spec cat- disa child"
		If HC_DISA_status = "20" then HC_DISA_status = "TEFRA- disabled"
		If HC_DISA_status = "21" then HC_DISA_status = "TEFRA- blind"
		If HC_DISA_status = "22" then HC_DISA_status = "MA-EPD"
		If HC_DISA_status = "23" then HC_DISA_status = "MA/waiver"
		If HC_DISA_status = "24" then HC_DISA_status = "SSA/SMRT appeal pends"
		If HC_DISA_status = "26" then HC_DISA_status = "SSA/SMRT disa deny"
		IF HC_DISA_verif = "?" OR HC_DISA_verif = "N" THEN
			HC_DISA_verif = ", no proof provided"
		ELSE
			HC_DISA_verif = ""
		END IF
		'cleaning to make variable to write
		IF CASH_DISA_status = "__" THEN 
			CASH_DISA_status = ""
		ELSE
			IF CASH_DISA_status = SNAP_DISA_status THEN
				SNAP_DISA_status = "__"
				CASH_DISA_status = "CASH/SNAP: " & CASH_DISA_status & " "
			ELSE	
				CASH_DISA_status = "CASH: " & CASH_DISA_status & " "
			END IF
		END IF
		IF SNAP_DISA_status = "__" THEN 
			SNAP_DISA_status = ""
		ELSE
			SNAP_DISA_status = "SNAP: " & SNAP_DISA_status & " "
		END IF
		IF HC_DISA_status = "__" THEN 
			HC_DISA_status = ""
		ELSE
			HC_DISA_status = "HC: " & HC_DISA_status & " "
		END IF
		'Adding verif code info if N or ?
		IF CASH_DISA_verif <> "" THEN CASH_DISA_status = CASH_DISA_status & CASH_DISA_verif & " "
		IF SNAP_DISA_verif <> "" THEN SNAP_DISA_status = SNAP_DISA_status & SNAP_DISA_verif & " "
		IF HC_DISA_verif <> "" THEN HC_DISA_status = HC_DISA_status & HC_DISA_verif & " "
		'Creating final variable
		IF CASH_DISA_status <> "" THEN FINAL_DISA_status = CASH_DISA_status
		IF SNAP_DISA_status <> "" THEN FINAL_DISA_status = FINAL_DISA_status & SNAP_DISA_status
		IF HC_DISA_status <> "" THEN FINAL_DISA_status = FINAL_DISA_status & HC_DISA_status
		
		variable_written_to = variable_written_to & "Member " & HH_member & "- "
		variable_written_to = variable_written_to & FINAL_DISA_status & "; "
	  END IF
    Next
  Elseif panel_read_from = "EATS" then '----------------------------------------------------------------------------------------------------EATS
    row = 14
    Do
      EMReadScreen reference_numbers_current_row, 40, row, 39
      reference_numbers = reference_numbers + reference_numbers_current_row  
      row = row + 1
    Loop until row = 18
    reference_numbers = replace(reference_numbers, "  ", " ")
    reference_numbers = split(reference_numbers)
    For each member in reference_numbers
      If member <> "__" and member <> "" then EATS_info = EATS_info & member & ", "
    Next
    EATS_info = trim(EATS_info)
    if right(EATS_info, 1) = "," then EATS_info = left(EATS_info, len(EATS_info) - 1)
    If EATS_info <> "" then variable_written_to = variable_written_to & ", p/p sep from memb(s) " & EATS_info & "."
  Elseif panel_read_from = "FACI" then '----------------------------------------------------------------------------------------------------FACI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen FACI_total, 1, 2, 78
      If FACI_total <> 0 then
        row = 14
        Do
          EMReadScreen date_in_check, 4, row, 53
		  EMReadScreen date_in_month_day, 5, row, 47
          EMReadScreen date_out_check, 4, row, 77
		  date_in_month_day = replace(date_in_month_day, " ", "/") & "/"
          If (date_in_check <> "____" and date_out_check <> "____") or (date_in_check = "____" and date_out_check = "____") then row = row + 1
          If row > 18 then
            EMReadScreen FACI_page, 1, 2, 73
            If FACI_page = FACI_total then 
              FACI_status = "Not in facility"
            Else
              transmit
              row = 14
            End if
          End if
        Loop until (date_in_check <> "____" and date_out_check = "____") or FACI_status = "Not in facility"
        EMReadScreen client_FACI, 30, 6, 43
        client_FACI = replace(client_FACI, "_", "")
        FACI_array = split(client_FACI)
        For each a in FACI_array
          If a <> "" then
            b = ucase(left(a, 1))
            c = LCase(right(a, len(a) -1))
            new_FACI = new_FACI & b & c & " "
          End if
        Next
        client_FACI = new_FACI
        If FACI_status = "Not in facility" then
          client_FACI = ""
        Else
          variable_written_to = variable_written_to & "Member " & HH_member & "- "
          variable_written_to = variable_written_to & client_FACI & " Date in: " & date_in_month_day & date_in_check & "; "
        End if
      End if
    Next
  Elseif panel_read_from = "FMED" then '----------------------------------------------------------------------------------------------------FMED
	For each HH_member in HH_member_array
	  ERRR_screen_check
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      fmed_row = 9 'Setting this variable for the next do...loop
      EMReadScreen fmed_total, 1, 2, 78
      If fmed_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
		  use_expense = False					'<--- Used to determine if an FMED expense that has an end date is going to be counted.
          EMReadScreen fmed_type, 2, fmed_row, 25
          EMReadScreen fmed_proof, 2, fmed_row, 32
          EMReadScreen fmed_amt, 8, fmed_row, 70
		  EMReadScreen fmed_end_date, 5, fmed_row, 60		'reading end date to see if this one even gets added.
		  IF fmed_end_date <> "__ __" THEN
			fmed_end_date = replace(fmed_end_date, " ", "/01/")
			fmed_end_date = dateadd("M", 1, fmed_end_date)
			fmed_end_date = dateadd("D", -1, fmed_end_date)
			IF datediff("D", date, fmed_end_date) > 0 THEN use_expense = True		'<--- If the end date of the FMED expense is the current month or a future month, the expense is going to be counted.
		  END IF
		  If fmed_end_date = "__ __" OR use_expense = TRUE then					'Skips entries with an end date or end dates in the past.
            If fmed_proof = "__" or fmed_proof = "?_" or fmed_proof = "NO" then 
              fmed_proof = ", no proof provided"
            Else
              fmed_proof = ""
            End if
            If fmed_amt = "________" then
              fmed_amt = ""
            Else
              fmed_amt = " ($" & trim(fmed_amt) & ")"
            End if
            If fmed_type = "01" then fmed_type = "Nursing Home"
            If fmed_type = "02" then fmed_type = "Hosp/Clinic"
            If fmed_type = "03" then fmed_type = "Physicians"
            If fmed_type = "04" then fmed_type = "Prescriptions"
            If fmed_type = "05" then fmed_type = "Ins Premiums"
            If fmed_type = "06" then fmed_type = "Dental"
            If fmed_type = "07" then fmed_type = "Medical Trans/Flat Amt"
            If fmed_type = "08" then fmed_type = "Vision Care"
            If fmed_type = "09" then fmed_type = "Medicare Prem"
            If fmed_type = "10" then fmed_type = "Mo. Spdwn Amt/Waiver Obl"
            If fmed_type = "11" then fmed_type = "Home Care"
            If fmed_type = "12" then fmed_type = "Medical Trans/Mileage Calc"
            If fmed_type = "15" then fmed_type = "Medi Part D premium"
            If fmed_type <> "__" then variable_written_to = variable_written_to & fmed_type & fmed_amt & fmed_proof & "; "
			IF fmed_end_date <> "__ __" THEN					'<--- If there is a counted FMED expense with a future end date, the script will modify the way that end date is displayed.
				fmed_end_date = datepart("M", fmed_end_date) & "/" & right(datepart("YYYY", fmed_end_date), 2)		'<--- Begins pulling apart fmed_end_date to format it to human speak.
				IF left(fmed_end_date, 1) <> "0" THEN fmed_end_date = "0" & fmed_end_date
				variable_written_to = left(variable_written_to, len(variable_written_to) - 2) & ", counted through " & fmed_end_date & "; "			'<--- Putting variable_written_to back together with FMED expense end date information.
			END IF	
          End if
          fmed_row = fmed_row + 1
          If fmed_row = 15 then
            PF20
            fmed_row = 9
            EMReadScreen last_page_check, 21, 24, 2
            If last_page_check <> "THIS IS THE LAST PAGE" then last_page_check = ""
          End if
        Loop until fmed_type = "__" or last_page_check = "THIS IS THE LAST PAGE"
      End if
    Next
  Elseif panel_read_from = "HCRE" then '----------------------------------------------------------------------------------------------------HCRE
    EMReadScreen variable_written_to, 8, 10, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If variable_written_to = "__/__/__" then EMReadScreen variable_written_to, 8, 11, 51
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then variable_written_to = cdate(variable_written_to) & ""
    If isdate(variable_written_to) = False then variable_written_to = ""
  Elseif panel_read_from = "HCRE-retro" then '----------------------------------------------------------------------------------------------HCRE-retro
    EMReadScreen variable_written_to, 5, 10, 64
    If isdate(variable_written_to) = True then
      variable_written_to = replace(variable_written_to, " ", "/01/")
      If DatePart("m", variable_written_to) <> DatePart("m", CAF_datestamp) or DatePart("yyyy", variable_written_to) <> DatePart("yyyy", CAF_datestamp) then
        variable_written_to = variable_written_to
      Else
        variable_written_to = ""
      End if
    End if
  Elseif panel_read_from = "HEST" then '----------------------------------------------------------------------------------------------------HEST
    EMReadScreen HEST_total, 1, 2, 78
    If HEST_total <> 0 then 
      EMReadScreen heat_air_check, 6, 13, 75
      If heat_air_check <> "      " then variable_written_to = variable_written_to & "Heat/AC.; "
      EMReadScreen electric_check, 6, 14, 75
      If electric_check <> "      " then variable_written_to = variable_written_to & "Electric.; "
      EMReadScreen phone_check, 6, 15, 75
      If phone_check <> "      " then variable_written_to = variable_written_to & "Phone.; "
    End if
  Elseif panel_read_from = "IMIG" then '----------------------------------------------------------------------------------------------------IMIG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen IMIG_total, 1, 2, 78
      If IMIG_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen IMIG_type, 30, 6, 48
        variable_written_to = variable_written_to & trim(IMIG_type) & "; "
      End if
    Next
  Elseif panel_read_from = "INSA" then '----------------------------------------------------------------------------------------------------INSA
    EMReadScreen INSA_amt, 1, 2, 78
    If INSA_amt <> 0 then
      'Runs once per INSA screen
		For i = 1 to INSA_amt step 1
			insurance_name = ""
			'Goes to the correct screen
			EMWriteScreen "0" & i, 20, 79
			transmit
			'Gather Insurance Name
			EMReadScreen INSA_name, 38, 10, 38
			INSA_name = replace(INSA_name, "_", "")
			INSA_name = split(INSA_name)
			For each word in INSA_name
				If trim(word) <> "" then
						first_letter_of_word = ucase(left(word, 1))
						rest_of_word = LCase(right(word, len(word) -1))
						If len(word) > 4 then
							insurance_name = insurance_name & first_letter_of_word & rest_of_word & " "
						Else
							insurance_name = insurance_name & word & " "
						End if
				End if
			Next
			'Create a list of members covered by this insurance
			y = 15 : x = 30
			insured_count = 0
			member_list = ""
			Do
				EMReadScreen insured_member, 2, y, x
				If insured_member <> "__" then 
					if member_list = "" then member_list = insured_member
					if member_list <> "" then member_list = member_list & ", " & insured_member
					x = x + 4
					If x = 70 then
						x = 30 : y = 16
					End If
				End If
			loop until insured_member = "__"
			'Retain "variable_written_to" as is while also adding members covered by the insurance policy
			'Example - "Members: 01, 03, 07 are covered by Blue Cross Blue Shield; " 
			variable_written_to = variable_written_to & "Members: " & member_list & " are covered by " & trim(insurance_name) & "; "
		Next
		'This will loop and add the above statement for all insurance policies listed
	End if
  Elseif panel_read_from = "JOBS" then '----------------------------------------------------------------------------------------------------JOBS
	For each HH_member in HH_member_array  
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen JOBS_total, 1, 2, 78
      If JOBS_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_JOBS_to_variable(variable_written_to)
          EMReadScreen JOBS_panel_current, 1, 2, 73
          If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
        Loop until cint(JOBS_panel_current) = cint(JOBS_total)
      End if
    Next
  Elseif panel_read_from = "MEDI" then '----------------------------------------------------------------------------------------------------MEDI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen MEDI_amt, 1, 2, 78
      If MEDI_amt <> "0" then variable_written_to = variable_written_to & "Medicare for member " & HH_member & ".; "
    Next
  Elseif panel_read_from = "MEMB" then '----------------------------------------------------------------------------------------------------MEMB
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen rel_to_applicant, 2, 10, 42
      EMReadScreen client_age, 3, 8, 76
      If client_age = "   " then client_age = 0
      If cint(client_age) >= 21 or rel_to_applicant = "02" then
        number_of_adults = number_of_adults + 1
      Else
        number_of_children = number_of_children + 1
      End if
    Next
    If number_of_adults > 0 then variable_written_to = number_of_adults & "a"
    If number_of_children > 0 then variable_written_to = variable_written_to & ", " & number_of_children & "c"
    If left(variable_written_to, 1) = "," then variable_written_to = right(variable_written_to, len(variable_written_to) - 1)
  Elseif panel_read_from = "MEMI" then '----------------------------------------------------------------------------------------------------MEMI
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen citizen, 1, 10, 49
      If citizen = "Y" then citizen = "US citizen"
      If citizen = "N" then citizen = "non-citizen"
      EMReadScreen citizenship_ver, 2, 10, 78
      EMReadScreen SSA_MA_citizenship_ver, 1, 11, 49
      If citizenship_ver = "__" or citizenship_ver = "NO" then cit_proof_indicator = ", no verifs provided"
      If SSA_MA_citizenship_ver = "R" then cit_proof_indicator = ", MEMI infc req'd"
      If (citizenship_ver <> "__" and citizenship_ver <> "NO") or (SSA_MA_citizenship_ver = "A") then cit_proof_indicator = ""
      variable_written_to = variable_written_to & "Member " & HH_member & "- "
      variable_written_to = variable_written_to & citizen & cit_proof_indicator & "; "
    Next
  ElseIf panel_read_from = "MONT" then '----------------------------------------------------------------------------------------------------MONT
    EMReadScreen variable_written_to, 8, 6, 39
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "OTHR" then '----------------------------------------------------------------------------------------------------OTHR
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen OTHR_total, 1, 2, 78
      If OTHR_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_OTHR_to_variable(variable_written_to)
          EMReadScreen OTHR_panel_current, 1, 2, 73
          If cint(OTHR_panel_current) < cint(OTHR_total) then transmit
        Loop until cint(OTHR_panel_current) = cint(OTHR_total)
      End if
    Next
  Elseif panel_read_from = "PBEN" then '----------------------------------------------------------------------------------------------------PBEN
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      transmit
      EMReadScreen panel_amt, 1, 2, 78
      If panel_amt <> "0" then
        PBEN = PBEN & "Member " & HH_member & "- "
        row = 8
        Do
          EMReadScreen PBEN_type, 12, row, 28
          EMReadScreen PBEN_disp, 1, row, 77
          If PBEN_disp = "A" then PBEN_disp = " appealing"
          If PBEN_disp = "D" then PBEN_disp = " denied"
          If PBEN_disp = "E" then PBEN_disp = " eligible"
          If PBEN_disp = "P" then PBEN_disp = " pends"
          If PBEN_disp = "N" then PBEN_disp = " not applied yet"
          If PBEN_disp = "R" then PBEN_disp = " refused"
          If PBEN_type <> "            " then PBEN = PBEN & trim(PBEN_type) & PBEN_disp & "; "
          row = row + 1
        Loop until row = 14
      End if
    Next
    If PBEN <> "" then variable_written_to = variable_written_to & PBEN
  Elseif panel_read_from = "PREG" then '----------------------------------------------------------------------------------------------------PREG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen PREG_total, 1, 2, 78
      If PREG_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen PREG_due_date, 8, 10, 53
        If PREG_due_date = "__ __ __" then
          PREG_due_date = "unknown"
        Else
          PREG_due_date = replace(PREG_due_date, " ", "/")
        End if
        variable_written_to = variable_written_to & "Due date is " & PREG_due_date & ".; "
      End if
    Next
  Elseif panel_read_from = "PROG" then '----------------------------------------------------------------------------------------------------PROG
    row = 6
    Do
      EMReadScreen appl_prog_date, 8, row, 33
      If appl_prog_date <> "__ __ __" then appl_prog_date_array = appl_prog_date_array & replace(appl_prog_date, " ", "/") & " "
      row = row + 1
    Loop until row = 13
    appl_prog_date_array = split(appl_prog_date_array)
    variable_written_to = CDate(appl_prog_date_array(0))
    for i = 0 to ubound(appl_prog_date_array) - 1
      if CDate(appl_prog_date_array(i)) > variable_written_to then 
        variable_written_to = CDate(appl_prog_date_array(i))
      End if
    next
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "RBIC" then '----------------------------------------------------------------------------------------------------RBIC
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen RBIC_total, 1, 2, 78
      If RBIC_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_RBIC_to_variable(variable_written_to)
          EMReadScreen RBIC_panel_current, 1, 2, 73
          If cint(RBIC_panel_current) < cint(RBIC_total) then transmit
        Loop until cint(RBIC_panel_current) = cint(RBIC_total)
      End if
    Next
  Elseif panel_read_from = "REST" then '----------------------------------------------------------------------------------------------------REST
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen REST_total, 1, 2, 78
      If REST_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_REST_to_variable(variable_written_to)
          EMReadScreen REST_panel_current, 1, 2, 73
          If cint(REST_panel_current) < cint(REST_total) then transmit
        Loop until cint(REST_panel_current) = cint(REST_total)
      End if
    Next
  Elseif panel_read_from = "REVW" then '----------------------------------------------------------------------------------------------------REVW
    EMReadScreen variable_written_to, 8, 13, 37
    variable_written_to = replace(variable_written_to, " ", "/")
    If isdate(variable_written_to) = True then
      variable_written_to = cdate(variable_written_to) & ""
    Else
      variable_written_to = ""
    End if
  Elseif panel_read_from = "SCHL" then '----------------------------------------------------------------------------------------------------SCHL
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen school_type, 2, 7, 40
      If school_type = "01" then school_type = "elementary school"
      If school_type = "11" then school_type = "middle school"
      If school_type = "02" then school_type = "high school"
      If school_type = "03" then school_type = "GED"
      If school_type = "07" then school_type = "IEP"
      If school_type = "08" or school_type = "09" or school_type = "10" then school_type = "post-secondary"
      If school_type = "06" or school_type = "__" or school_type = "?_" then
        school_type = ""
      Else
        EMReadScreen SCHL_ver, 2, 6, 63
        If SCHL_ver = "?_" or SCHL_ver = "NO" then
          school_proof_type = ", no proof provided"
        Else
          school_proof_type = ""
        End if
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        variable_written_to = variable_written_to & school_type & school_proof_type & "; "
      End if
    Next
  Elseif panel_read_from = "SECU" then '----------------------------------------------------------------------------------------------------SECU
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SECU_total, 1, 2, 78
      If SECU_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_SECU_to_variable(variable_written_to)
          EMReadScreen SECU_panel_current, 1, 2, 73
          If cint(SECU_panel_current) < cint(SECU_total) then transmit
        Loop until cint(SECU_panel_current) = cint(SECU_total)
      End if
    Next
  Elseif panel_read_from = "SHEL" then '----------------------------------------------------------------------------------------------------SHEL
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen SHEL_total, 1, 2, 78
      If SHEL_total <> 0 then 
        member_number_designation = "Member " & HH_member & "- "
        row = 11
        Do
          EMReadScreen SHEL_amount, 8, row, 56
          If SHEL_amount <> "________" then
            EMReadScreen SHEL_type, 9, row, 24
            EMReadScreen SHEL_proof_check, 2, row, 67
            If SHEL_proof_check = "NO" or SHEL_proof_check = "?_" then 
              SHEL_proof = ", no proof provided"
            Else
              SHEL_proof = ""
            End if
            SHEL_expense = SHEL_expense & "$" & trim(SHEL_amount) & "/mo " & lcase(trim(SHEL_type)) & SHEL_proof & ". ;"
          End if
          row = row + 1
        Loop until row = 19
        variable_written_to = variable_written_to & member_number_designation & SHEL_expense
      End if
      SHEL_expense = ""
    Next
  Elseif panel_read_from = "STWK" then '----------------------------------------------------------------------------------------------------STWK
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen STWK_total, 1, 2, 78
      If STWK_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        EMReadScreen STWK_verification, 1, 7, 63
        If STWK_verification = "N" then
          STWK_verification = ", no proof provided"
        Else
          STWK_verification = ""
        End if
        EMReadScreen STWK_employer, 30, 6, 46
        STWK_employer = replace(STWK_employer, "_", "")
        STWK_employer = split(STWK_employer)
        For each a in STWK_employer
          If a <> "" then
            b = ucase(left(a, 1))
            c = LCase(right(a, len(a) -1))
            If len(a) > 3 then
              new_STWK_employer = new_STWK_employer & b & c & " "
            Else
              new_STWK_employer = new_STWK_employer & a & " "
            End if
          End if
        Next
        EMReadScreen STWK_income_stop_date, 8, 8, 46
        If STWK_income_stop_date = "__ __ __" then
          STWK_income_stop_date = "at unknown date"
        Else
          STWK_income_stop_date = replace(STWK_income_stop_date, " ", "/")
        End if
      EMReadScreen voluntary_quit, 1, 10, 46
	vol_quit_info = ", Vol. Quit " & voluntary_quit
	  IF voluntary_quit = "Y" THEN
		EMReadScreen good_cause, 1, 12, 67
		EMReadScreen fs_pwe, 1, 14, 46
		vol_quit_info = ", Vol Quit " & voluntary_quit & ", Good Cause " & good_cause & ", FS PWE " & fs_pwe
	  END IF
        variable_written_to = variable_written_to & new_STWK_employer & "income stopped " & STWK_income_stop_date & STWK_verification & vol_quit_info & ".; "
      End if
      new_STWK_employer = "" 'clearing variable to prevent duplicates
    Next
  Elseif panel_read_from = "UNEA" then '----------------------------------------------------------------------------------------------------UNEA
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
      EMReadScreen UNEA_total, 1, 2, 78
      If UNEA_total <> 0 then 
        variable_written_to = variable_written_to & "Member " & HH_member & "- "
        Do
          call add_UNEA_to_variable(variable_written_to)
          EMReadScreen UNEA_panel_current, 1, 2, 73
          If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
        Loop until cint(UNEA_panel_current) = cint(UNEA_total)
      End if
    Next
  Elseif panel_read_from = "WREG" then '---------------------------------------------------------------------------------------------------WREG
	For each HH_member in HH_member_array
      EMWriteScreen HH_member, 20, 76
      EMWriteScreen "01", 20, 79
      transmit
    EMReadScreen wreg_total, 1, 2, 78
    EMReadScreen snap_case_yn, 1, 6, 50
    IF wreg_total <> "0" and snap_case_yn = "Y" THEN 
	EmWriteScreen "x", 13, 57
	transmit
	 bene_mo_col = (15 + (4*cint(footer_month)))
	  bene_yr_row = 10
       abawd_counted_months = 0
       second_abawd_period = 0
 	 month_count = 0
 	   DO
  		  EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
  		    IF is_counted_month = "X" or is_counted_month = "M" THEN abawd_counted_months = abawd_counted_months + 1
		    IF is_counted_month = "Y" or is_counted_month = "N" THEN second_abawd_period = second_abawd_period + 1
   		  bene_mo_col = bene_mo_col - 4
    		    IF bene_mo_col = 15 THEN
        		bene_yr_row = bene_yr_row - 1
   	     		bene_mo_col = 63
   	   	    END IF
    		  month_count = month_count + 1
  	   LOOP until month_count = 36
  	PF3
	EmreadScreen read_WREG_status, 2, 8, 50
	If read_WREG_status = 03 THEN  WREG_status = "WREG = incap"
	If read_WREG_status = 04 THEN  WREG_status = "WREG = resp for incap HH memb"
	If read_WREG_status = 05 THEN  WREG_status = "WREG = age 60+"
	If read_WREG_status = 06 THEN  WREG_status = "WREG = < age 16"
	If read_WREG_status = 07 THEN  WREG_status = "WREG = age 16-17, live w/prnt/crgvr"
	If read_WREG_status = 08 THEN  WREG_status = "WREG = resp for child < 6 yrs old"
	If read_WREG_status = 09 THEN  WREG_status = "WREG = empl 30 hrs/wk or equiv"
	If read_WREG_status = 10 THEN  WREG_status = "WREG = match grant part"
	If read_WREG_status = 11 THEN  WREG_status = "WREG = rec/app for unemp ins"
	If read_WREG_status = 12 THEN  WREG_status = "WREG = in schl, train prog or higher ed"
	If read_WREG_status = 13 THEN  WREG_status = "WREG = in CD prog"
	If read_WREG_status = 14 THEN  WREG_status = "WREG = rec MFIP"
	If read_WREG_status = 20 THEN  WREG_status = "WREG = pend/rec DWP or WB"
	If read_WREG_status = 22 THEN  WREG_status = "WREG = app for SSI"
	If read_WREG_status = 15 THEN  WREG_status = "WREG = age 16-17 not live w/ prnt/crgvr"
	If read_WREG_status = 16 THEN  WREG_status = "WREG = 50-59 yrs old"
	If read_WREG_status = 21 THEN  WREG_status = "WREG = resp for child < 18"
	If read_WREG_status = 17 THEN  WREG_status = "WREG = rec RCA or GA"
	If read_WREG_status = 18 THEN  WREG_status = "WREG = provide home schl"
	If read_WREG_status = 30 THEN  WREG_status = "WREG = mand FSET part"
	If read_WREG_status = 02 THEN  WREG_status = "WREG = non-coop w/ FSET"
	If read_WREG_status = 33 THEN  WREG_status = "WREG = non-coop w/ referral"
	If read_WREG_status = "__" THEN  WREG_status = "WREG = blank"
	
	EmreadScreen read_abawd_status, 2, 13, 50
	If read_abawd_status = 01 THEN  abawd_status = "ABAWD = work reg exempt."
      If read_abawd_status = 02 THEN  abawd_status = "ABAWD = < age 18."
	If read_abawd_status = 03 THEN  abawd_status = "ABAWD = age 50+."
	If read_abawd_status = 04 THEN  abawd_status = "ABAWD = crgvr of minor child."		
	If read_abawd_status = 05 THEN  abawd_status = "ABAWD = pregnant."
	If read_abawd_status = 06 THEN  abawd_status = "ABAWD = emp ave 20 hrs/wk."
	If read_abawd_status = 07 THEN  abawd_status = "ABAWD = work exp participant."	
	If read_abawd_status = 08 THEN  abawd_status = "ABAWD = othr E & T service."
	If read_abawd_status = 09 THEN  abawd_status = "ABAWD = reside in waiver area."
	If read_abawd_status = 10 THEN  abawd_status = "ABAWD = ABAWD & has used " & abawd_counted_months & " mo"
	If read_abawd_status = 11 THEN  abawd_status = "ABAWD = using 2nd three mo period of elig."
	If read_abawd_status = 12 THEN  abawd_status = "ABAWD = RCA or GA recip."
	If read_abawd_status = 13 THEN  abawd_status = "ABAWD = ABAWD extension."
	If read_abawd_status = "__" THEN  abawd_status = "WREG = blank"


	variable_written_to = variable_written_to & "Member " & HH_member & "- " & WREG_status & ", " & abawd_status & "; "
     END IF
    Next
  End if
  variable_written_to = trim(variable_written_to) '-----------------------------------------------------------------------------------------cleaning up editbox
  if right(variable_written_to, 1) = ";" then variable_written_to = left(variable_written_to, len(variable_written_to) - 1)


End function

function back_to_SELF
  Do
    EMSendKey "<PF3>"
    EMWaitReady 0, 0
    EMReadScreen SELF_check, 4, 2, 50
  Loop until SELF_check = "SELF"
End function

'This function asks if you want to cancel. If you say yes, it sends StopScript.
FUNCTION cancel_confirmation
	If ButtonPressed = 0 then 
		cancel_confirm = MsgBox("Are you sure you want to cancel the script? Press YES to cancel. Press NO to return to the script.", vbYesNo)
		If cancel_confirm = vbYes then stopscript
	End if
END FUNCTION

Function check_for_MAXIS(end_script)
	Do
		transmit
		EMReadScreen MAXIS_check, 5, 1, 39
		If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then 
			If end_script = True then 
				script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
			Else
				warning_box = MsgBox("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
				If warning_box = vbCancel then stopscript
			End if
		End if
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
End function

Function check_for_PRISM(end_script)
	EMReadScreen PRISM_check, 5, 1, 36
	if end_script = True then
		If PRISM_check <> "PRISM" then script_end_procedure("You do not appear to be in PRISM. You may be passworded out. Please check your PRISM screen and try again.")
	else
		If PRISM_check <> "PRISM" then MsgBox "You do not appear to be in PRISM. You may be passworded out. Please enter your password before pressing OK."
	end if
end function

'This function converts an array into a droplist to be used by a dialog
Function convert_array_to_droplist_items(array_to_convert, output_droplist_box)
	For each item in array_to_convert
		If output_droplist_box = "" then 
			output_droplist_box = item
		Else
			output_droplist_box = output_droplist_box & chr(9) & item
		End if
	Next
End Function

'This function converts a date (MM/DD/YY or MM/DD/YYYY format) into a separate footer month and footer year variables. For best results, always use footer_month and footer_year as the appropriate variables.
FUNCTION convert_date_into_MAXIS_footer_month(date_to_convert, footer_month, footer_year)
	footer_month = DatePart("m", date_to_convert)						'Uses DatePart function to copy the month from date_to_convert into the footer_month variable.
	IF Len(footer_month) = 1 THEN footer_month = "0" & footer_month		'Uses Len function to determine if the footer_month is a single digit month. If so, it adds a 0, which MAXIS needs.
	footer_year = DatePart("yyyy", date_to_convert)						'Uses DatePart function to copy the year from date_to_convert into the footer_year variable.
	footer_year = Right(footer_year, 2)									'Uses Right function to reduce the footer_year variable to it's right 2 characters (allowing for a 2 digit footer year).
END FUNCTION

'This function converts a numeric digit to an Excel column, up to 104 digits (columns).
function convert_digit_to_excel_column(col_in_excel)
	'Create string with the alphabet
	alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

	'Assigning a letter, based on that column. Uses "mid" function to determine it. If number > 26, it handles by adding a letter (per Excel).
	convert_digit_to_excel_column = Mid(alphabet, col_in_excel, 1)		
	If col_in_excel >= 27 and col_in_excel < 53 then convert_digit_to_excel_column = "A" & Mid(alphabet, col_in_excel - 26, 1)
	If col_in_excel >= 53 and col_in_excel < 79 then convert_digit_to_excel_column = "B" & Mid(alphabet, col_in_excel - 52, 1)
	If col_in_excel >= 79 and col_in_excel < 105 then convert_digit_to_excel_column = "C" & Mid(alphabet, col_in_excel - 78, 1)

	'Closes script if the number gets too high (very rare circumstance, just errorproofing)
	If col_in_excel >= 105 then script_end_procedure("This script is only able to assign excel columns to 104 distinct digits. You've exceeded this number, and this script cannot continue.")
end function

Function create_array_of_all_active_x_numbers_in_county(array_name, county_code)
	'Getting to REPT/USER
	call navigate_to_screen("rept", "user")

	'Hitting PF5 to force sorting, which allows directly selecting a county
	PF5

	'Inserting county
	EMWriteScreen county_code, 21, 6
	transmit

	'Declaring the MAXIS row
	MAXIS_row = 7

	'Blanking out array_name in case this has been used already in the script
	array_name = ""

	Do
		Do
			'Reading MAXIS information for this row, adding to spreadsheet
			EMReadScreen worker_ID, 8, MAXIS_row, 5					'worker ID
			If worker_ID = "        " then exit do					'exiting before writing to array, in the event this is a blank (end of list)
			array_name = trim(array_name & " " & worker_ID)				'writing to variable
			MAXIS_row = MAXIS_row + 1
		Loop until MAXIS_row = 19

		'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
		EMReadScreen more_pages_check, 7, 19, 3
		If more_pages_check = "More: +" then 
			PF8			'getting to next screen
			MAXIS_row = 7	'redeclaring MAXIS row so as to start reading from the top of the list again
		End if
	Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank
	array_name = split(array_name)
End function

'Creates a MM DD YY date entry at screen_row and screen_col. The variable_length variable is the amount of days to offset the date entered. I.e., 10 for 10 days, -10 for 10 days in the past, etc.
Function create_MAXIS_friendly_date(date_variable, variable_length, screen_row, screen_col) 
	var_month = datepart("m", dateadd("d", variable_length, date_variable))
	If len(var_month) = 1 then var_month = "0" & var_month
	EMWriteScreen var_month, screen_row, screen_col
	var_day = datepart("d", dateadd("d", variable_length, date_variable))
	If len(var_day) = 1 then var_day = "0" & var_day
	EMWriteScreen var_day, screen_row, screen_col + 3
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
	EMWriteScreen right(var_year, 2), screen_row, screen_col + 6
End function

FUNCTION create_MAXIS_friendly_date_three_spaces_between(date_variable, variable_length, screen_row, screen_col) 
	var_month = datepart("m", dateadd("d", variable_length, date_variable))		'determines the date based on the variable length: month 
	If len(var_month) = 1 then var_month = "0" & var_month				'adds a '0' in front of a single digit month
	EMWriteScreen var_month, screen_row, screen_col					'writes in var_month at coordinates set in FUNCTION line
	var_day = datepart("d", dateadd("d", variable_length, date_variable)) 		'determines the date based on the variable length: day
	If len(var_day) = 1 then var_day = "0" & var_day 				'adds a '0' in front of a single digit day
	EMWriteScreen var_day, screen_row, screen_col + 5 				'writes in var_day at coordinates set in FUNCTION line, and starts 5 columns into date field in MAXIS
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable)) 	'determines the date based on the variable length: year
	EMWriteScreen right(var_year, 2), screen_row, screen_col + 10 			'writes in var_year at coordinates set in FUNCTION line , and starts 5 columns into date field in MAXIS
END FUNCTION

'Creates a MM DD YYYY date entry at screen_row and screen_col. The variable_length variable is the amount of days to offset the date entered. I.e., 10 for 10 days, -10 for 10 days in the past, etc.
FUNCTION create_MAXIS_friendly_date_with_YYYY(date_variable, variable_length, screen_row, screen_col) 
	var_month = datepart("m", dateadd("d", variable_length, date_variable))
	IF len(var_month) = 1 THEN var_month = "0" & var_month
	EMWriteScreen var_month, screen_row, screen_col
	var_day = datepart("d", dateadd("d", variable_length, date_variable))
	IF len(var_day) = 1 THEN var_day = "0" & var_day
	EMWriteScreen var_day, screen_row, screen_col + 3
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
	EMWriteScreen var_year, screen_row, screen_col + 6
END FUNCTION

FUNCTION create_MAXIS_friendly_phone_number(phone_number_variable, screen_row, screen_col)
	WITH (new RegExp)                                                            	'Uses RegExp to bring in special string functions to remove the unneeded strings
                .Global = True                                                   	'I don't know what this means but David made it work so we're going with it
                .Pattern = "\D"                                                	 	'Again, no clue. Just do it.
                phone_number_variable = .Replace(phone_number_variable, "")    	 	'This replaces the non-digits of the phone number with nothing. That leaves us with a bunch of numbers
	END WITH
	EMWriteScreen left(phone_number_variable, 3), screen_row, screen_col 		'writes in left 3 digits of the phone number in variable
	EMWriteScreen mid(phone_number_variable, 4, 3), screen_row, screen_col + 6	'writes in middle 3 digits of the phone number in variable
	EMWriteScreen right(phone_number_variable, 4), screen_row, screen_col + 12	'writes in right 4 digits of the phone number in variable
END FUNCTION

Function create_panel_if_nonexistent()
	EMWriteScreen reference_number , 20, 76
	transmit
	EMReadScreen case_panel_check, 44, 24, 2
	If case_panel_check = "REFERENCE NUMBER IS NOT VALID FOR THIS PANEL" then
		EMReadScreen quantity_of_screens, 1, 2, 78
		If quantity_of_screens <> "0" then
			PF9
		ElseIf quantity_of_screens = "0" then
			EMWriteScreen "__", 20, 76
			EMWriteScreen "NN", 20, 79
			Transmit
		End If
	ElseIf case_panel_check <> "REFERENCE NUMBER IS NOT VALID FOR THIS PANEL" then
		EMReadScreen error_scan, 80, 24, 1
		error_scan = trim(error_scan)
		EMReadScreen quantity_of_screens, 1, 2, 78
		If error_scan = "" and quantity_of_screens <> "0" then
			PF9
		ElseIf error_scan <> "" and quantity_of_screens <> "0" then
			'FIX ERROR HERE
			msgbox("Error: " & error_scan)
		ElseIf error_scan <> "" and quantity_of_screens = "0" then
			EMWriteScreen reference_number, 20, 76
			EMWriteScreen "NN", 20, 79
			Transmit
		End If
	End If
End Function

Function end_excel_and_script
  objExcel.Workbooks.Close
  objExcel.quit
  stopscript
End function

Function find_variable(x, y, z) 'x is string, y is variable, z is length of new variable
  row = 1
  col = 1
  EMSearch x, row, col
  If row <> 0 then EMReadScreen y, z, row, col + len(x)
End function

'This function fixes the case for a phrase. For example, "ROBERT P. ROBERTSON" becomes "Robert P. Robertson". 
'	It capitalizes the first letter of each word.
Function fix_case(phrase_to_split, smallest_length_to_skip)										'Ex: fix_case(client_name, 3), where 3 means skip words that are 3 characters or shorter
	phrase_to_split = split(phrase_to_split)													'splits phrase into an array
	For each word in phrase_to_split															'processes each word independently
		If word <> "" then																		'Skip blanks
			first_character = ucase(left(word, 1))												'grabbing the first character of the string, making uppercase and adding to variable
			remaining_characters = LCase(right(word, len(word) -1))								'grabbing the remaining characters of the string, making lowercase and adding to variable
			If len(word) > smallest_length_to_skip then											'skip any strings shorter than the smallest_length_to_skip variable
				output_phrase = output_phrase & first_character & remaining_characters & " "	'output_phrase is the output of the function, this combines the first_character and remaining_characters
			Else															
				output_phrase = output_phrase & word & " "										'just pops the whole word in if it's shorter than the smallest_length_to_skip variable
			End if
		End if
	Next
	phrase_to_split = output_phrase																'making the phrase_to_split equal to the output, so that it can be used by the rest of the script.
End function

Function get_to_MMIS_session_begin
  Do 
    EMSendkey "<PF6>"
    EMWaitReady 0, 0
    EMReadScreen session_start, 18, 1, 7
  Loop until session_start = "SESSION TERMINATED"
End function

Function MAXIS_background_check
	Do
		call navigate_to_screen("STAT", "SUMM")
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then
			PF3
			Pause 2
		End if
	Loop until SELF_check <> "SELF"
End function

Function MAXIS_case_number_finder(variable_for_MAXIS_case_number)
	row = 1
	col = 1
	EMSearch "Case Nbr:", row, col
	If row <> 0 then 
		EMReadScreen variable_for_MAXIS_case_number, 8, row, col + 10
		variable_for_MAXIS_case_number = replace(variable_for_MAXIS_case_number, "_", "")
		variable_for_MAXIS_case_number = trim(variable_for_MAXIS_case_number)
	End if
End function

Function HH_member_custom_dialog(HH_member_array)

CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name. 

DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadscreen last_name, 5, 6, 30
	EMReadscreen first_name, 7, 6, 63
	EMReadscreen Mid_intial, 1, 6, 79
	last_name = replace(last_name, "_", "") & " "
	first_name = replace(first_name, "_", "") & " "
	mid_initial = replace(mid_initial, "_", "")
	client_string = ref_nbr & last_name & first_name & mid_intial
	client_array = client_array & client_string & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2	
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row. 

client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

DIM all_client_array()
ReDim all_clients_array(total_clients, 1)

FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(x, 0) = Interim_array(x)
	all_clients_array(x, 1) = 1
NEXT

BEGINDIALOG HH_memb_dialog, 0, 0, 191, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
	Text 10, 5, 105, 10, "Household members to look at:"						
	FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
		IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 120, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
	NEXT
	ButtonGroup ButtonPressed
	OkButton 135, 10, 50, 15
	CancelButton 135, 30, 50, 15
ENDDIALOG
													'runs the dialog that has been dynamically created. Streamlined with new functions.
Dialog HH_memb_dialog
If buttonpressed = 0 then stopscript
check_for_maxis(True)

HH_member_array = ""					

FOR i = 0 to total_clients
	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts. 
		IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
			'msgbox all_clients_
			HH_member_array = HH_member_array & left(all_clients_array(i, 0), 2) & " "
		END IF
	END IF
NEXT

HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
HH_member_array = SPLIT(HH_member_array, " ")

End function

function log_usage_stats_without_closing 'For use when logging usage stats but then running another script, i.e. DAIL scrubber
	stop_time = timer
	script_run_time = stop_time - start_time
	If is_county_collecting_stats = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork") 
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Opening DB
		objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\DHS-MAXIS-Scripts\Statistics\usage statistics.accdb"

		'Opening usage_log and adding a record
		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & "" & "')", objConnection, adOpenStatic, adLockOptimistic
	End if
end function

'This function navigates to various panels in MAXIS. You need to name your buttons using the button names in the function.
FUNCTION MAXIS_dialog_navigation
	'This part works with the prev/next buttons on several of our dialogs. You need to name your buttons prev_panel_button, next_panel_button, prev_memb_button, and next_memb_button in order to use them.
	EMReadScreen STAT_check, 4, 20, 21
	If STAT_check = "STAT" then
		If ButtonPressed = prev_panel_button then 
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel = 1 then new_panel = current_panel
			If current_panel > 1 then new_panel = current_panel - 1
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = next_panel_button then 
			EMReadScreen current_panel, 1, 2, 73
			EMReadScreen amount_of_panels, 1, 2, 78
			If current_panel < amount_of_panels then new_panel = current_panel + 1
			If current_panel = amount_of_panels then new_panel = current_panel
			If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
			transmit
		ELSEIF ButtonPressed = prev_memb_button then 
			HH_memb_row = HH_memb_row - 1
			EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
			If isnumeric(prev_HH_memb) = False then
				HH_memb_row = HH_memb_row + 1
			Else
				EMWriteScreen prev_HH_memb, 20, 76
				EMWriteScreen "01", 20, 79
			End if
			transmit
		ELSEIF ButtonPressed = next_memb_button then 
			HH_memb_row = HH_memb_row + 1
			EMReadScreen next_HH_memb, 2, HH_memb_row, 3
			If isnumeric(next_HH_memb) = False then
				HH_memb_row = HH_memb_row + 1
			Else
				EMWriteScreen next_HH_memb, 20, 76
				EMWriteScreen "01", 20, 79
			End if
			transmit
		End if
	End if
	
	'This part takes care of remaining navigation buttons, designed to go to a single panel.
	If ButtonPressed = ABPS_button then call navigate_to_screen("stat", "ABPS")
	If ButtonPressed = ACCI_button then call navigate_to_screen("stat", "ACCI")
	If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
	If ButtonPressed = ADDR_button then call navigate_to_screen("stat", "ADDR")
	If ButtonPressed = ALTP_button then call navigate_to_screen("stat", "ALTP")
	If ButtonPressed = AREP_button then call navigate_to_screen("stat", "AREP")
	If ButtonPressed = BILS_button then call navigate_to_screen("stat", "BILS")
	If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
	If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
	If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
	If ButtonPressed = COEX_button then call navigate_to_screen("stat", "COEX")
	If ButtonPressed = DCEX_button then call navigate_to_screen("stat", "DCEX")
	If ButtonPressed = DIET_button then call navigate_to_screen("stat", "DIET")
	If ButtonPressed = DISA_button then call navigate_to_screen("stat", "DISA")
	If ButtonPressed = EATS_button then call navigate_to_screen("stat", "EATS")
	If ButtonPressed = ELIG_DWP_button then call navigate_to_screen("elig", "DWP_")
	If ButtonPressed = ELIG_FS_button then call navigate_to_screen("elig", "FS__")
	If ButtonPressed = ELIG_GA_button then call navigate_to_screen("elig", "GA__")
	If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
	If ButtonPressed = ELIG_MFIP_button then call navigate_to_screen("elig", "MFIP")
	If ButtonPressed = ELIG_MSA_button then call navigate_to_screen("elig", "MSA_")
	If ButtonPressed = ELIG_WB_button then call navigate_to_screen("elig", "WB__")
	If ButtonPressed = FACI_button then call navigate_to_screen("stat", "FACI")
	If ButtonPressed = FMED_button then call navigate_to_screen("stat", "FMED")
	If ButtonPressed = HCRE_button then call navigate_to_screen("stat", "HCRE")
	If ButtonPressed = HEST_button then call navigate_to_screen("stat", "HEST")
	If ButtonPressed = IMIG_button then call navigate_to_screen("stat", "IMIG")
	If ButtonPressed = INSA_button then call navigate_to_screen("stat", "INSA")
	If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
	If ButtonPressed = MEDI_button then call navigate_to_screen("stat", "MEDI")
	If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
	If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
	If ButtonPressed = MONT_button then call navigate_to_screen("stat", "MONT")
	If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
	If ButtonPressed = PBEN_button then call navigate_to_screen("stat", "PBEN")
	If ButtonPressed = PDED_button then call navigate_to_screen("stat", "PDED")
	If ButtonPressed = PREG_button then call navigate_to_screen("stat", "PREG")
	If ButtonPressed = PROG_button then call navigate_to_screen("stat", "PROG")
	If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
	If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
	If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
	If ButtonPressed = SCHL_button then call navigate_to_screen("stat", "SCHL")
	If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
	If ButtonPressed = STIN_button then call navigate_to_screen("stat", "STIN")
	If ButtonPressed = STEC_button then call navigate_to_screen("stat", "STEC")
	If ButtonPressed = STWK_button then call navigate_to_screen("stat", "STWK")
	If ButtonPressed = SHEL_button then call navigate_to_screen("stat", "SHEL")
	If ButtonPressed = SWKR_button then call navigate_to_screen("stat", "SWKR")
	If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
	If ButtonPressed = TYPE_button then call navigate_to_screen("stat", "TYPE")
	If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
END FUNCTION

FUNCTION MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)'Grabbing the footer month/year
	'Does this to check to see if we're on SELF screen
	EMReadScreen SELF_check, 4, 2, 50
	IF SELF_check = "SELF" THEN
		EMReadScreen MAXIS_footer_month, 2, 20, 43
		EMReadScreen MAXIS_footer_year, 2, 20, 46
	ELSE
	Call find_variable("Month: ", MAXIS_footer_month, 2)
		If isnumeric(MAXIS_footer_month) = true then 	'checking to see if a footer month 'number' is present 
			footer_month = MAXIS_footer_month	
			call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
			If isnumeric(MAXIS_footer_year) = true then footer_year = MAXIS_footer_year 'checking to see if a footer year 'number' is present
		Else 'If we don’t have one found, we’re going to assign the current month/year.
			MAXIS_footer_month = DatePart("m", date)   'Datepart delivers the month number to the variable
			If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month   'If it’s a single digit month, add a zero
			MAXIS_footer_year = right(DatePart("yyyy", date), 2)	'We only need the right two characters of the year for MAXIS
		End if
	End if
END FUNCTION

Function memb_navigation_next
  HH_memb_row = HH_memb_row + 1
  EMReadScreen next_HH_memb, 2, HH_memb_row, 3
  If isnumeric(next_HH_memb) = False then
    HH_memb_row = HH_memb_row + 1
  Else
    EMWriteScreen next_HH_memb, 20, 76
    EMWriteScreen "01", 20, 79
  End if
End function

Function memb_navigation_prev
  HH_memb_row = HH_memb_row - 1
  EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
  If isnumeric(prev_HH_memb) = False then
    HH_memb_row = HH_memb_row + 1
  Else
    EMWriteScreen prev_HH_memb, 20, 76
    EMWriteScreen "01", 20, 79
  End if
End function

Function MMIS_RKEY_finder
  'Now we use a Do Loop to get to the start screen for MMIS.
  Do 
    EMSendkey "<PF6>"
    EMWaitReady 0, 0
    EMReadScreen session_start, 18, 1, 7
  Loop until session_start = "SESSION TERMINATED"
  'Now we get back into MMIS. We have to skip past the intro screens.
  EMWriteScreen "mw00", 1, 2
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  'This section may not work for all OSAs, since some only have EK01. This will find EK01 and enter it.
  MMIS_row = 1
  MMIS_col = 1
  EMSearch "EK01", MMIS_row, MMIS_col
  If MMIS_row <> 0 then
    EMWriteScreen "x", MMIS_row, 4
    EMSendKey "<enter>"
    EMWaitReady 0, 0
  End if
  'This section starts from EK01. OSAs may need to skip the previous section.
  EMWriteScreen "x", 10, 3
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

Function navigate_to_MAXIS_screen(function_to_go_to, command_to_go_to)
  EMSendKey "<enter>"
  EMWaitReady 0, 0
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " then
    row = 1
    col = 1
    EMSearch "Function: ", row, col
    If row <> 0 then 
      EMReadScreen MAXIS_function, 4, row, col + 10
      EMReadScreen STAT_note_check, 4, 2, 45
      row = 1
      col = 1
      EMSearch "Case Nbr: ", row, col
      EMReadScreen current_case_number, 8, row, col + 10
      current_case_number = replace(current_case_number, "_", "")
      current_case_number = trim(current_case_number)
    End if
    If current_case_number = case_number and MAXIS_function = ucase(function_to_go_to) and STAT_note_check <> "NOTE" then 
      row = 1
      col = 1
      EMSearch "Command: ", row, col
      EMWriteScreen command_to_go_to, row, col + 9
      EMSendKey "<enter>"
      EMWaitReady 0, 0
    Else
      Do
        EMSendKey "<PF3>"
        EMWaitReady 0, 0
        EMReadScreen SELF_check, 4, 2, 50
      Loop until SELF_check = "SELF"
      EMWriteScreen function_to_go_to, 16, 43
      EMWriteScreen "________", 18, 43
      EMWriteScreen case_number, 18, 43
      EMWriteScreen footer_month, 20, 43
      EMWriteScreen footer_year, 20, 46
      EMWriteScreen command_to_go_to, 21, 70
      EMSendKey "<enter>"
      EMWaitReady 0, 0
      EMReadScreen abended_check, 7, 9, 27
      If abended_check = "abended" then
        EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
	  EMReadScreen ERRR_screen_check, 4, 2, 52
	  If ERRR_screen_check = "ERRR" then 
	    EMSendKey "<enter>"
        EMWaitReady 0, 0
      End if
    End if
  End if
End function

Function navigate_to_PRISM_screen(x) 'x is the name of the screen
  EMWriteScreen x, 21, 18
  EMSendKey "<enter>"
  EMWaitReady 0, 0
End function

function navigation_buttons 'this works by calling the navigation_buttons function when the buttonpressed isn't -1
  If ButtonPressed = ABPS_button then call navigate_to_screen("stat", "ABPS")
  If ButtonPressed = ACCI_button then call navigate_to_screen("stat", "ACCI")
  If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
  If ButtonPressed = ADDR_button then call navigate_to_screen("stat", "ADDR")
  If ButtonPressed = ALTP_button then call navigate_to_screen("stat", "ALTP")
  If ButtonPressed = AREP_button then call navigate_to_screen("stat", "AREP")
  If ButtonPressed = BILS_button then call navigate_to_screen("stat", "BILS")
  If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
  If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
  If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
  If ButtonPressed = COEX_button then call navigate_to_screen("stat", "COEX")
  If ButtonPressed = DCEX_button then call navigate_to_screen("stat", "DCEX")
  If ButtonPressed = DIET_button then call navigate_to_screen("stat", "DIET")
  If ButtonPressed = DISA_button then call navigate_to_screen("stat", "DISA")
  If ButtonPressed = EATS_button then call navigate_to_screen("stat", "EATS")
  If ButtonPressed = ELIG_DWP_button then call navigate_to_screen("elig", "DWP_")
  If ButtonPressed = ELIG_FS_button then call navigate_to_screen("elig", "FS__")
  If ButtonPressed = ELIG_GA_button then call navigate_to_screen("elig", "GA__")
  If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
  If ButtonPressed = ELIG_MFIP_button then call navigate_to_screen("elig", "MFIP")
  If ButtonPressed = ELIG_MSA_button then call navigate_to_screen("elig", "MSA_")
  If ButtonPressed = ELIG_WB_button then call navigate_to_screen("elig", "WB__")
  If ButtonPressed = FACI_button then call navigate_to_screen("stat", "FACI")
  If ButtonPressed = FMED_button then call navigate_to_screen("stat", "FMED")
  If ButtonPressed = HCRE_button then call navigate_to_screen("stat", "HCRE")
  If ButtonPressed = HEST_button then call navigate_to_screen("stat", "HEST")
  If ButtonPressed = IMIG_button then call navigate_to_screen("stat", "IMIG")
  If ButtonPressed = INSA_button then call navigate_to_screen("stat", "INSA")
  If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
  If ButtonPressed = MEDI_button then call navigate_to_screen("stat", "MEDI")
  If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
  If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
  If ButtonPressed = MONT_button then call navigate_to_screen("stat", "MONT")
  If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
  If ButtonPressed = PBEN_button then call navigate_to_screen("stat", "PBEN")
  If ButtonPressed = PDED_button then call navigate_to_screen("stat", "PDED")
  If ButtonPressed = PREG_button then call navigate_to_screen("stat", "PREG")
  If ButtonPressed = PROG_button then call navigate_to_screen("stat", "PROG")
  If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
  If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
  If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
  If ButtonPressed = SCHL_button then call navigate_to_screen("stat", "SCHL")
  If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
  If ButtonPressed = STIN_button then call navigate_to_screen("stat", "STIN")
  If ButtonPressed = STEC_button then call navigate_to_screen("stat", "STEC")
  If ButtonPressed = STWK_button then call navigate_to_screen("stat", "STWK")
  If ButtonPressed = SHEL_button then call navigate_to_screen("stat", "SHEL")
  If ButtonPressed = SWKR_button then call navigate_to_screen("stat", "SWKR")
  If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
  If ButtonPressed = TYPE_button then call navigate_to_screen("stat", "TYPE")
  If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
End function

function new_BS_BSI_heading
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 4 then 
    EMSendKey "--------BURIAL SPACE/ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_row = 5
  end if
End function

function new_CAI_heading
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 4 then 
    EMSendKey "--------CASH ADVANCE ITEMS---------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_row = 5
  end if
End function

function new_page_check
  EMGetCursor MAXIS_row, MAXIS_col
  If MAXIS_row = 17 then
    EMSendKey ">>>>MORE>>>>"
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
    MAXIS_row = 4
  End if
end function

function new_service_heading
  EMGetCursor MAXIS_service_row, MAXIS_service_col
  If MAXIS_service_row = 4 then 
    EMSendKey "--------------SERVICE--------------------AMOUNT----------STATUS--------------" & "<newline>"
    MAXIS_service_row = 5
  end if
End function

Function panel_navigation_next
  EMReadScreen current_panel, 1, 2, 73
  EMReadScreen amount_of_panels, 1, 2, 78
  If current_panel < amount_of_panels then new_panel = current_panel + 1
  If current_panel = amount_of_panels then new_panel = current_panel
  If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
End function

Function panel_navigation_prev
  EMReadScreen current_panel, 1, 2, 73
  EMReadScreen amount_of_panels, 1, 2, 78
  If current_panel = 1 then new_panel = current_panel
  If current_panel > 1 then new_panel = current_panel - 1
  If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
End function

Function PF1
  EMSendKey "<PF1>"
  EMWaitReady 0, 0
End function

Function PF2
  EMSendKey "<PF2>"
  EMWaitReady 0, 0
End function

function PF3
  EMSendKey "<PF3>"
  EMWaitReady 0, 0
end function

Function PF4
  EMSendKey "<PF4>"
  EMWaitReady 0, 0
End function

Function PF5
  EMSendKey "<PF5>"
  EMWaitReady 0, 0
End function

Function PF6
  EMSendKey "<PF6>"
  EMWaitReady 0, 0
End function

Function PF7
  EMSendKey "<PF7>"
  EMWaitReady 0, 0
End function

function PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function

function PF9
  EMSendKey "<PF9>"
  EMWaitReady 0, 0
end function

function PF10
  EMSendKey "<PF10>"
  EMWaitReady 0, 0
end function

Function PF11
  EMSendKey "<PF11>"
  EMWaitReady 0, 0
End function

Function PF12
  EMSendKey "<PF12>"
  EMWaitReady 0, 0
End function

Function PF13
  EMSendKey "<PF13>"
  EMWaitReady 0, 0
End function

Function PF14
  EMSendKey "<PF14>"
  EMWaitReady 0, 0
End function

Function PF15
  EMSendKey "<PF15>"
  EMWaitReady 0, 0
End function

Function PF16
  EMSendKey "<PF16>"
  EMWaitReady 0, 0
End function

Function PF17
  EMSendKey "<PF17>"
  EMWaitReady 0, 0
End function

Function PF18
  EMSendKey "<PF18>"
  EMWaitReady 0, 0
End function

function PF19
  EMSendKey "<PF19>"
  EMWaitReady 0, 0
end function

function PF20
  EMSendKey "<PF20>"
  EMWaitReady 0, 0
end function

function PF21
  EMSendKey "<PF21>"
  EMWaitReady 0, 0
end function

function PF22
  EMSendKey "<PF22>"
  EMWaitReady 0, 0
end function

function PF23
  EMSendKey "<PF23>"
  EMWaitReady 0, 0
end function

function PF24
  EMSendKey "<PF24>"
  EMWaitReady 0, 0
end function

'Asks the user if they want to proceed. Result_of_msgbox parameter returns TRUE if Yes is pressed, and FALSE if No is pressed.
FUNCTION proceed_confirmation(result_of_msgbox)
	If ButtonPressed = -1 then 
		proceed_confirm = MsgBox("Are you sure you want to proceed? Press Yes to continue, No to return to the previous screen, and Cancel to end the script.", vbYesNoCancel)
		If proceed_confirm = vbCancel then stopscript
		If proceed_confirm = vbYes then result_of_msgbox = TRUE
		If proceed_confirm = vbNo then result_of_msgbox = FALSE
	End if
END FUNCTION

function run_another_script(script_path)
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile(script_path)
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  Execute text_from_the_other_script
end function

FUNCTION run_from_GitHub(url)
	'Creates a list of items to remove from anything run from GitHub. This will allow for counties to use Option Explicit handling without fear.
	list_of_things_to_remove = array("OPTION EXPLICIT", _
									"option explicit", _
									"Option Explicit", _
									"dim case_number", _
									"DIM case_number", _
									"Dim case_number")
	If run_locally = "" or run_locally = False then					'Runs the script from GitHub if we're not set up to run locally.
		Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
		req.open "GET", url, False									'Attempts to open the URL
		req.send													'Sends request
		If req.Status = 200 Then									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			script_contents = req.responseText						'Empties the response into a variable called script_contents
			'Uses a for/next to remove the list_of_things_to_remove
			FOR EACH phrase IN list_of_things_to_remove		
				script_contents = replace(script_contents, phrase, "")
			NEXT
			Execute script_contents									'Executes the remaining script code
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
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		call run_another_script(url)
	END IF
END FUNCTION

function script_end_procedure(closing_message)
	If closing_message <> "" then MsgBox closing_message
	stop_time = timer
	script_run_time = stop_time - start_time
	If is_county_collecting_stats  = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork") 
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Fixing a bug when the script_end_procedure has an apostrophe (this interferes with Access)
		closing_message = replace(closing_message, "'", "")

		'Opening DB
		objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""

		'Opening usage_log and adding a record
		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic	
	End if
	stopscript
end function

function script_end_procedure_wsh(closing_message) 'For use when running a script outside of the BlueZone Script Host
	If closing_message <> "" then MsgBox closing_message
	stop_time = timer
	script_run_time = stop_time - start_time
	If is_county_collecting_stats = True then
		'Getting user name
		Set objNet = CreateObject("WScript.NetWork") 
		user_ID = objNet.UserName

		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'Opening DB
		objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "" & stats_database_path & ""

		'Opening usage_log and adding a record
		objRecordSet.Open "INSERT INTO usage_log (USERNAME, SDATE, STIME, SCRIPT_NAME, SRUNTIME, CLOSING_MSGBOX)" &  _
		"VALUES ('" & user_ID & "', '" & date & "', '" & time & "', '" & name_of_script & "', " & script_run_time & ", '" & closing_message & "')", objConnection, adOpenStatic, adLockOptimistic
	End if
	Wscript.Quit
end function

'Navigates you to a blank case note, presses PF9, and checks to make sure you're in edit mode (keeping you from writing all of the case note on an inquiry screen).
FUNCTION start_a_blank_CASE_NOTE
	call navigate_to_screen("case", "note")
	DO
		PF9
		EMReadScreen case_note_check, 17, 2, 33
		EMReadScreen mode_check, 1, 20, 09
		If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then msgbox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."
	Loop until (mode_check = "A" or mode_check = "E")
END FUNCTION

function stat_navigation
  EMReadScreen STAT_check, 4, 20, 21
  If STAT_check = "STAT" then
    If ButtonPressed = prev_panel_button then 
      EMReadScreen current_panel, 1, 2, 73
      EMReadScreen amount_of_panels, 1, 2, 78
      If current_panel = 1 then new_panel = current_panel
      If current_panel > 1 then new_panel = current_panel - 1
      If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
    End if
    If ButtonPressed = next_panel_button then 
      EMReadScreen current_panel, 1, 2, 73
      EMReadScreen amount_of_panels, 1, 2, 78
      If current_panel < amount_of_panels then new_panel = current_panel + 1
      If current_panel = amount_of_panels then new_panel = current_panel
      If amount_of_panels > 1 then EMWriteScreen "0" & new_panel, 20, 79
    End if
    If ButtonPressed = prev_memb_button then 
      HH_memb_row = HH_memb_row - 1
      EMReadScreen prev_HH_memb, 2, HH_memb_row, 3
      If isnumeric(prev_HH_memb) = False then
        HH_memb_row = HH_memb_row + 1
      Else
        EMWriteScreen prev_HH_memb, 20, 76
        EMWriteScreen "01", 20, 79
      End if
    End if
    If ButtonPressed = next_memb_button then 
      HH_memb_row = HH_memb_row + 1
      EMReadScreen next_HH_memb, 2, HH_memb_row, 3
      If isnumeric(next_HH_memb) = False then
        HH_memb_row = HH_memb_row + 1
      Else
        EMWriteScreen next_HH_memb, 20, 76
        EMWriteScreen "01", 20, 79
      End if
    End if
  End if
End function

Function step_through_handling 'This function will introduce "warning screens" before each transmit, which is very helpful for testing new scripts
	'To use this function, simply replace the "Execute text_from_the_other_script" line with:
	'Execute replace(text_from_the_other_script, "EMWaitReady 0, 0", "step_through_handling")
	step_through = MsgBox("Step " & step_number & chr(13) & chr(13) & "If you see something weird on your screen (like a MAXIS or PRISM error), PRESS CANCEL then email your script administrator about it. Make sure you include the step you're on.", vbOKCancel)
	If step_number = "" then step_number = 1	'Declaring the variable
	If step_through = vbCancel then
		stopscript
	Else
		EMWaitReady 0, 0
		step_number = step_number + 1
	End if
End Function

function transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

Function worker_county_code_determination(worker_county_code_variable, two_digit_county_code_variable)		'Determines worker_county_code and two_digit_county_code for multi-county agencies and DHS staff
	If left(code_from_installer, 2) = "PT" then 'special handling for Pine Tech
		worker_county_code_variable = "PWVTS"
		county_name = "Pine Tech"
	Else
		If worker_county_code_variable = "MULTICOUNTY" or worker_county_code_variable = "" then 		'If the user works for many counties (i.e. SWHHS) or isn't assigned (i.e. a scriptwriter) it asks.
			Do
				two_digit_county_code_variable = inputbox("Select the county to proxy as. Ex: ''01''")
				If two_digit_county_code_variable = "" then stopscript
				If len(two_digit_county_code_variable) <> 2 or isnumeric(two_digit_county_code_variable) = False then MsgBox "Your county proxy code should be two digits and numeric."
			Loop until len(two_digit_county_code_variable) = 2 and isnumeric(two_digit_county_code_variable) = True 
			worker_county_code_variable = "x1" & two_digit_county_code_variable
			If two_digit_county_code_variable = "91" then worker_county_code_variable = "PW"	'For DHS folks without proxy
			
			'Determining county name
			if worker_county_code_variable = "x101" then 
				county_name = "Aitkin County"
			elseif worker_county_code_variable = "x102" then 
				county_name = "Anoka County"
			elseif worker_county_code_variable = "x103" then 
				county_name = "Becker County"
			elseif worker_county_code_variable = "x104" then 
				county_name = "Beltrami County"
			elseif worker_county_code_variable = "x105" then 
				county_name = "Benton County"
			elseif worker_county_code_variable = "x106" then 
				county_name = "Big Stone County"
			elseif worker_county_code_variable = "x107" then 
				county_name = "Blue Earth County"
			elseif worker_county_code_variable = "x108" then 
				county_name = "Brown County"
			elseif worker_county_code_variable = "x109" then 
				county_name = "Carlton County"
			elseif worker_county_code_variable = "x110" then 
				county_name = "Carver County"
			elseif worker_county_code_variable = "x111" then 
				county_name = "Cass County"
			elseif worker_county_code_variable = "x112" then 
				county_name = "Chippewa County"
			elseif worker_county_code_variable = "x113" then 
				county_name = "Chisago County"
			elseif worker_county_code_variable = "x114" then 
				county_name = "Clay County"
			elseif worker_county_code_variable = "x115" then 
				county_name = "Clearwater County"
			elseif worker_county_code_variable = "x116" then 
				county_name = "Cook County"
			elseif worker_county_code_variable = "x117" then 
				county_name = "Cottonwood County"
			elseif worker_county_code_variable = "x118" then 
				county_name = "Crow Wing County"
			elseif worker_county_code_variable = "x119" then 
				county_name = "Dakota County"
			elseif worker_county_code_variable = "x120" then 
				county_name = "Dodge County"
			elseif worker_county_code_variable = "x121" then 
				county_name = "Douglas County"
			elseif worker_county_code_variable = "x122" then 
				county_name = "Faribault County"
			elseif worker_county_code_variable = "x123" then 
				county_name = "Fillmore County"
			elseif worker_county_code_variable = "x124" then 
				county_name = "Freeborn County"
			elseif worker_county_code_variable = "x125" then 
				county_name = "Goodhue County"
			elseif worker_county_code_variable = "x126" then 
				county_name = "Grant County"
			elseif worker_county_code_variable = "x127" then 
				county_name = "Hennepin County"
			elseif worker_county_code_variable = "x128" then 
				county_name = "Houston County"
			elseif worker_county_code_variable = "x129" then 
				county_name = "Hubbard County"
			elseif worker_county_code_variable = "x130" then 
				county_name = "Isanti County"
			elseif worker_county_code_variable = "x131" then 
				county_name = "Itasca County"
			elseif worker_county_code_variable = "x132" then 
				county_name = "Jackson County"
			elseif worker_county_code_variable = "x133" then 
				county_name = "Kanabec County"
			elseif worker_county_code_variable = "x134" then
				county_name = "Kandiyohi County"
			elseif worker_county_code_variable = "x135" then 	
				county_name = "Kittson County"
			elseif worker_county_code_variable = "x136" then 	
				county_name = "Koochiching County"
			elseif worker_county_code_variable = "x137" then 	
				county_name = "Lac Qui Parle County"
			elseif worker_county_code_variable = "x138" then 	
				county_name = "Lake County"
			elseif worker_county_code_variable = "x139" then 	
				county_name = "Lake of the Woods County"
			elseif worker_county_code_variable = "x140" then 	
				county_name = "LeSueur County"
			elseif worker_county_code_variable = "x141" then 	
				county_name = "Lincoln County"
			elseif worker_county_code_variable = "x142" then 	
				county_name = "Lyon County"
			elseif worker_county_code_variable = "x143" then 	
				county_name = "Mcleod County"
			elseif worker_county_code_variable = "x144" then 	
				county_name = "Mahnomen County"
			elseif worker_county_code_variable = "x145" then 	
				county_name = "Marshall County"
			elseif worker_county_code_variable = "x146" then 	
				county_name = "Martin County"
			elseif worker_county_code_variable = "x147" then 	
				county_name = "Meeker County"
			elseif worker_county_code_variable = "x148" then 	
				county_name = "Mille Lacs County"
			elseif worker_county_code_variable = "x149" then 	
				county_name = "Morrison County"
			elseif worker_county_code_variable = "x150" then 	
				county_name = "Mower County"
			elseif worker_county_code_variable = "x151" then 	
				county_name = "Murray County"
			elseif worker_county_code_variable = "x152" then 	
				county_name = "Nicollet County"
			elseif worker_county_code_variable = "x153" then 	
				county_name = "Nobles County"
			elseif worker_county_code_variable = "x154" then 	
				county_name = "Norman County"
			elseif worker_county_code_variable = "x155" then 	
				county_name = "Olmsted County"
			elseif worker_county_code_variable = "x156" then 	
				county_name = "Otter Tail County"
			elseif worker_county_code_variable = "x157" then 	
				county_name = "Pennington County"
			elseif worker_county_code_variable = "x158" then 	
				county_name = "Pine County"
			elseif worker_county_code_variable = "x159" then 	
				county_name = "Pipestone County"
			elseif worker_county_code_variable = "x160" then 	
				county_name = "Polk County"
			elseif worker_county_code_variable = "x161" then 	
				county_name = "Pope County"
			elseif worker_county_code_variable = "x162" then 	
				county_name = "Ramsey County"
			elseif worker_county_code_variable = "x163" then 	
				county_name = "Red Lake County"
			elseif worker_county_code_variable = "x164" then 	
				county_name = "Redwood County"
			elseif worker_county_code_variable = "x165" then 	
				county_name = "Renville County"
			elseif worker_county_code_variable = "x166" then 	
				county_name = "Rice County"
			elseif worker_county_code_variable = "x167" then 	
				county_name = "Rock County"
			elseif worker_county_code_variable = "x168" then 	
				county_name = "Roseau County"
			elseif worker_county_code_variable = "x169" then 	
				county_name = "St. Louis County"
			elseif worker_county_code_variable = "x170" then 	
				county_name = "Scott County"
			elseif worker_county_code_variable = "x171" then 	
				county_name = "Sherburne County"
			elseif worker_county_code_variable = "x172" then 	
				county_name = "Sibley County"
			elseif worker_county_code_variable = "x173" then 	
				county_name = "Stearns County"
			elseif worker_county_code_variable = "x174" then 	
				county_name = "Steele County"
			elseif worker_county_code_variable = "x175" then 	
				county_name = "Stevens County"
			elseif worker_county_code_variable = "x176" then 	
				county_name = "Swift County"
			elseif worker_county_code_variable = "x177" then 	
				county_name = "Todd County"
			elseif worker_county_code_variable = "x178" then 	
				county_name = "Traverse County"
			elseif worker_county_code_variable = "x179" then 	
				county_name = "Wabasha County"
			elseif worker_county_code_variable = "x180" then 	
				county_name = "Wadena County"
			elseif worker_county_code_variable = "x181" then 	
				county_name = "Waseca County"
			elseif worker_county_code_variable = "x182" then 	
				county_name = "Washington County"
			elseif worker_county_code_variable = "x183" then 	
				county_name = "Watonwan County"
			elseif worker_county_code_variable = "x184" then 	
				county_name = "Wilkin County"
			elseif worker_county_code_variable = "x185" then 	
				county_name = "Winona County"
			elseif worker_county_code_variable = "x186" then 	
				county_name = "Wright County"
			elseif worker_county_code_variable = "x187" then 	
				county_name = "Yellow Medicine County"
			elseif worker_county_code_variable = "x188" then 
				county_name = "Mille Lacs Band"
			elseif worker_county_code_variable = "x192" then 
				county_name = "White Earth Nation"
			end if
		End If
	End if
End function

Function write_bullet_and_variable_in_CASE_NOTE(bullet, variable)
	If trim(variable) <> "" then
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
			If character_test <> " " or noting_row >= 18 then 
				noting_row = noting_row + 1
				
				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then 
					EMSendKey "<PF8>"
					EMWaitReady 0, 0
					
					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
					Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
					End if
				End if
			End if
		Loop until character_test = " "
	
		'Looks at the length of the bullet. This determines the indent for the rest of the info. Going with a maximum indent of 18.
		If len(bullet) >= 14 then
			indent_length = 18	'It's four more than the bullet text to account for the asterisk, the colon, and the spaces.
		Else
			indent_length = len(bullet) + 4 'It's four more for the reason explained above.
		End if
	
		'Writes the bullet
		EMWriteScreen "* " & bullet & ": ", noting_row, noting_col
	
		'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
		noting_col = noting_col + (len(bullet) + 4)
	
		'Splits the contents of the variable into an array of words
		variable_array = split(variable, " ")
	
		For each word in variable_array
			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then 
				noting_row = noting_row + 1
				noting_col = 3
			End if
			
			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0
				
				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
				EMReadScreen end_of_case_note_check, 1, 24, 2
				If end_of_case_note_check = "A" then
					EMSendKey "<PF3>"												'PF3s
					EMWaitReady 0, 0
					EMSendKey "<PF9>"												'PF9s (opens new note)
					EMWaitReady 0, 0
					EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
					EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
					noting_row = 5													'Resets this variable to work in the new locale
				Else
					noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
				End if
			End if
			
			'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
			If noting_col = 3 then 
				EMWriteScreen space(indent_length), noting_row, noting_col	
				noting_col = noting_col + indent_length
			End if
	
			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col
			
			'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
			If right(word, 1) = ";" then
				noting_row = noting_row + 1
				noting_col = 3
				EMWriteScreen space(indent_length), noting_row, noting_col	
				noting_col = noting_col + indent_length
			End if
			
			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next 
	
		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
End function

Function write_bullet_and_variable_in_CCOL_NOTE(bullet, variable)

	EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
	noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
	'The following figures out if we need a new page, or if we need a new case note entirely as well.
	Do
		EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
		If character_test <> " " or noting_row >= 19 then 
			noting_row = noting_row + 1
			
			'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
			If noting_row >= 19 then 
				EMSendKey "<PF8>"
				EMWaitReady 0, 0
				
				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
				EMReadScreen end_of_case_note_check, 1, 24, 2
				If end_of_case_note_check = "A" then
					EMSendKey "<PF3>"												'PF3s
					EMWaitReady 0, 0
					EMSendKey "<PF9>"												'PF9s (opens new note)
					EMWaitReady 0, 0
					EMWriteScreen "~~~continued from previous note~~~", 5, 	3		'enters a header
					EMSetCursor 6, 3												'Sets cursor in a good place to start noting.
					noting_row = 6													'Resets this variable to work in the new locale
				Else
					noting_row = 5													'Resets this variable to 5 if we did not need a brand new note.
				End if
			End if
		End if
	Loop until character_test = " "

	'Looks at the length of the bullet. This determines the indent for the rest of the info. Going with a maximum indent of 18.
	If len(bullet) >= 14 then
		indent_length = 18	'It's four more than the bullet text to account for the asterisk, the colon, and the spaces.
	Else
		indent_length = len(bullet) + 4 'It's four more for the reason explained above.
	End if

	'Writes the bullet
	EMWriteScreen "* " & bullet & ": ", noting_row, noting_col

	'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
	noting_col = noting_col + (len(bullet) + 4)

	'Splits the contents of the variable into an array of words
	variable_array = split(variable, " ")

	For each word in variable_array
		'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
		If len(word) + noting_col > 80 then 
			noting_row = noting_row + 1
			noting_col = 3
		End if
		
		'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
		If noting_row >= 18 then
			EMSendKey "<PF8>"
			EMWaitReady 0, 0
			
			'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
			EMReadScreen end_of_case_note_check, 1, 24, 2
			If end_of_case_note_check = "A" then
				EMSendKey "<PF3>"												'PF3s
				EMWaitReady 0, 0
				EMSendKey "<PF9>"												'PF9s (opens new note)
				EMWaitReady 0, 0
				EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
				EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
				noting_row = 6													'Resets this variable to work in the new locale
			Else
				noting_row = 5													'Resets this variable to 4 if we did not need a brand new note.
			End if
		End if
		
		'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
		If noting_col = 3 then 
			EMWriteScreen space(indent_length), noting_row, noting_col	
			noting_col = noting_col + indent_length
		End if

		'Writes the word and a space using EMWriteScreen
		EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col
		
		'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
		If right(word, 1) = ";" then
			noting_row = noting_row + 1
			noting_col = 3
			EMWriteScreen space(indent_length), noting_row, noting_col	
			noting_col = noting_col + indent_length
		End if
		
		'Increases noting_col the length of the word + 1 (for the space)
		noting_col = noting_col + (len(word) + 1)
	Next 

	'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
	EMSetCursor noting_row + 1, 3

End function

Function write_three_columns_in_CASE_NOTE(col_01_start_point, col_01_variable, col_02_start_point, col_02_variable, col_03_start_point, col_03_variable)
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80 + 1 ) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMGetCursor row, col
  EMWriteScreen "                                                                              ", row, 3
  EMSetCursor row, col_01_start_point
  EMSendKey col_01_variable
  EMSetCursor row, col_02_start_point
  EMSendKey col_02_variable
  EMSetCursor row, col_03_start_point
  EMSendKey col_03_variable
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 17 and col + (len(x)) >= 80) or (row = 4 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

FUNCTION write_value_and_transmit(input_value, MAXIS_row, MAXIS_col)
	EMWriteScreen input_value, MAXIS_row, MAXIS_col
	transmit
END FUNCTION

Function write_variable_in_CASE_NOTE(variable)
	If trim(variable) <> "" THEN
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
			If character_test <> " " or noting_row >= 18 then 
				noting_row = noting_row + 1
				
				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then 
					EMSendKey "<PF8>"
					EMWaitReady 0, 0
					
					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
					Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
					End if
				End if
			End if
		Loop until character_test = " "
	
		'Splits the contents of the variable into an array of words
		variable_array = split(variable, " ")
	
		For each word in variable_array
	
			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 80 then 
				noting_row = noting_row + 1
				noting_col = 3
			End if
			
			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0
				
				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
				EMReadScreen end_of_case_note_check, 1, 24, 2
				If end_of_case_note_check = "A" then
					EMSendKey "<PF3>"												'PF3s
					EMWaitReady 0, 0
					EMSendKey "<PF9>"												'PF9s (opens new note)
					EMWaitReady 0, 0
					EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
					EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
					noting_row = 5													'Resets this variable to work in the new locale
				Else
					noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
				End if
			End if
	
			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col
			
			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next 
	
		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
End function

Function write_variable_in_CCOL_NOTE(variable)

	EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
	noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
	'The following figures out if we need a new page, or if we need a new case note entirely as well.
	Do
		EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
		If character_test <> " " or noting_row >= 19 then 
			noting_row = noting_row + 1
			
			'If we get to row 19 (which can't be read here), it will go to the next panel (PF8).
			If noting_row >= 19 then 
				EMSendKey "<PF8>"
				EMWaitReady 0, 0
				
				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
				EMReadScreen end_of_case_note_check, 1, 24, 2
				If end_of_case_note_check = "A" then
					EMSendKey "<PF3>"												'PF3s
					EMWaitReady 0, 0
					EMSendKey "<PF9>"												'PF9s (opens new note)
					EMWaitReady 0, 0
					EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
					EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
					noting_row = 6													'Resets this variable to work in the new locale
				Else
					noting_row = 5													'Resets this variable to 5 if we did not need a brand new note.
				End if
			End if
		End if
	Loop until character_test = " "

	'Splits the contents of the variable into an array of words
	variable_array = split(variable, " ")

	For each word in variable_array

		'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
		If len(word) + noting_col > 80 then 
			noting_row = noting_row + 1
			noting_col = 3
		End if
		
		'If the next line is row 19 (you can't write to row 19), it will PF8 to get to the next page
		If noting_row >= 19 then
			EMSendKey "<PF8>"
			EMWaitReady 0, 0
			
			'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
			EMReadScreen end_of_case_note_check, 1, 24, 2
			If end_of_case_note_check = "A" then
				EMSendKey "<PF3>"												'PF3s
				EMWaitReady 0, 0
				EMSendKey "<PF9>"												'PF9s (opens new note)
				EMWaitReady 0, 0
				EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
				EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
				noting_row = 6													'Resets this variable to work in the new locale
			Else
				noting_row = 5													'Resets this variable to 5 if we did not need a brand new note.
			End if
		End if

		'Writes the word and a space using EMWriteScreen
		EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col
		
		'Increases noting_col the length of the word + 1 (for the space)
		noting_col = noting_col + (len(word) + 1)
	Next 

	'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
	EMSetCursor noting_row + 1, 3

End function

Function write_variable_in_SPEC_MEMO(variable)
	EMGetCursor memo_row, memo_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
	memo_col = 15										'The memo col should always be 15 at this point, because it's the beginning. But, this will be dynamically recreated each time.
	'The following figures out if we need a new page
	Do
		EMReadScreen character_test, 1, memo_row, memo_col 	'Reads a single character at the memo row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond memo range).
		If character_test <> " " or memo_row >= 18 then 
			memo_row = memo_row + 1
			
			'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
			If memo_row >= 18 then
				PF8
				memo_row = 3					'Resets this variable to 3
			End if
		End if
	Loop until character_test = " "
	
	'Each word becomes its own member of the array called variable_array.
	variable_array = split(variable, " ")					
  
	For each word in variable_array 
		'If the length of the word would go past col 74 (you can't write to col 74), it will kick it to the next line
		If len(word) + memo_col > 74 then 
			memo_row = memo_row + 1
			memo_col = 15
		End if
		
		'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
		If memo_row >= 18 then
			PF8
			memo_row = 3					'Resets this variable to 3
		End if
	
		'Writes the word and a space using EMWriteScreen
		EMWriteScreen word & " ", memo_row, memo_col
			
		'Increases memo_col the length of the word + 1 (for the space)
		memo_col = memo_col + (len(word) + 1)
	Next 
	
	'After the array is processed, set the cursor on the following row, in col 15, so that the user can enter in information here (just like writing by hand). 
	EMSetCursor memo_row + 1, 15
End function

Function write_variable_in_TIKL(variable)
	IF len(variable) <= 60 THEN
		tikl_line_one = variable
	ELSE
		tikl_line_one_len = 61
		tikl_line_one = left(variable, tikl_line_one_len)
		IF right(tikl_line_one, 1) = " " THEN
			whats_left_after_one = right(variable, (len(variable) - tikl_line_one_len))
		ELSE
			DO
				tikl_line_one = left(variable, (tikl_line_one_len - 1))
				IF right(tikl_line_one, 1) <> " " THEN tikl_line_one_len = tikl_line_one_len - 1
			LOOP UNTIL right(tikl_line_one, 1) = " "
			whats_left_after_one = right(variable, (len(variable) - (tikl_line_one_len - 1)))
		END IF
	END IF

	IF (whats_left_after_one <> "" AND len(whats_left_after_one) <= 60) THEN
		tikl_line_two = whats_left_after_one
	ELSEIF (whats_left_after_one <> "" AND len(whats_left_after_one) > 60) THEN
		tikl_line_two_len = 61
		tikl_line_two = left(whats_left_after_one, tikl_line_two_len)
		IF right(tikl_line_two, 1) = " " THEN
			whats_left_after_two = right(whats_left_after_one, (len(whats_left_after_one) - tikl_line_two_len))
		ELSE
			DO
				tikl_line_two = left(whats_left_after_one, (tikl_line_two_len - 1))
				IF right(tikl_line_two, 1) <> " " THEN tikl_line_two_len = tikl_line_two_len - 1
			LOOP UNTIL right(tikl_line_two, 1) = " "
			whats_left_after_two = right(whats_left_after_one, (len(whats_left_after_one) - (tikl_line_two_len - 1)))
		END IF
	END IF

	IF (whats_left_after_two <> "" AND len(whats_left_after_two) <= 60) THEN
		tikl_line_three = whats_left_after_two
	ELSEIF (whats_left_after_two <> "" AND len(whats_left_after_two) > 60) THEN
		tikl_line_three_len = 61
		tikl_line_three = right(whats_left_after_two, tikl_line_three_len)
		IF right(tikl_line_three, 1) = " " THEN
			whats_left_after_three = right(whats_left_after_two, (len(whats_left_after_two) - tikl_line_three_len))
		ELSE
			DO
				tikl_line_three = left(whats_left_after_two, (tikl_line_three_len - 1))
				IF right(tikl_line_three, 1) <> " " THEN tikl_line_three_len = tikl_line_three_len - 1
			LOOP UNTIL right(tikl_line_three, 1) = " "
			whats_left_after_three = right(whats_left_after_two, (len(whats_left_after_two) - (tikl_line_three_len - 1)))
		END IF
	END IF

	IF (whats_left_after_three <> "" AND len(whats_left_after_three) <= 60) THEN
		tikl_line_four = whats_left_after_three
	ELSEIF (whats_left_after_three <> "" AND len(whats_left_after_three) > 60) THEN
		tikl_line_four_len = 61
		tikl_line_four = left(whats_left_after_three, tikl_line_four_len)
		IF right(tikl_line_four, 1) = " " THEN
			tikl_line_five = right(whats_left_after_three, (len(whats_left_after_three) - tikl_line_four_len))
		ELSE
			DO
				tikl_line_four = left(whats_left_after_three, (tikl_line_four_len - 1))
				IF right(tikl_line_four, 1) <> " " THEN tikl_line_four_len = tikl_line_four_len - 1
			LOOP UNTIL right(tikl_line_four, 1) = " "
			tikl_line_five = right(whats_left_after_three, (tikl_line_four_len - 1))
		END IF
	END IF

	EMWriteScreen tikl_line_one, 9, 3
	IF tikl_line_two <> "" THEN EMWriteScreen tikl_line_two, 10, 3
	IF tikl_line_three <> "" THEN EMWriteScreen tikl_line_three, 11, 3
	IF tikl_line_four <> "" THEN EMWriteScreen tikl_line_four, 12, 3
	IF tikl_line_five <> "" THEN EMWriteScreen tikl_line_five, 13, 3
	transmit
End function

'--------DEPRECIATED FUNCTIONS KEPT FOR COMPATIBILITY PURPOSES, THE NEW FUNCTIONS ARE INDICATED WITHIN THE OLD FUNCTIONS
Function ERRR_screen_check 'Checks for error prone cases				'DEPRECIATED AS OF 01/20/2015.
	EMReadScreen ERRR_check, 4, 2, 52	'Now included in NAVIGATE_TO_MAXIS_SCREEN
	If ERRR_check = "ERRR" then transmit
End Function

Function maxis_check_function											'DEPRECIATED AS OF 01/20/2015.
	call check_for_MAXIS(True)	'Always true, because the original function always exited, and this needs to match the original function for reverse compatibility reasons.
End function

Function navigate_to_screen(x, y)										'DEPRECIATED AS OF 03/09/2015.
	call navigate_to_MAXIS_screen(x, y)
End function

Function write_editbox_in_case_note(bullet, variable, length_of_indent) 'DEPRECIATED AS OF 01/20/2015. 
	call write_bullet_and_variable_in_case_note(bullet, variable)
End function

Function write_new_line_in_case_note(variable)							'DEPRECIATED AS OF 01/20/2015. 
	call write_variable_in_CASE_NOTE(variable)
End function

Function write_new_line_in_SPEC_MEMO(variable_to_enter)					'DEPRECIATED AS OF 01/20/2015. 
	call write_variable_in_SPEC_MEMO(variable_to_enter)
End function

FUNCTION write_TIKL_function(variable)									'DEPRECIATED AS OF 01/20/2015.
	call write_variable_in_TIKL(variable)
END FUNCTION



'<<<<<<<<<<<<THESE VARIABLES ARE TEMPORARY, DESIGNED TO KEEP CERTAIN COUNTIES FROM ACCIDENTALLY JOINING THE BETA, DUE TO A GLITCH IN THE INSTALLER WHICH WAS CORRECTED IN VERSION 1.3.1
If beta_agency = True then 
	'These counties are NOT part of the beta. Because of that, if they are showing up as part of the beta, it will manually remove them from the beta branch and set them back to release.
	'	As of 06/03/2015, it will also deliver a MsgBox telling them they need to update. These counties should have fixed this back in January when this was first posted on SIR.
	If worker_county_code = "x101" or _
	   worker_county_code = "x103" or _
	   worker_county_code = "x106" or _
	   worker_county_code = "x107" or _
	   worker_county_code = "x108" or _
	   worker_county_code = "x109" or _
	   worker_county_code = "x110" or _
	   worker_county_code = "x111" or _
	   worker_county_code = "x112" or _
	   worker_county_code = "x113" or _
	   worker_county_code = "x114" or _
	   worker_county_code = "x115" or _
	   worker_county_code = "x116" or _
	   worker_county_code = "x117" or _
	   worker_county_code = "x121" or _
	   worker_county_code = "x122" or _
	   worker_county_code = "x124" or _
	   worker_county_code = "x126" or _
	   worker_county_code = "x128" or _
	   worker_county_code = "x129" or _
	   worker_county_code = "x130" or _
	   worker_county_code = "x131" or _
	   worker_county_code = "x132" or _
	   worker_county_code = "x134" or _
	   worker_county_code = "x135" or _
	   worker_county_code = "x136" or _
	   worker_county_code = "x137" or _
	   worker_county_code = "x138" or _
	   worker_county_code = "x139" or _
	   worker_county_code = "x140" or _
	   worker_county_code = "x143" or _
	   worker_county_code = "x144" or _
	   worker_county_code = "x145" or _
	   worker_county_code = "x146" or _
	   worker_county_code = "x148" or _
	   worker_county_code = "x149" or _
	   worker_county_code = "x152" or _
	   worker_county_code = "x153" or _
	   worker_county_code = "x154" or _
	   worker_county_code = "x156" or _
	   worker_county_code = "x158" or _
	   worker_county_code = "x161" or _
	   worker_county_code = "x163" or _
	   worker_county_code = "x165" or _
	   worker_county_code = "x166" or _
	   worker_county_code = "x168" or _
	   worker_county_code = "x170" or _
	   worker_county_code = "x171" or _
	   worker_county_code = "x172" or _
	   worker_county_code = "x175" or _
	   worker_county_code = "x176" or _
	   worker_county_code = "x177" or _
	   worker_county_code = "x178" or _
	   worker_county_code = "x180" or _
	   worker_county_code = "x182" or _
	   worker_county_code = "x183" or _
	   worker_county_code = "x184" or _
	   worker_county_code = "x185" or _
	   worker_county_code = "x187" then 
		MsgBox "If you are seeing this message, it's because a script glitch has been detected, which requires an alpha user to reinstall the scripts for your county." & vbNewLine & vbNewLine & _
		  "Instructions for updating your scripts can be found on SIR, in a document titled ""Beta agency bug fix 01.27.2015"". Please ask an alpha user to follow these instructions to correct this issue." & vbNewLine & vbNewLine & _
		  "This script will now stop."
		stopscript
	End if
End if
