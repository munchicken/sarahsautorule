# Sarah's AutoRule

#### Description:

    Outlook VBA Macro to automatically create a "Contact Group" rule from the selected email
    
    Based on the MS product team article ["Best practices for Outlook 2010"](https://support.office.com/en-us/article/Best-practices-for-Outlook-2010-f90e5f69-8832-4d89-95b3-bfdf76c82ef8)
    
#### Instructions:

    Run the macro on an email that is selected in the Outlook Explorer that you want automatically moved to a "Contact Group" folder
    
#### Actions:

    It checks to see if there is an existing rule for this sender
    
    If there is not an exising rule, then it creates one with the following settings:
    
      Move messages from "Sender" to "Folder"
      
      It checks for a "Contact Groups" folder, and creates one if necessary
      
      It then checks for a folder in Contact Groups named "Sender", and creates one if necessary
      
      Except if users name is in the To or Cc box
      
      Except if "specific words" are in the subject or body (see the array marked with "+++++" if you would like to change these)
      
      Stop processing more rules
      
      It moves the new rule to the bottom of the rule list
      
      It then runs the new rule
      
    If there is an existing rule, it checks to see if this is a new email address, if so it adds it to the existing rule and re-runs the rule
    
    If the rule exists and has the correct email addresses, then this email is in your Inbox due to one of the exceptions
    
      If you choose not to delete the email, but rather run AutoRule on it, then it assumes you just want to move it to the proper folder and does so
      
#### Notes:

    The notification box indicates all actions taken
    
    You can check & modify any created rules in the Outlook Rules & Alerts Manager
    
#### Installation:

    Download the module (SarahsAutoRule.bas) & forms (SarahsAutoRuleUserForm .frm & .frx)
    
    Enable the "Developer" tab on the Outlook ribbon
    
    From the Developer tab, click Visual Basic to open the editor
    
    Under "Froms" in the Project Explorer (left), import the two forms
    
    Under "Modules", import the module
    
    The macro can be run from from the Developer tab or can be placed in a menu like the ribbon, etc.
    
    You may need to adjust your Macro security settings (macro can also be self-signed with SelfCert, if desired)
