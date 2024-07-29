# MonacoScripting-batchexport
Automate DICOM export in Elekta Monaco for a list of patients.

This script was developed to bulk export patient DICOM data (CT, Structures, DOSE) to Elekta's application 'ProKnow', 
to be used as a cloud-based backup solution. However, any export location can be used.
The script only opens and exports approved plans.

Steps:
1. First, create a csv file in Excel with patient IDs in the first column.
2. Edit the App.config file with your Installation, Clinic, csv file path, and export location name.
3. Open the C# program file in Visual Studio,
4. Edit the export location fields to match your Monaco instance.
5. Compile the program pressing 'Start'.
6. Upload the contents of the 'bin/Release' folder to Monaco Scripting.
7. Approve and run the Script.


Notes:
Some of the Monaco API methods have built-in exeptions which will stop the script if encountered. This makes custom error handling difficult. 
Some basic error handling is included, such as skipping if the patient is open by another user, can't find CSV etc.

**ProKnow Users**
Due to the limited storage of ProKnow, a Python script is included which utilises the ProKnow API to search your ProKnow workspace, and delete plans >1 year old.
This ensures only the last year of plans are backed up, and storage limit is not reached.
