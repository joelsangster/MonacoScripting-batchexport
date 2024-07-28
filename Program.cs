/*************************************************************************************************************************************
 * Batch Export to ProKnow
 * 
 * AUTHOR: JOEL SANGSTER
 * DATE: 16/07/2024
 * INSTITUTION: ST GEORGES CANCER CARE CENTRE
 * 
 * This script automates the export of patient plans to ProKnow
 * Inputs:
 *      - csv file as input with the patient NHIs in the first column.
 *      - config file with Installation, Clinic, and ProKnow workspace defined.
 *      
 * It will export the approved plans for the patients listed in the csv file.
 * 
 * This script includes below steps: 
 * 1. Minimize the windows console window which appears and displays script logs when a Monaco script is started.
 * 2. Launch Monaco
 * 3. Get a list of patients from the provided csv file.
 * 4. For each patient:
 *      - Open the patient
 *      - Load the approved plan
 *      - DICOM Export the Images, Structure Set, and Total Dose to specified workspace.
 *      - Repeat for any other approved plans
 *      - Close Patient.
 * 5. Close Monaco and log any failures
 * 
 ************************************************************************************************************************************/

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Configuration;
using System.Windows.Forms;
using Elekta.MonacoScripting.API;
using Elekta.MonacoScripting.Log;
using Elekta.MonacoScripting.API.DICOMExport;
using Elekta.MonacoScripting.API.General;
using Elekta.MonacoScripting.DataType;
using System.Collections.Generic;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Linq;

namespace BatchDicomExport
{
    class Program
    {
        private static string Installation = ConfigurationManager.AppSettings["Installation"];
        private static string Clinic = ConfigurationManager.AppSettings["Clinic"];
        private static string CSVFilePath = ConfigurationManager.AppSettings["CSVFilePath"];
        private static string ProKnow = ConfigurationManager.AppSettings["ProKnow"];
        

        #region MinimizeCmdWindow
        //Works together with following script step to minimize windows command window during script run
        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName); 
        #endregion


        static void Main(string[] args)
        {
            #region MinimizeCmdWindow
            //Minimize the windows command window which appears during script run to display script logs
            IntPtr intptr = FindWindow("ConsoleWindowClass", null);
            if (intptr != IntPtr.Zero)
                ShowWindow(intptr, 6);
            #endregion
            // Check if the configuration values are not null or empty
            if (string.IsNullOrEmpty(Installation) || string.IsNullOrEmpty(Clinic) ||
                string.IsNullOrEmpty(CSVFilePath) || string.IsNullOrEmpty(ProKnow))
            {
                Logger.Instance.Error("One or more configuration values are missing or empty.");
                return;
            }


            //begin export
            try
            {
                //new Monaco instance
                MonacoApplication app = MonacoApplication.Instance;

                //launch and login to Monaco
                app.LaunchMonaco();


                //Load the CSV file of patient NHIs.
                List<string> PatientIDList = null;
                try
                {
                    // Read all lines from the CSV file to list 
                    PatientIDList = File.ReadAllLines(CSVFilePath).ToList();
                    Logger.Instance.Info("Successfully loaded CSV file.");
                    Logger.Instance.Info(PatientIDList[0]);

                }
                catch (Exception ex)
                {
                    // exceptions (e.g., file not found, access denied)
                    Logger.Instance.Error("Warning message:\n" + ex.Message);
                    app.ExitMonaco();
                }

                List<String> failures = new List<string>(); // initialise list for failed exports. For reporting.               


                //Load each patient on the list one by one
                PatientSelection patient_selection = null;
                foreach (string PatientID in PatientIDList) 
                {
                    Logger.Instance.Info("Loading Patient: " + PatientID + "....");

                    //Load Patient
                    
                    try
                    {
                        // have to do these methods separately, because LoadPatient() method will kill the script if it fails. This way we can use our own logic to handle errors.
                        patient_selection = app.GetPatientSelection();
                        patient_selection.SelectPatient(Installation, Clinic, PatientID);
                        patient_selection.LoadSelectedPatient();
                        if (Warning.Instance.IsVisible()) //close the warning window if open
                        {
                            Warning.Instance.ClickCancel();
                            Logger.Instance.Warn("Could not load patient: " + PatientID);
                            failures.Add(PatientID); //add to failures list.
                            continue; // go to next patient
                        }
                    }
                        
                    

                    catch //exception e.g. if patient open by another user. 
                    {
                        Logger.Instance.Warn("Could not load patient: \n" + PatientID);
                        failures.Add(PatientID); //add to failures list.
                        
                        if (Warning.Instance.IsVisible()) //close the warning window if open
                        {
                            Warning.Instance.ClickCancel();
                        }
                        continue; // go to next patient
                    }
                    


                    //Iterate over plans. Open and export if plan is approved.
                    foreach (var Plan in app.GetPlanList())
                    {
                        if (Plan.Status != PlanStatus.Unapproved)
                        {
                            try
                            {
                                app.LoadPlan(Plan.Name); //load the approved plan
                            }
                            catch (Exception ex) // if plan doesn't load for some reason.
                            {
                                Logger.Instance.Warn("Could not load plan:" + PatientID + Plan.Name + "\n" + ex.Message);
                                failures.Add(PatientID + Plan.Name); //add to failures list
                                if (Warning.Instance.IsVisible())
                                {
                                    Warning.Instance.ClickOK();
                                }
                                continue; //load the next plan
                            }

                            //Open Dicom Export window
                            DicomExport dicomExport = null;
                            try
                            {
                                dicomExport = app.GetDicomExport(MonacoApplication.DicomExportOffsetOption.Continue); //load dicom export window, click OK on offset message
                            }
                            catch (Exception ex) // if dicom export unavailable for some reason.
                            {
                                Logger.Instance.Warn("Could not load plan:" + PatientID + Plan.Name + "\n" + ex.Message);
                                failures.Add(PatientID + Plan.Name);
                                if (Warning.Instance.IsVisible())
                                {
                                    Warning.Instance.ClickOK();
                                }
                                continue; //load the next plan
                            }


                            // Select Images, Structure Set, Total Plan Dose
                            // Ensure only ProKnow Backup export location is selected. 
                            // EDIT THIS SECTION BASED ON YOUR MONACO SETUP.
                            if (dicomExport != null &&
                                dicomExport.SelectDICOMExportModalities(ExportModality.StructureSet | ExportModality.TotalPlanDose | ExportModality.Images) &&
                                dicomExport.ToggleDestination("File", false) &&
                                dicomExport.ToggleDestination("IMPAC", false) &&
                                dicomExport.ToggleDestination("SNC_Clinical", false) &&
                                dicomExport.ToggleDestination("FOCAL", false) &&
                                dicomExport.ToggleDestination(ProKnow, true))
                            {
                                
                                // If the above methods all returned true, Click Export
                                try
                                {
                                    dicomExport.ClickExport();
                                    // go to next plan if succesful. dialog closes automatically. Script will terminate if any export fails... it's built in to the method.
                                }

                                catch
                                {
                                    Logger.Instance.Warn("Export Failed for: \n" + PatientID);
                                    failures.Add(PatientID+Plan.Name); //add plan to failures list

                                    //close the warning dialog, export window, and patient
                                    if (Warning.Instance.IsVisible())
                                    {
                                        Warning.Instance.ClickOK();
                                    }
                                    dicomExport.ClickCancel();
                                    continue; //load the next plan
                                }
                            }

                            // If we can't select a destination or one of the modalities, close dialog and go to next plan
                            else
                            {
                                Logger.Instance.Warn("Export Failed for: \n" + PatientID + "\n Could not select modalities, or could not select destination.");
                                failures.Add(PatientID + Plan.Name); //add plan to failures list
                                if (dicomExport != null)
                                {
                                    dicomExport.ClickCancel();
                                    continue; //load the next plan
                                }
                                else
                                {
                                    continue; //load the next plan
                                }
                            }
                        }
                    }


                    //close patient after all approved plans exported
                    //if successful, export dialog should have closed.
                    
                    app.ClosePatient();

                }
                //Log the failed exports
                foreach (var item in failures)
                {
                    Logger.Instance.Info("Failed Export for: " + item);
                }

              
                app.ExitMonaco();
            }



            // exception if monaco doesn't open.
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception occurred", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }
    }
}
