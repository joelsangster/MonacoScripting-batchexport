from proknow import ProKnow
import datetime
import urllib3
urllib3.disable_warnings()

'''
Author: Joel Sangster
Date: 29/07/2024

This script finds all the patients within a specified workspace, and checks if the latest plan for each
patient is more than one year old. If so, the patient is deleted from the workspace.
'''

workspace_name = "Name of your workspace"  # workspace_name should be a string exactly matching workspace name.
base_url = 'ProKnow Base Url (example.au.proknow.com)'

credentials_file = "Path to your ProKnow API credentials file."

# Connect with ProKnow server.
pk = ProKnow(base_url, credentials_file=None)
pk.requestor.get('/api', verify=False)  # bypass SSL certificates. 


def is_date_greater_than_one_year_ago(date_str):

    # this method returns True if the date is more than 1 year ago
    current_date = datetime.datetime.now(datetime.timezone.utc)
    one_year_ago = current_date - datetime.timedelta(days=365)
    parsed_date = datetime.datetime.fromisoformat(date_str.replace('Z', '+00:00'))  # Parse ISO 8601 date string
    return parsed_date < one_year_ago


# get all patients
patients = pk.patients.query(workspace_name)  # query() returns a list of 'PatientSummary' objects from the specified workspace.
delete_count = int(0) # counter for deletions

# iterate through patients, and delete if plan date > 1 year ago.
for patient in patients:
    print("")
    print(f'Patient: {patient.name}')
    Patient = patient.get()  # get() returns a 'PatientItem' object.

    # find the patient's plans
    plans = Patient.find_entities(
        lambda entity: entity.data["type"] == "plan")

    if len(plans) == 0:
        print('No plan data..')
        pass  # if there are no plans, ignore it.

    else:
        plan = plans[0].get()  # first plan will be the latest.
        date = plan.data["created_at"]  # get plan date
        print(f'Date: {date}')

        if is_date_greater_than_one_year_ago(date):  # check if plan > one year old.
            print("Deleting..")
            #Patient.delete()  # delete the patient from the workspace
            delete_count += 1  # increment delete count
        else:
            print("Don't Delete")

print("")
print(f'{delete_count} patient(s) deleted from <{workspace_name}> workspace.')
