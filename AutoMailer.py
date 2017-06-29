from selenium import webdriver
from time import sleep
import json
import requests
import os
import sys

# Login information for outlook
username = "------------------"
password = "--------"

# URL being connected through for the API key
baseURL = "https://outlook.office.com/api/v2.0/me"

# Setting the path for files in which to keep track of emails
root = "C:-------------"

# Have the option to go through mutiple structures, Ie is picked just for example
# driver = webdriver.PhantomJS(os.getcwd()+'/SaaS Application/UniTrace/phantomjs.exe')
# driver = webdriver.Chrome()

driver = webdriver.Ie("C:\\Users\\bypass\\Desktop\\AutoMailer\\IEDriverServer.exe")


# Obtains the access token by opening outlooks sandbox and redeeming it
def get_access_token():
    driver.get(str("https://oauthplay.azurewebsites.net/"))
    driver.find_element_by_id("authorize").click()
    usernameField = driver.find_element_by_id("cred-userid-inputtext")
    usernameField.send_keys(str(username))
    passwordField = driver.find_element_by_id("cred-password-inputtext")
    passwordField.send_keys(str(password))

    driver.find_element_by_id("submit-button").submit()
    driver.find_element_by_id("getToken").click()
    sleep(5)
    responseBody = driver.find_element_by_id('responseBody').text

    responseBody_json = responseBody.split('{')
    driver.close()

    if len(responseBody_json) > 1:
        response = "{" + responseBody_json[1]

        content = json.loads(response)

        return (content['access_token'])


# Accesses outlook's api through the access token
def fetch_emails_from_API(access_token):
    if (access_token):
        client_request_id = 'bc83d76c-ff49-4944-adad-d828050d7e9b'

        headers = {'Authorization': 'bearer ' + access_token,
                   'Accept': 'application/json',
                   'X-AnchorMailbox': username,
                   'client-request-id': client_request_id}

        # Spefecy a spefefic folder for the api to check into
        Folder = "AAMkAGJjNDRhNDEwLTJlNDMtNDAwZS1hNGFkLTExZTYxZmJlYmMzYQAuAAAAAACaNKHe6I8qTLIi0Nl1yM9MAQAJbJYCFuRyQL9u74Hbz%2B2uAAAzYk5vAAA%3D"

        url = baseURL + "/mailfolders/" + Folder + "/messages"

        counter = 0

        # Use's the api to check how many emails are inside the given folder
        APIEmailCounter = (requests.get(url=url + "?$count=true", headers=headers).json()['@odata.count'])

        # Goes through each email and writes it into current folder and then deletes it from outlook
        while (counter < APIEmailCounter):
            counter += 1
            response = requests.get(url=url + "?$top=1&$select=body", headers=headers)

            with open(root + "CurrentEmails/ActualEmail" + str(counter) + ".txt", "w") as outfile:
                outfile.write(str(response.json()['value'][0]['Body']['Content']))

            deleteEmail = requests.delete(url=url + "/" + str(response.json()['value'][0]['Id']), headers=headers)

        # Shows how many emails match our known emails
        Match_counter = 0

        KnownEmails_dir = os.listdir(root + "KnownEmails/")

        # Check's to see if the amount of emails that have been written equal the number of emails that are already known
        if (len(os.listdir(root + "KnownEmails/")) == len(os.listdir(root + "CurrentEmails/"))):

            # Open's the current emails that were taken from outlook and checks to see if they match with an email in the
            # known email
            # known email folder. If they match then they get deleted from the current emails. If there are any emails
            # left they did not match any emails
            for file1 in KnownEmails_dir:
                with open(root + "KnownEmails/" + file1) as knownfile:
                    knownEmail = knownfile.read()

                OtherEmails_dir = os.listdir(root + "CurrentEmails/")
                for file in OtherEmails_dir:
                    with open(root + "CurrentEmails/" + file) as otherfile:
                        otherEmail = otherfile.read()
                    if knownEmail in otherEmail:

                        Match_counter += 1

                        os.remove(root + "CurrentEmails/" + file)
                    else:
                        pass
            print(len(OtherEmails_dir))

            # If all the emails left in the current directory have found a match then the test has passed
            if (len(OtherEmails_dir) == 1):
                write_to_file("PASS")
        else:
            write_to_file("FAIL")
            OtherEmails_dir = os.listdir(root + "CurrentEmails/")
            if (OtherEmails_dir):
                for file in OtherEmails_dir:
                    os.remove(root + "CurrentEmails/" + file)

        if (counter != Match_counter):
            write_to_file("FAIL")


def write_to_file(state):
    file = open(root + "temp.txt", "w+")
    file.write(state)
    file.close()


fetch_emails_from_API(get_access_token())
