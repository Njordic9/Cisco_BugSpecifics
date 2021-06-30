#!/.pyenv/versions/3.7.2/bin/python
"""Module docstring."""
import json
import os
import ssl
import requests
import urllib3
import xlsxwriter
import argparse

ssl._create_default_https_context = ssl._create_unverified_context

urllib3.disable_warnings()  # Suppresses warnings from unsigned Cert warning.


def token(keyuname, keypasswd):
    """Retrieves Product Series and Affect release Bugs from Cisco Bug Search Tool API.
    Returns:
        [JSON] -- Returns all bugs relating to affected release and product series.
    """
    url = "https://cloudsso.cisco.com/as/token.oauth2"
    headers = {'Content-Type': "application/x-www-form-urlencoded"}
    datatype = {"grant_type": "client_credentials"}
    response = requests.post(url, verify=False, stream=True, data=datatype, params={"client_id": keyuname, "client_secret": keypasswd}, headers=headers)
    if response is not None:
        return json.loads(response.text)["access_token"]
    else:
        return None

def get_bugs(token, url):
    """Retrieves Product Series and Affect release Bugs from Cisco Bug Search Tool API.
    Returns:
        [JSON] -- Returns all bugs relating to affected release and product series.
    """
    headers = {
        'Accept': "application/json",
        'Content-Type': "application/json",
        'Authorization': "Bearer {0}".format(token)}
    response = requests.get(url, verify=False, stream=True, headers=headers)
    getstatuscode = response.status_code
    getresponse = response.json()
    if getstatuscode == 200:
        return getresponse
    else:
        response.raise_for_status()

def main(keyuname, keypasswd):
    projectpath = os.getenv("WORKSPACE")
    try:
        os.remove('{}/WeeklyCiscoBugReport.xlsx'.format(projectpath))
    except OSError as e:
        print(e)
    try:
        urldict = {
            "Cisco ASA 5500X Series": "https://api.cisco.com/bug/v2.0/bugs/product_series/Cisco%20ASA%205500-X%20Series%20Firewalls/affected_releases/9.8(3)",
            # "Cisco FTD 5500X Series": "https://api.cisco.com/bug/v2.0/bugs/product_series/Cisco%20ASA%205500-X%20Series%20Firewalls/affected_releases/6.4",
            "Firepower 9000 Series": "https://api.cisco.com/bug/v2.0/bugs/product_series/Cisco%20Firepower%209000%20Series/affected_releases/2.3(1.91)",
            # Removing 4100 series for now, as it's covered by 9000 series in how the bugs are formatted (only lists 9000 series as product series affected).
            # "Firepower 4110 Appliance": "https://api.cisco.com/bug/v2.0/bugs/product_series/Cisco%20Firepower%204110%20Security%20Appliance/affected_releases/2.3(1.91)",
            # "Firepower 4120 Appliance": "https://api.cisco.com/bug/v2.0/bugs/product_series/Cisco%20Firepower%204140%20Security%20Appliance/affected_releases/2.3(1.91)",
            # "Firepower 4140 Appliance": "https://api.cisco.com/bug/v2.0/bugs/product_series/Cisco%20Firepower%204140%20Security%20Appliance/affected_releases/2.3(1.91)",
            "Firepower 7125 Appliance": "https://api.cisco.com/bug/v2.0/bugs/product_name/Cisco%20FirePOWER%20Appliance%207125/affected_releases/6.4",
            "ISE 2.6": "https://api.cisco.com/bug/v2.0/bugs/product_name/Cisco%20Identity%20Services%20Engine/affected_releases/2.6",
            "FMC 4000": "https://api.cisco.com/bug/v2.0/bugs/product_name/Cisco%20Firepower%20Management%20Center%204000/affected_releases/6.4"
            }
        workbook = xlsxwriter.Workbook('{}/WeeklyCiscoBugReport.xlsx'.format(projectpath))
        workbook.close()
        col = 0
        workbook = xlsxwriter.Workbook('{}/WeeklyCiscoBugReport.xlsx'.format(projectpath))
        for x in urldict.items():
            row = 1
            buglist = []
            oauthtoken = token(keyuname, keypasswd)
            bugs = get_bugs(oauthtoken, str(x[1]))
            worksheet = str(x[0])
            worksheet = workbook.add_worksheet(name=str(x[0]))
            bold = workbook.add_format({'bold': 1})
            worksheet.write('A1', '*** {} Bugs ***'.format(str(x[0])), bold)
            for b in bugs['bugs']:
                int_sev = int(b['severity'])
                if int_sev <= 2:
                    buglist.append(b)
                else:
                    continue
            for bb in buglist:
                worksheet.write_string(row, col, "Bug ID : {}: ".format(bb['bug_id']))
                row += 1
                worksheet.write_string(row, col, "Headline : {}: ".format(bb['headline']))
                row += 1
                worksheet.write_string(row, col, "Description : {}: ".format(bb['description']))
                row += 1
                worksheet.write_string(row, col, "Affected Releases : {}: ".format(b['known_affected_releases']))
                row += 1
                worksheet.write_string(row, col, "Fixed Releases : {}: ".format(bb['known_fixed_releases']))
                row += 2
        workbook.close()
    except Exception as e:
        print("ERROR: {}".format(e))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Get Vars')
    parser.add_argument("svc_uname")
    parser.add_argument("svc_passwd")
    args = vars(parser.parse_args())
    keyuname = args["svc_uname"]
    keypasswd = args["svc_passwd"]
    main(keyuname, keypasswd)
