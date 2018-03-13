####
# Created by Bill Talley
# AOS, a ConvergeOne Company
# 913-307-2330
#
# This script has been tested on MacOS High Sierra and Windows 7 using Python 3.6.4, zeep, suds-jurko, xlsxwriter and
# the CUCM AXL files (v10.5 and 11.5).  After installing Python3 from python.org, install the modules using the
# commands below using your OS CLI.
#
#  NOTE: on MacOS run as sudo
#
#  python -m pip install -U pip
#  python -m pip install zeep
#  python -m pip install suds-jurko
#  python -m pip install XlsxWriter
#
#  Download the CUCM AXL Toolkit from the Application >> Plugins menu in CUCM. Save this
#  python script and the extract CUCM AXL 'axlsqltoolkit' directory structure in the same directory.
#  Execute this script from the CLI with the command 'python3 discovery.py' and follow the prompting
#  to perform a discovery of the CUCM/IMP cluster.
#
####

import ssl
import getpass
import os
import time
import datetime
import sys
import urllib.error
import xlsxwriter
from suds.xsd.doctor import Import
from suds.xsd.doctor import ImportDoctor
from suds.client import Client
from collections import Counter

ssl._create_default_https_context = ssl._create_unverified_context

####
# get current working diretory to generate path for saving applicaiton output
# and creates client directory.
####
def createDir(wsdl,location,clientpath,workbook):
    cwd = os.getcwd()
    isDirectory = os.path.isdir(clientpath)
    print("")
    print("Discovery file will be saved to " + clientpath)
    time.sleep(1)
#    print("Client directory exists? " + str(isDirectory))
    if (isDirectory):
        print("")
        print("Client directory already exists at " + clientpath )
        time.sleep(.5)
    else:
        print("")
        print("Creating client directory at " + clientpath )
        time.sleep(.5)
        os.mkdir(clientpath, mode=0o777)
    print("Current working directory: " + cwd )
    time.sleep(.5)
    print("Final Path to configuration files: " + clientpath )
    print("")
    time.sleep(.5)
    login(wsdl,location,clientpath,workbook)

####
# Prompt user for authentication info to authenticate via AXL with CUCM server.  Create
# excel workbook for writing configurating data into.  Query CUCM publisher for node
# and role info, then query each node for active software version.
####

def login(wsdl,location,clientpath,workbook):
    tns = 'http://schemas.cisco.com/ast/soap/'
    imp = Import('http://schemas.xmlsoap.org/soap/encoding/', 'http://schemas.xmlsoap.org/soap/encoding/')
    imp.filter.add(tns)
    origwsdl = wsdl
    try:
        username = input('Please Enter the CUCM username: ')
        password = getpass.getpass('Please Enter the CUCM password: ')
        client = Client(wsdl,location=location, username=username, password=password, plugins=[ImportDoctor(imp)])
        result = client.service.listProcessNode({'name':'%'},{'name':'','nodeUsage':'','processNodeRole':''})
    except Exception as e:
        if str(e) == "(401" + "," + " 'Unauthorized')":
            print("")
            print("Authentication failed!  Please try again. ")
            print("")
            login(wsdl,location,clientpath, workbook)
        else:
            print("")
            print(e + "Please try again. ")
            print("")
            login(wsdl,location,clientpath, workbook)
    else:
        print("")
        print("Authentication Successful! ")
        print("")
        worksheet_nodes = workbook.add_worksheet('Server Nodes')
        time.sleep(1)
        dt = datetime.datetime.now()
        dt = dt.replace(microsecond=0)
        strdt = str(dt)
        strdt = (strdt.replace(':', '').replace('-','').replace(' ','').replace('.',''))
        #file_path = path + '/' + clientname + '_server_nodes_' + strdt + '.csv'
        #status = open(file_path,'w')
        #status.write("Server Name" + "," + "Node Type" + '\n')
        hostlist = []
        count = Counter()
        countCell = Counter()
        cell = 1
        format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
        worksheet_nodes.set_column('A:D', 30)
        worksheet_nodes.write("A1","Server Name", format)
        worksheet_nodes.write("B1","Server Type", format)
        worksheet_nodes.write("C1","Server Role", format)
        worksheet_nodes.write("D1","Server Version", format)
        for node in result['return']['processNode']:
            print (str(node['name']) + ',' + str(node['nodeUsage']) + ',' + str(node['processNodeRole']))
            cell = cell +1
            worksheet_nodes.write("A" + str(cell),str(node['name']))
            worksheet_nodes.write("B" + str(cell),str(node['nodeUsage']))
            worksheet_nodes.write("C" + str(cell),str(node['processNodeRole']))
            count[node.nodeUsage] += 1
            hostlist.append(node['name'])
        countClean = str(count)
        countNew = (countClean.replace('Counter({', '').replace('})',''))
        #hosts = hostlist.split(",")
        cell = 1
        print("")
        for host in hostlist:
            try:
                print("Querying server " + host)
                print("")
                location = 'https://' + host + ':' + cmport + '/realtimeservice/services/RisPort'
                wsdl = 'https://' + host + ':' + cmport + '/realtimeservice/services/RisPort?wsdl'
                client = Client(wsdl,location=location, username=username, password=password, plugins=[ImportDoctor(imp)])
                result = client.service.GetServerInfo()
            except:
                if host == "EnterpriseWideData":
                    print("Entry intentionally skipped.")
                    print("")
                else:
                    print("Unkown error.")
                    print("")
            else:
                if cell == 1:
                    cell = cell +2
                else:
                    cell = cell+1
                result[0]['call-manager-version'] = result[0]['call-manager-version'].replace('.i386', '')
                version = result[0]['call-manager-version'].replace('[[?1;2c','')
                worksheet_nodes.write("D" + str(cell),str(result[0]['call-manager-version']))
                print(result[0]['call-manager-version'] + '\n')
                print("")
        print("")
        print("")
        location = 'https://' + cmserver + ':' + cmport + '/axl/'
        wsdl = origwsdl
        client = Client(wsdl,location=location, username=username, password=password, plugins=[ImportDoctor(imp)])
        mainMenu(wsdl,location,clientpath,username,password,imp,workbook)
    finally:
        print()

####
# present menu options to user
####
def mainMenu(wsdl,location,clientpath,username,password,imp,workbook):
    print("")
    print("Select from the following items: " + '\n')
    #print("1:  Phones" + '\n')
    #print("1:  Phones - SEP Devices Only")
    #print("2:  Gateways - MGCP, SCCP, h323")
    #print("3:  CTI Devices - Route Points and Ports")
    #print("4:  Hunt Pilots")
    #print("5:  Media Devices - CFB, XCODE, MTP")
    print("")
    print("9:  Discover All")
    print("q:  Quit")
    print("")
    selection = input("Your Selection?")
    if selection.isdigit():
        if selection == "1":
            discoverall = False
            phones(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
        elif selection == "2":
            discoverall = False
            gateways(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
        elif selection == "3":
            discoverall = False
            cti(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
        elif selection == "4":
            discoverall = False
            hunt(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
        elif selection == "5":
            discoverall = False
            media(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
        elif selection == "9":
            discoverall = True
            phones(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
        else:
            print("Invalid selection, try again!")
            time.sleep(2)
            mainMenu(wsdl,location,clientpath,username,password,imp,workbook)
    else:
        if selection == "q":
            workbook.close()
            print("")
            print("Goodbye ")
            time.sleep(2)
            exit
        else:
            print("Invalid selection, try again!")
            time.sleep(2)
            mainMenu(wsdl,location,clientpath,username,password,imp,workbook)

####
# Query CUCM via AXL for a list of all SEP phone devices. Create worksheet tab in excel
# workbook and write devicename, description, model and protocol to the worksheet.
####
def phones(wsdl,location,clientpath,username,password,imp,discoverall,workbook):
    print("")
    print("Querying CUCM for a list of phone devices.")
    print("")
    worksheet_phones = workbook.add_worksheet('Phone Devices')
    client = Client(wsdl,location=location, username=username, password=password, plugins=[ImportDoctor(imp)])
    result = client.service.listPhone({'name':'SEP%'},{'name':'','description':'','model':'','protocol':''})
    dt = datetime.datetime.now()
    dt = dt.replace(microsecond=0)
    strdt = str(dt)
    strdt = (strdt.replace(':', '').replace('-','').replace(' ','').replace('.',''))
    count = Counter()
    countCell = Counter()
    cell = 1
    format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
    worksheet_phones.set_column('A:D', 30)
    worksheet_phones.write("A1","Device Name", format)
    worksheet_phones.write("B1","Description", format)
    worksheet_phones.write("C1","Phone Model", format)
    worksheet_phones.write("D1","Protocol", format)
    for node in result['return']['phone']:
        cell = cell +1
        worksheet_phones.write("A" + str(cell),str(node['name']))
        worksheet_phones.write("B" + str(cell),str(node['description']))
        worksheet_phones.write("C" + str(cell),str(node['model']))
        worksheet_phones.write("D" + str(cell),str(node['protocol']))
        print (str(node['name']), str(node['description']), str(node['model']), str(node['protocol']))
        count[node.model] += 1
    print("")
    print("Device Quantities:")
    countClean = str(count)
    countNew = (countClean.replace('Counter({', '').replace('})',''))
    print(countNew)
    cell = cell +2
    worksheet_phones.write("A" + str(cell),str(countNew))
    print("")
    print("Command completed successfully! ")
    print("")
    if discoverall:
        gateways(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
    else:
        mainMenu(wsdl,location,clientpath,username,password,imp,workbook)

####
# Query CUCM via AXL for a list of gateway devices. Create worksheet tab in excel
# workbook and write gateway name, description, model and protocol to the worksheet.
####
def gateways(wsdl,location,clientpath,username,password,imp,discoverall,workbook):
    print("")
    print("Querying CUCM for a list of gateways.")
    print("")
    worksheet_gateways = workbook.add_worksheet('Gateway Devices')
    client = Client(wsdl,location=location, username=username, password=password, plugins=[ImportDoctor(imp)])
    result = client.service.listGateway({'domainName':'%'},{'domainName':'','description':'','product':'','protocol':''})
    dt = datetime.datetime.now()
    strdt = str(dt)
    strdt = (strdt.replace(':', '').replace('-','').replace(' ','').replace('.',''))
    count = Counter()
    countCell = Counter()
    cell = 1
    format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
    worksheet_gateways.set_column('A:D', 30)
    worksheet_gateways.write("A1","Device Name", format)
    worksheet_gateways.write("B1","Description", format)
    worksheet_gateways.write("C1","Product", format)
    worksheet_gateways.write("D1","Protocol", format)
    if result['return'] == "":
        print("No gateway devices configured. ")
        print("")
    else:
        for node in result['return']['gateway']:
            cell = cell +1
            worksheet_gateways.write("A" + str(cell),str(node['domainName']))
            worksheet_gateways.write("B" + str(cell),str(node['description']))
            worksheet_gateways.write("C" + str(cell),str(node['product']))
            worksheet_gateways.write("D" + str(cell),str(node['protocol']))
            print (str(node['domainName']), str(node['description']), str(node['product']), str(node['protocol']))
            count[node.product] += 1
    print("")
    print("Device Quantities:")
    countClean = str(count)
    countNew = (countClean.replace('Counter({', '').replace('})',''))
    print(countNew)
    cell = cell +2
    worksheet_gateways.write("A" + str(cell),str(countNew))
    print("Command completed successfully! ")
    print("")
    if discoverall:
        cti(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
    else:
        mainMenu(wsdl,location,clientpath,username,password,imp,workbook)

####
# Query via AXL for a list of CTI Route Points and CTI ports.  Create worksheet in workbook
# and save Route Point and Port name and description to the worksheet.
####
def cti(wsdl,location,clientpath,username,password,imp,discoverall,workbook):
    print("")
    print("Querying CUCM for a list of CTI Route Points.")
    print("")
    worksheet_cti = workbook.add_worksheet('CTI Devices')
    client = Client(wsdl,location=location, username=username, password=password, plugins=[ImportDoctor(imp)])
    result = client.service.listCtiRoutePoint({'name':'%'},{'name':'','description':''})
    count = Counter()
    countCell = Counter()
    cell = 1
    format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
    worksheet_cti.set_column('A:B', 30)
    worksheet_cti.write("A1","Device Name", format)
    worksheet_cti.write("B1","Description", format)
    if result['return'] == "":
        print("No CTI Route Point devices configured. ")
        print("")
    else:
        for node in result['return']['ctiRoutePoint']:
            cell = cell +1
            worksheet_cti.write("A" + str(cell),str(node['name']))
            worksheet_cti.write("B" + str(cell),str(node['description']))
            print (str(node['name']) + ',' + str(node['description']))
            count['ctiRoutePoint'] += 1
    print("")
    countClean = str(count['ctiRoutePoint'])
    countNew = (countClean.replace('Counter({', '').replace('})',''))
    cell = cell +2
    print("Total CTI Route Points: " + str(countNew))
    worksheet_cti.write("A" + str(cell),"Total CTI Route Points: " + str(countNew))
    print("")
    print("Querying CUCM for a list of CTI Ports.")
    print("")
    print("")
    result = client.service.listPhone({'name':'%'},{'name':'','description':'','model':''})
    cell = cell +4
    format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
    worksheet_cti.set_column('A:B', 30)
    worksheet_cti.write("A" + str(cell),"Device Name", format)
    worksheet_cti.write("B" + str(cell),"Description", format)
    worksheet_cti.write("C" + str(cell),"Phone Model", format)
    if result['return'] == "":
        print("No CTI Port devices configured. ")
        #status.write("No No CTI Port devices configured." + '\n')
        print("")
    else:
        modelCTI = []
        device = 0
        for node in result['return']['phone']:
            if node['model']  == 'CTI Port':
                modelCTI.append(node)
                cell = cell +1
                worksheet_cti.write("A" + str(cell),str(modelCTI[device]['name']))
                worksheet_cti.write("B" + str(cell),str(modelCTI[device]['description']))
                worksheet_cti.write("C" + str(cell),str(modelCTI[device]['model']))
                print(modelCTI[device]['name'] + "," + modelCTI[device]['description'] + "," + modelCTI[device]['model'])
                count['phone'] += 1
                device = device +1
    print("")
    countClean = str(count['phone'])
    countNew = (countClean.replace('Counter({', '').replace('})',''))
    cell = cell +2
    print("Total CTI Ports: " + str(countNew))
    worksheet_cti.write("A" + str(cell),"Total CTI Ports: " + str(countNew))
    print("Command completed successfully! ")
    print("")
    if discoverall:
        hunt(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
    else:
        mainMenu(wsdl,location,clientpath,username,password,imp,workbook)

####
# Query CUCM for a list of hunt pilots, create worksheet tab in workbook and save
# hunt pilot DN and Description to the worksheet.
####
def hunt(wsdl,location,clientpath,username,password,imp,discoverall,workbook):
    print("")
    print("Querying CUCM for a list of hunt pilots.")
    print("")
    worksheet_hunt = workbook.add_worksheet('Hunt Devices')
    client = Client(wsdl,location=location, username=username, password=password, plugins=[ImportDoctor(imp)])
    result = client.service.listHuntPilot({'pattern':'%'},{'pattern':'','description':'','huntListName':''})
    dt = datetime.datetime.now()
    strdt = str(dt)
    strdt = (strdt.replace(':', '').replace('-','').replace(' ','').replace('.',''))
    count = Counter()
    countCell = Counter()
    cell = 1
    format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
    worksheet_hunt.set_column('A:C', 30)
    worksheet_hunt.write("A1","Hunt Pilot", format)
    worksheet_hunt.write("B1","Description", format)
    worksheet_hunt.write("C1","Hunt List", format)
    for node in result['return']['huntPilot']:
        cell = cell +1
        worksheet_hunt.write("A" + str(cell),str(node['pattern']))
        worksheet_hunt.write("B" + str(cell),str(node['description']))
        worksheet_hunt.write("C" + str(cell),str(node['huntListName']['value']))
        print (str(node['pattern']), str(node['description']), str(node['huntListName']['value']))
        count[node.pattern] += 1
    print("")
    countClean = str(count)
    countNew = (countClean.replace('Counter({', '').replace('})',''))
    print(countNew)
    cell = cell +2
    print("Total Hunt Pilots: " + str(countNew))
    worksheet_hunt.write("A" + str(cell),"Total Hunt Pilots: " + str(countNew))
    print("Command completed successfully! ")
    print("")
    if discoverall:
        media(wsdl,location,clientpath,username,password,imp,discoverall,workbook)
    else:
        mainMenu(wsdl,location,clientpath,username,password,imp,workbook)

####
# Query CUCM via AXL for a list of media devices, including CFB, Transcoders and MTP Devices.
# Create worksheet tab in workbook and save meda device name and type to the worksheet.
####
def media(wsdl,location,clientpath,username,password,imp,discoverall,workbook):

    ####
    # CONFERENCE BRIDGES
    ####

    print("")
    print("Querying CUCM for a list of conference bridges.")
    print("")
    worksheet_media = workbook.add_worksheet('Media Devices')
    client = Client(wsdl,location=location, username=username, password=password, plugins=[ImportDoctor(imp)])
    result = client.service.listConferenceBridge({'name':'%'},{'name':'','product':''})
    dt = datetime.datetime.now()
    strdt = str(dt)
    strdt = (strdt.replace(':', '').replace('-','').replace(' ','').replace('.',''))
    count = Counter()
    countCell = Counter()
    cell = 1
    format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
    worksheet_media.set_column('A:C', 30)
    worksheet_media.write("A1","CFB Name", format)
    worksheet_media.write("B1","CFB Type", format)
    for node in result['return']['conferenceBridge']:
        cell = cell +1
        worksheet_media.write("A" + str(cell),str(node['name']))
        worksheet_media.write("B" + str(cell),str(node['product']))
        print (str(node['name']), str(node['product']))
        count['product'] += 1
    print("")
    countClean = str(count['product'])
    countNew = (countClean.replace('Counter({', '').replace('})',''))
    cell = cell +2
    print("Total CFB Devices: " + str(countNew))
    worksheet_media.write("A" + str(cell),"Total CFB Devices: " + str(countNew))
    print("")
    print("Conference bridge query completed successfully! ")
    print("")
    time.sleep(1)

    ####
    # TRANSCODERS
    ####

    print("")
    print("Now querying CUCM for a list of transcoders.")
    print("")
    time.sleep(1)
    result = client.service.listTranscoder({'name':'%'},{'name':'','product':''})
    dt = datetime.datetime.now()
    strdt = str(dt)
    strdt = (strdt.replace(':', '').replace('-','').replace(' ','').replace('.',''))
    countCell = Counter()
    cell = cell +4
    format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
    worksheet_media.write("A" + str(cell),"Transcoder Name", format)
    worksheet_media.write("B" + str(cell),"Transcoder Type", format)
    count['product'] = 0
    if result['return'] == "":
        print("No transcoding devices configured. ")
    else:
        for node in result['return']['transcoder']:
            cell = cell +1
            worksheet_media.write("A" + str(cell),str(node['name']))
            worksheet_media.write("B" + str(cell),str(node['product']))
            print (str(node['name']), str(node['product']))
            count['product'] += 1
    print("")
    countClean = str(count['product'])
    countNew = (countClean.replace('Counter({', '').replace('})',''))
    cell = cell +2
    print("Total Transcoder Devices: " + str(countNew))
    worksheet_media.write("A" + str(cell),"Total Transcoder Devices: " + str(countNew))
    print("")
    print("Transcoder query completed successfully! ")
    print("")
    time.sleep(1)

    ####
    # MTP DEVICES
    ####

    print("")
    print("Now querying CUCM for a list of MTP devices.")
    print("")
    time.sleep(1)
    result = client.service.listMtp({'name':'%'},{'name':'','description':'','mtpType':''})
    dt = datetime.datetime.now()
    strdt = str(dt)
    strdt = (strdt.replace(':', '').replace('-','').replace(' ','').replace('.',''))
    countCell = Counter()
    cell = cell +4
    format = workbook.add_format({'bold': True, 'bg_color': 'silver', 'align': 'center'})
    worksheet_media.write("A" + str(cell),"MTP Name", format)
    worksheet_media.write("B" + str(cell),"Description", format)
    worksheet_media.write("C" + str(cell),"MTP Type", format)
    if result['return'] == "":
        print("No MTP devices configured. ")
    else:
        for node in result['return']['mtp']:
            cell = cell +1
            worksheet_media.write("A" + str(cell),str(node['name']))
            worksheet_media.write("B" + str(cell),str(node['description']))
            worksheet_media.write("C" + str(cell),str(node['mtpType']))
            print (str(node['name']) + "," + str(node['description']) + "," + str(node['mtpType']) )
            count['mtpType'] += 1
    print("")
    countClean = str(count['mtpType'])
    countNew = (countClean.replace('Counter({', '').replace('})',''))
    cell = cell +2
    print("Total MTP Devices: " + str(countNew))
    worksheet_media.write("A" + str(cell),"Total MTP Devices: " + str(countNew))
    print("")
    print("MTP query completed successfully! ")
    print("")
    if discoverall:
        print("Cluster Discovery Completed!")
        workbook.close()
        mainMenu(wsdl,location,clientpath,username,password,imp,workbook)
    else:
        mainMenu(wsdl,location,clientpath,username,password,imp,workbook)

# load login module, prompt user for AXL version, CUCM server address and Discovery
# local file path.
ospath = os.getcwd()
print("")
clientname = input('Please Enter the Client Name: ')
ccmversion = input('Please input the CCM major version(e.g. 10.5): ')
if ccmversion == '9.0':
    print()
elif ccmversion == '9.1':
    print()
elif ccmversion == '10.0':
    print()
elif ccmversion == '10.5':
    print()
elif ccmversion == '11.0':
    print()
elif ccmversion == '11.5':
    print()
else:
    ccmversion == 'Current'
if os.name == "nt":
    ospath = (ospath.replace('\\','/'))
    wsdl = 'file:///' + ospath + '/axlsqltoolkit/schema/' + ccmversion + '/AXLAPI.wsdl'
    clientpath = ospath + '/' + clientname + '/'
elif os.name == "posix":
    wsdl = 'file://' + ospath + '/axlsqltoolkit/schema/' + ccmversion + '/AXLAPI.wsdl'
    clientpath = ospath + '/' + clientname + '/'
    time.sleep(1)
else:
    clientpath = ospath + '/' + clientname + '/'
dt = datetime.datetime.now()
dt = dt.replace(microsecond=0)
strdt = str(dt)
strdt = (strdt.replace(':', '').replace('-','').replace(' ','').replace('.',''))
workbook = xlsxwriter.Workbook(clientpath + clientname + '_discovery_' + strdt + '.xlsx')
cmserver = input('Please enter the CUCM Server IP address: ')
discoverall = False
cmport = '8443'
location = 'https://' + cmserver + ':' + cmport + '/axl/'
createDir(wsdl,location,clientpath,workbook)
