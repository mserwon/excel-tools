import sys
import os
import json
import re
import argparse
import jinja2
import datetime
from datetime import timedelta
import pandas as pd


# from ccxobjects import *
#from ucxnobjects import *

def dim(a):
    if not type(a) == list:
        return [1]
    return [len(a)] + dim(a[0])

def parse_languagefile(lcode):

    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    dpath_c1 = os.path.join(SITE_ROOT, 'data', f'LanguageMap.json')
    langabbr = ''
    jf = open(dpath_c1, 'r')
    langmap = json.load(jf)
    nl = int(langmap['@total'])
    i = 0
    while i <= nl:
        if langmap['LanguageMapping'][i]['LanguageCode'] == lcode:
            return(langmap['LanguageMapping'][i]['LanguageAbbreviation'])
        else:
            i += 1
    #return default EN-US
    return('ENU')

def parse_tzonefile(tzid):

    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    dpath_c1 = os.path.join(SITE_ROOT, 'data', f'TimeZones.json')
    langabbr = ''
    jf = open(dpath_c1, 'r')
    tzmap = json.load(jf)
    nl = int(tzmap['@total'])
    i = 0
    while i <= nl:
        if tzmap['TimeZone'][i]['TimeZoneId'] == tzid:
            return(tzmap['TimeZone'][i]['DisplayName'])
        else:
            i += 1
    #return default EN-US
    return('ENU')

def parse_ch_rpt_data(ch):
    ch = parse_rpt_basic(ch)
    ch = parse_rpt_menuentries(ch)
    ch = parse_rpt_greetings(ch)
    ch = parse_rpt_transfers(ch)
    ch = parse_rpt_schedules(ch)
    return(ch)

def parse_rpt_basic(ch):
    # modify basic report settings
    ch['Language'] = parse_languagefile(ch['Language'])
    ch['TimeZone'] = parse_tzonefile(ch['TimeZone'])
    return(ch)

def parse_rpt_greetings(ch):
    # modify report settings
    GreetingStream_Init = {
        "@total": "1",
        "PrimaryGreetingStream": "0",
        "GreetingStreamFile": {
          "CallHandlerObjectId": "0",
          "GreetingType": "Standard",
          "LanguageCode": "1033",
          "StreamFile": "Standard.wav",
          "GreetingStreamFileLanguageURI": " ",
          "GreetingStreamFileName": " ",
          "GreetingStreamFilePath": " ",
          "LanguageAbbr": ""
        }
      }
    i = 0
    greetings = ch["Greeting"]
    while i < len(greetings):
        if greetings[i]['AfterGreetingAction'] == '0':
            greetings[i]['AfterGreetingAction'] = 'Ignore Key'
            greetings[i]['AfterGreetingTargetConversation'] = ''
            greetings[i]['DisplayName'] = ''
            greetings[i]['GreetingStream']['GreetingStreamFile']['LanguageAbbr'] = ''
            greetings[i]['GreetingStream']['GreetingStreamFile']['GreetingStreamFilePath'] = ''
        elif greetings[i]['AfterGreetingAction'] == '1':
            greetings[i]['AfterGreetingAction'] = 'Hang Up'
            greetings[i]['AfterGreetingTargetConversation'] = ''
            greetings[i]['AfterGreetingDisplayName'] = ''
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
        elif greetings[i]['AfterGreetingAction'] == '2' and greetings[i]['AfterGreetingTargetConversation'] == 'PHTransfer':
            greetings[i]['AfterGreetingAction'] = 'Transfer to'
            if greetings[i]['AfterGreetingTargetHandler']['IsPrimary'] == 'true':
                greetings[i]['AfterGreetingDisplayName'] = 'User'
            else:
                greetings[i]['AfterGreetingDisplayName'] = 'Call Handler'
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
            greetings[i]['AfterGreetingTargetConversation'] = greetings[i]['AfterGreetingTargetHandler']['DisplayName']
        elif greetings[i]['AfterGreetingAction'] == '2' and greetings[i]['AfterGreetingTargetConversation'] == 'PHGreeting':
            greetings[i]['AfterGreetingAction'] = 'Goto Greeting'
            if greetings[i]['AfterGreetingTargetHandler']['IsPrimary'] == 'true':
                greetings[i]['AfterGreetingDisplayName'] = 'User'
            else:
                greetings[i]['AfterGreetingDisplayName'] = 'Call Handler'
            greetings[i]['AfterGreetingTargetConversation'] = greetings[i]['AfterGreetingTargetHandler']['DisplayName']
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
        elif greetings[i]['AfterGreetingAction'] == '2' and greetings[i]['AfterGreetingTargetConversation'] == 'AD':
            greetings[i]['AfterGreetingAction'] = 'Transfer to'
            greetings[i]['AfterGreetingTargetConversation'] = 'Dial By Name'
            greetings[i]['AfterGreetingDisplayName'] = 'Directory Handler'
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
        elif greetings[i]['AfterGreetingAction'] == '2' and greetings[i]['AfterGreetingTargetConversation'] == 'SubSignIn':
            greetings[i]['AfterGreetingAction'] = 'Send Caller To'
            greetings[i]['AfterGreetingTargetConversation'] = 'Sign-In'
            greetings[i]['AfterGreetingDisplayName'] = 'Call Handler'
            greetings[i]['GreetingStream'] = GreetingStream_Init
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
        elif greetings[i]['AfterGreetingAction'] == '4':
            greetings[i]['AfterGreetingAction'] = 'Take Message'
            greetings[i]['AfterGreetingTargetConversation'] = ''
            greetings[i]['AfterGreetingDisplayName'] = ''
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
        elif greetings[i]['AfterGreetingAction'] == '5':
            greetings[i]['AfterGreetingAction'] = 'Skip Greeting'
            greetings[i]['AfterGreetingTargetConversation'] = ''
            greetings[i]['AfterGreetingDisplayName'] = ''
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
        elif greetings[i]['AfterGreetingAction'] == '6':
            greetings[i]['AfterGreetingAction'] = 'Restart Greeting'
            greetings[i]['AfterGreetingTargetConversation'] = ''
            greetings[i]['AfterGreetingDisplayName'] = ''
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
            #greetings[i]['GreetingStream'] = GreetingStream_Init
        elif greetings[i]['AfterGreetingAction'] == '7':
            greetings[i]['AfterGreetingAction'] = 'Transfer to'
            greetings[i]['AfterGreetingTargetConversation'] = greetings[i]['TransferNumber']
            greetings[i]['AfterGreetingDisplayName'] = ''
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
        elif greetings[i]['AfterGreetingAction'] == '8':
            greetings[i]['AfterGreetingAction'] = 'Goto Next Routing Rule'
            greetings[i]['AfterGreetingTargetConversation'] = ''
            greetings[i]['AfterGreetingDisplayName'] = ''
            if greetings[i]['PlayWhat'] != '1':
                greetings[i]['GreetingStream'] = GreetingStream_Init
        if greetings[i]['PlayWhat'] == '0':
            greetings[i]['PlayWhat'] = 'System Greeting'
        elif greetings[i]['PlayWhat'] == '1':
            greetings[i]['PlayWhat'] = 'Personal Greeting'
        elif greetings[i]['PlayWhat'] == '2':
            greetings[i]['PlayWhat'] = 'Nothing'
        gsn = int(greetings[i]['GreetingStream']['@total'])
        if gsn > 1:
            j = 0 
            while j < gsn:
                if greetings[i]['GreetingStream']['GreetingStreamFile'][j]['LanguageAbbr'] == ch['Language']:
                    greetings[i]['GreetingStream']['PrimaryGreetingStream'] = j
                j +=1
        i +=1
    ch["Greeting"] = greetings
    return (ch)

def parse_rpt_menuentries(ch):
    # modify report settings
    i = 0
    menuentries = ch["MenuEntry"]
    while i < len(menuentries):
        if menuentries[i]['Action'] == '0':
            menuentries[i]['Action'] = 'Ignore Key'
            menuentries[i]['TargetConversation'] = ''
            menuentries[i]['DisplayName'] = ''
        elif menuentries[i]['Action'] == '1':
            menuentries[i]['Action'] = 'Hang Up'
            menuentries[i]['TargetConversation'] = ''
            menuentries[i]['DisplayName'] = ''
        elif menuentries[i]['Action'] == '2' and menuentries[i]['TargetConversation'] == 'PHTransfer':
            menuentries[i]['Action'] = 'Transfer to'
            if menuentries[i]['TargetHandler']['IsPrimary'] == 'true':
                menuentries[i]['DisplayName'] = 'User'
            else:
                menuentries[i]['DisplayName'] = 'Call Handler'
            menuentries[i]['TargetConversation'] = menuentries[i]['TargetHandler']['DisplayName']
        elif menuentries[i]['Action'] == '2' and menuentries[i]['TargetConversation'] == 'PHGreeting':
            menuentries[i]['Action'] = 'Goto Greeting'
            if menuentries[i]['TargetHandler']['IsPrimary'] == 'true':
                menuentries[i]['DisplayName'] = 'User'
            else:
                menuentries[i]['DisplayName'] = 'Call Handler'
            menuentries[i]['TargetConversation'] = menuentries[i]['TargetHandler']['DisplayName']
        elif menuentries[i]['Action'] == '2' and menuentries[i]['TargetConversation'] == 'AD':
            menuentries[i]['Action'] = 'Transfer to'
            menuentries[i]['TargetConversation'] = 'Dial By Name'
            menuentries[i]['DisplayName'] = 'Directory Handler'
        elif menuentries[i]['Action'] == '2' and menuentries[i]['TargetConversation'] == 'SubSignIn':
            menuentries[i]['Action'] = 'Send Caller To'
            menuentries[i]['TargetConversation'] = 'Sign-In'
            menuentries[i]['DisplayName'] = 'Call Handler'
        elif menuentries[i]['Action'] == '4':
            menuentries[i]['Action'] = 'Take Message'
            menuentries[i]['TargetConversation'] = ''
            menuentries[i]['DisplayName'] = ''
        elif menuentries[i]['Action'] == '5':
            menuentries[i]['Action'] = 'Skip Greeting'
            menuentries[i]['TargetConversation'] = ''
            menuentries[i]['DisplayName'] = ''
        elif menuentries[i]['Action'] == '6':
            menuentries[i]['Action'] = 'Restart Greeting'
            menuentries[i]['TargetConversation'] = ''
            menuentries[i]['DisplayName'] = ''
        elif menuentries[i]['Action'] == '7':
            menuentries[i]['Action'] = 'Transfer to'
            menuentries[i]['TargetConversation'] = menuentries[i]['TransferNumber']
        elif menuentries[i]['Action'] == '8':
            menuentries[i]['Action'] = 'Goto Next Routing Rule'
            menuentries[i]['TargetConversation'] = ''
            menuentries[i]['DisplayName'] = ''
        i +=1
    return (ch)

def parse_rpt_transfers(ch):
    # modify transfer option report settings
    i = 0
    transfers = ch["TransferOption"]
    while i < len(transfers):
        if transfers[i]['Enabled'] == 'true':
            if transfers[i]['Action'] == '0':
                transfers[i]['TransferOptionType'] = ''
                transfers[i]['Extension'] = ''
                transfers[i]['TransferType'] = ''
            elif transfers[i]['Action'] == '1':
                transfers[i]['TransferType'] = 'Extension'
            else:
                transfers[i]['TransferOptionType'] = ''
                transfers[i]['Extension'] = ''
                transfers[i]['TransferType'] = ''
        i +=1
    return (ch)

def parse_rpt_schedules(ch):
    # modify transfer option report settings

    schedules = ch["ScheduleDetail"]
    if type (schedules) is list:
        i = 0
        while i < len(schedules):
            if schedules[i]['StartTime'] != '0' or schedules[i]['EndTime'] != '0':
                stime = str(timedelta(minutes=int(schedules[i]['StartTime'])))
                schedules[i]['StartTime'] = stime[:-3]
                etime = str(timedelta(minutes=int(schedules[i]['EndTime'])))
                schedules[i]['EndTime'] = etime[:-3]
            else:
                schedules[i]['Subject'] = 'AllHours'
                schedules[i]['StartTime'] = '00:00'
                schedules[i]['EndTime'] = '24:00'
            #if 'Subject' not in schedules:
                #schedules[i]['Subject'] = ''
            i +=1
        ch["ScheduleDetail"] = schedules
    else:
        schedlist = []
        if schedules['StartTime'] != '0' and schedules['EndTime'] != '0':
            stime = str(timedelta(minutes=int(schedules['StartTime'])))
            schedules['StartTime'] = stime[:-3]
            etime = str(timedelta(minutes=int(schedules['EndTime'])))
            schedules['EndTime'] = etime[:-3]
        else:
            schedules['Subject'] = 'AllHours'
            schedules['StartTime'] = '00:00'
            schedules['EndTime'] = '24:00'
        if 'Subject' not in schedules:
            schedules['Subject'] = ''
        schedlist.append(schedules)
        ch["ScheduleDetail"] = schedlist
    
    return (ch)


def create_ch_report(cluster,searchobj,searchq):
    cs = {0:['Green','xl687995','xl667995'],1:['Red','xl697995','xl707995'],2:['Blue','xl717995','xl727995'],3:['Yellow','xl767995','xl777995']}
    fulljson = f"{searchobj}-{searchq}-full.json"
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, 'data', fulljson)
    tpath = os.path.join(SITE_ROOT, 'templates/CH')
    jf = open(apath_json, 'r')
    chl = json.load(jf)
    rpath = re.sub('.json','.html',apath_json)
    rf = open(rpath,'w')
    print('\nCreate HTML Report - ',rpath)
    templateLoader = jinja2.FileSystemLoader(tpath)
    templateEnv = jinja2.Environment(loader=templateLoader)
    tfile = 'base.html'
    template = templateEnv.get_template(tfile)
    template.globals['now'] = datetime.datetime.utcnow
    section = 'header'
    htext = template.render(section=section, cluster=cluster.upper())
    print(htext, file=rf)
    section = 'rows'
    tfile = 'tablerows.html'
    template = templateEnv.get_template(tfile)
    count = 0
    for count,ch in enumerate(chl):
        if ch['IsPrimary'] == 'false': # IsPrimary = true for Users not Callhandlers
            ch = parse_ch_rpt_data(ch)
            print(ch['DisplayName'],' +')
            htext = template.render(section=section,ch=ch,chm = ch['MenuEntry'],chg = ch['Greeting'],cht = ch['TransferOption'],chs = ch['ScheduleDetail'],cs1=cs[count % 4][1],cs2=cs[count % 4][2])
            print(htext,file=rf)
    tfile = 'footer.html'
    template = templateEnv.get_template(tfile)
    section = 'footer'
    htext = template.render(section=section)
    print(htext, file=rf)
    rf.close()
  
def parse_user_rpt_data(us,dim_us):
    us[0] = parse_rpt_basic_user(us[0]['User'])
    us[0] = parse_rpt_basic_vm(us[0])
    us[4] = parse_rpt_mwi(us[4])
    us[6] = parse_rpt_menuentries(us[6])
    us[6] = parse_rpt_greetings(us[6])
    us[5] = parse_rpt_notifications(us[5])
    us[2] = parse_rpt_message_actions(us[2])
    us[3] = parse_rpt_alt_extensions(us[3])
    return(us)

def parse_rpt_basic_user(us):
    # modify basic report settings
    us['Language'] = parse_languagefile(us['Language'])
    us['TimeZone'] = parse_tzonefile(us['TimeZone'])
    us['Building'] = us['Building'].upper()
    return(us)
    
def parse_rpt_basic_vm(us):
    # modify basic report settings
    if us['LdapType'] == '0':
        us['LdapType'] = 'No'
    else:
        us['LdapType'] = 'Yes'

    us['CosURI'] = 'Standard User'
    return(us)

def parse_rpt_mwi(us):
    # modify mwi report settings

    if us['Mwi']['MwiOn'] == 'true':
        us['Mwi']['MwiOn'] = 'On'
    else:
        us['Mwi']['MwiOn'] = 'Off'
    return(us)

def parse_rpt_alt_extensions(us):
    # modify alternate extension report settings

    if int(us['@total'])> 1:
        # There are alternate extensions
        us['AlternateExtension'][1]['Type'] = 'Work Phone'
        us['AlternateExtension'].pop(0)
        us['AlternateExtension'] = us['AlternateExtension'][0]
    else:
        # There are no alternate extensions
        us['AlternateExtension']['DtmfAccessId'] = ''
    return(us)

def parse_rpt_notifications(us):
    # modify report settings
    i = 0
    notifications = us["NotificationDevice"]
    while i < len(notifications):
        if notifications[i]['Type'] != '8':
            notifications[i]['DisplayName'] = 'Other'
            notifications[i]['EventList'] = ''
            notifications[i]['SmtpAddress'] = ''
        i +=1
    return (us)

def parse_rpt_message_actions(us):
    # modify message actions settings
    maction = {'0': 'Reject', '1': 'Accept', '2': 'Relay', '3': 'Accept and Relay'}

    us['MessageHandler']['VoicemailAction'] = maction[us['MessageHandler']['VoicemailAction']]
    us['MessageHandler']['DeliveryReceiptAction'] = maction[us['MessageHandler']['DeliveryReceiptAction']]

    return(us)


def parse_trigger_output(apath_json):

    cdict = {
        r"'{": r"'",
        r"}'": r"'",
        r"'None'": r"None",
        r" None": r" 'None'",
        r" deque.*\)": r" 'None'",
        r'False,': r'"false",',
        r'True,': r'"true",',
        r"'": r'"',
        r">": r'>"',
        r"<": r'"<',
        r'}\n{': r'},\n{',
        r'}\n  {': r'},\n{'
    }
    cdict = {r'(https.*adminapi\/)': r''}

    with open(apath_json) as ff:
        ffs = ff.read()
    ff.close()

    for key, value in cdict.items():
        ffs, cn = re.subn(key, value, ffs)
        #print(key, ' - ', cn)

    ff = open(apath_json, 'w')
    ff.write(ffs)
    ff.close()

def create_resource_report(jfile,format,searchobj):
    cdict = {r"(^https.*adminapi\/)(.*$)": r"\2"}
    ccxobj = {'skills':'skill','applications':'application','csqs':'csq','resources':'resource','triggers':'trigger','teams':'team'}
    ccxdfcol = {'csqs':[1,2,3,6,7,8,9,10,11,14,17], 'triggers':[1,3,4,6,8,9,10,11,12,13,14,16,17,19,20,26]}
    ccxobjdep = {'csqs':'skills', 'triggers':'applications', 'resources': ['skills','teams']}
    ccxdfcolname = {'userID':'User ID','firstName':'First Name','lastName':'Last Name','extension':'DN','alias':'Alias','type':'Agent/Supv','team.refURL':'Team','skillMap.skillCompetency':'Skills','primarySupervisorOf.suervisorOfTeamName':'Primary Supervisor Of','secondarySupervisorOf.suervisorOfTeamName':'Secondary Supervisor Of'}
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, jfile)
    parse_trigger_output(apath_json)
    dpath_json = os.path.join(SITE_ROOT, 'resources.json')
    jf = open(apath_json, 'r')
    jobj = json.load(jf)
    jf2 = open(dpath_json, 'r')
    jobj2 = json.load(jf2)
    rpath = re.sub('json',format,apath_json)
    print('\nCreate Resource Report - ',rpath)
    df = pd.json_normalize(jobj[searchobj], record_path =[ccxobj[searchobj]])
    df2 = pd.json_normalize(jobj2['skills'], record_path =[ccxobj['skills']])
    df = pd.merge(df,df2, how = 'left', left_on='poolSpecificInfo.skillGroup.skillCompetency.skillNameUriPair.refURL', right_on='self')
    df2['self'].replace(r'(https.*adminapi\/)',r'',regex=True,inplace=True)
    df = pd.merge(df,df2, how = 'left', left_on='application.refURL', right_on='self')
    df = df.iloc[:,ccxdfcol[searchobj]]
    df.rename(ccxdfcolname,axis = 1,inplace=True)

    if format == 'csv':
        df.to_csv(rpath, encoding='utf-8', index = False)
    elif format == 'xlsx':
        df.to_excel(rpath, index = False)

def create_trigger_report(jfile,format,searchobj):
    cdict = {r"(^https.*adminapi\/)(.*$)": r"\2"}
    ccxobj = {'skills':'skill','applications':'application','csqs':'csq','resources':'resource','triggers':'trigger','teams':'team'}
    ccxdfcol = {'csqs':[1,2,3,6,7,8,9,10,11,14,17], 'triggers':[1,3,4,6,8,9,10,11,12,13,14,16,17,19,20,26]}
    ccxobjdep = {'csqs':'skills', 'triggers':'applications'}
    ccxdfcolname = {'directoryNumber':'DN','deviceName':'CTIRP Device','description_x':'CTIRP Description','maxNumOfSessions':'Max Sessions','alertingNameAscii':'CTIRP Alerting Name','devicePool':'Device Pool','location':'Location','partition':'Partition','voiceMailProfile':'VM Profile','callingSearchSpace':'CSS','callingSearchSpaceForRedirect':'Redirect CSS','display':'CTIRP Display','externalPhoneMaskNumber':'External Mask','application.refURL':'Application','callControlGroup.refURL':'Call Control Group','overrideMediaTermination.dialogGroup.refURL':'Dialog Group'}
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, jfile)
    parse_trigger_output(apath_json)
    dpath_json = os.path.join(SITE_ROOT, 'applications.json')
    jf = open(apath_json, 'r')
    jobj = json.load(jf)
    jf2 = open(dpath_json, 'r')
    jobj2 = json.load(jf2)
    rpath = re.sub('json',format,apath_json)
    print('\nCreate Trigger Report - ',rpath)
    df = pd.json_normalize(jobj[searchobj], record_path =[ccxobj[searchobj]])
    df2 = pd.json_normalize(jobj2['applications'], record_path =[ccxobj['applications']])
    df2['self'].replace(r'(https.*adminapi\/)',r'',regex=True,inplace=True)
    df = pd.merge(df,df2, how = 'left', left_on='application.refURL', right_on='self')
    df = df.iloc[:,ccxdfcol[searchobj]]
    df.rename(ccxdfcolname,axis = 1,inplace=True)

    if format == 'csv':
        df.to_csv(rpath, encoding='utf-8', index = False)
    elif format == 'xlsx':
        df.to_excel(rpath, index = False)

def create_application_report(jfile,format,searchobj):
    ccxobj = {'skills':'skill','applications':'application','csqs':'csq','resources':'resource','triggers':'trigger','teams':'team'}
    ccxdfcol = {'skills':[1,2,3]}
    ccxdfcolname = {'id':'CSQ ID','name':'CSQ Name','queueType':'CSQ Type','autoWork':'Auto Work','wrapupTime':'Wrapup Time','resourcePoolType':'Resource Group','serviceLevel':'Service Level','serviceLevelPercentage':'Service Level %','poolSpecificInfo.skillGroup.skillCompetency.competencelevel':'Skill Level','skillName':'Skill Name','poolSpecificInfo.skillGroup.selectionCriteria':'Selection Criteria'}
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, jfile)
    jf = open(apath_json, 'r')
    jobj = json.load(jf)
    rpath = re.sub('json',format,apath_json)
    print('\nCreate Application Report - ',rpath)
    df = pd.json_normalize(jobj[searchobj], record_path =[ccxobj[searchobj]])
    if format == 'csv':
        df.to_csv(rpath, encoding='utf-8', index = False)
    elif format == 'xlsx':
        df.to_excel(rpath, index = False)

def create_csq_report(jfile,format,searchobj):
    ccxobj = {'skills':'skill','applications':'application','csqs':'csq','resources':'resource','triggers':'trigger','teams':'team'}
    ccxdfcol = {'csqs':[1,2,3,6,7,8,9,10,11,14,17]}
    ccxobjdep = {'csqs':'skills'}
    ccxdfcolname = {'id':'CSQ ID','name':'CSQ Name','queueType':'CSQ Type','autoWork':'Auto Work','wrapupTime':'Wrapup Time','resourcePoolType':'Resource Group','serviceLevel':'Service Level','serviceLevelPercentage':'Service Level %','poolSpecificInfo.skillGroup.skillCompetency.competencelevel':'Skill Level','skillName':'Skill Name','poolSpecificInfo.skillGroup.selectionCriteria':'Selection Criteria'}
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, jfile)
    dpath_json = os.path.join(SITE_ROOT, 'skills.json')
    jf = open(apath_json, 'r')
    jobj = json.load(jf)
    jf2 = open(dpath_json, 'r')
    jobj2 = json.load(jf2)
    rpath = re.sub('json',format,apath_json)
    print('\nCreate CSQ Report - ',rpath)
    df = pd.json_normalize(jobj[searchobj], record_path =[ccxobj[searchobj]])
    df2 = pd.json_normalize(jobj2['skills'], record_path =[ccxobj['skills']])
    df = pd.merge(df,df2, how = 'left', left_on='poolSpecificInfo.skillGroup.skillCompetency.skillNameUriPair.refURL', right_on='self')
    df = df.iloc[:,ccxdfcol[searchobj]]
    df.rename(ccxdfcolname,axis = 1,inplace=True)
    colmv = df.pop('Skill Level')
    df.insert(len(df.columns),'Skill Level',colmv)
    #df = pd.merge(df,df2, how = 'left', left_on='Skill Name', right_on='self')

    if format == 'csv':
        df.to_csv(rpath, encoding='utf-8', index = False)
    elif format == 'xlsx':
        df.to_excel(rpath, index = False)

def create_skill_report(jfile,format,searchobj):
    ccxobj = {'skills':'skill','applications':'application','csqs':'csq','resources':'resource','triggers':'trigger','teams':'team'}
    ccxdfcol = {'skills':[1,2,3]}
    ccxdfcolname = {'id':'CSQ ID','name':'CSQ Name','queueType':'CSQ Type','autoWork':'Auto Work','wrapupTime':'Wrapup Time','resourcePoolType':'Resource Group','serviceLevel':'Service Level','serviceLevelPercentage':'Service Level %','poolSpecificInfo.skillGroup.skillCompetency.competencelevel':'Skill Level','skillName':'Skill Name','poolSpecificInfo.skillGroup.selectionCriteria':'Selection Criteria'}
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, jfile)
    jf = open(apath_json, 'r')
    jobj = json.load(jf)
    rpath = re.sub('json',format,apath_json)
    print('\nCreate Skill Report - ',rpath)
    df = pd.json_normalize(jobj[searchobj], record_path =[ccxobj[searchobj]])
    if format == 'csv':
        df.to_csv(rpath, encoding='utf-8', index = False)
    elif format == 'xlsx':
        df.to_excel(rpath, index = False)

def create_us_report(cluster,searchobj,searchq):
    cs = {0:['Green','xl677995','xl657995'],1:['Red','xl717995','xl737995'],2:['Blue','xl747995','xl687995'],3:['Yellow','xl767995','xl777995']}
    fulljson = f"UCXN-{searchobj}-{searchq}-full.json"
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, 'data', fulljson)
    tpath = os.path.join(SITE_ROOT, 'templates/USER')
    jf = open(apath_json, 'r')
    usl = json.load(jf)
    rpath = re.sub('.json','.html',apath_json)
    rf = open(rpath,'w')
    print('\nCreate HTML Report - ',rpath)
    templateLoader = jinja2.FileSystemLoader(tpath)
    templateEnv = jinja2.Environment(loader=templateLoader)
    tfile = 'base.html'
    template = templateEnv.get_template(tfile)
    template.globals['now'] = datetime.datetime.utcnow

    section = 'header'
    dim_us = dim(usl)
    htext = template.render(section=section, cluster=cluster.upper())
    print(htext, file=rf)
    section = 'rows'
    tfile = 'tablerows.html'
    template = templateEnv.get_template(tfile)
    count = 0
    for count,us in enumerate(usl):
        if dim_us[1] == 1:
            us = usl
        us = parse_user_rpt_data(us,dim_us)
        print(us[0]['DisplayName'],' +')
        htext = template.render(section=section,us=us[0],usm = us[6]['MenuEntry'],usg = us[6]['Greeting'],usvm = us[1],usn = us[5]['NotificationDevice'],usmw = us[4]['Mwi'],usma = us[2]['MessageHandler'],usae = us[3]['AlternateExtension'], cs1=cs[count % 4][1],cs2=cs[count % 4][2])
        print(htext,file=rf)
        if dim_us[1] == 1:
            break
    tfile = 'footer.html'
    template = templateEnv.get_template(tfile)
    section = 'footer'
    htext = template.render(section=section)
    print(htext, file=rf)
    rf.close()

def create_ctirp_report(cluster,searchobj,searchq):
    cs = {0:['Green','xl687995','xl667995'],1:['Red','xl697995','xl707995'],2:['Blue','xl717995','xl727995'],3:['Yellow','xl767995','xl777995']}
    fulljson = f"{searchobj}-{searchq}-full.json"
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, 'data', fulljson)
    tpath = os.path.join(SITE_ROOT, 'templates/CH')
    jf = open(apath_json, 'r')
    chl = json.load(jf)
    rpath = re.sub('.json','.html',apath_json)
    rf = open(rpath,'w')
    print('\nCreate HTML Report - ',rpath)
    templateLoader = jinja2.FileSystemLoader(tpath)
    templateEnv = jinja2.Environment(loader=templateLoader)
    tfile = 'base.html'
    template = templateEnv.get_template(tfile)
    template.globals['now'] = datetime.datetime.utcnow
    section = 'header'
    htext = template.render(section=section, cluster=cluster.upper())
    print(htext, file=rf)
    section = 'rows'
    tfile = 'tablerows.html'
    template = templateEnv.get_template(tfile)
    count = 0
    for count,ch in enumerate(chl):
        if ch['IsPrimary'] == 'false': # IsPrimary = true for Users not Callhandlers
            ch = parse_ch_rpt_data(ch)
            print(ch['DisplayName'],' +')
            htext = template.render(section=section,ch=ch,chm = ch['MenuEntry'],chg = ch['Greeting'],cht = ch['TransferOption'],chs = ch['ScheduleDetail'],cs1=cs[count % 4][1],cs2=cs[count % 4][2])
            print(htext,file=rf)
    tfile = 'footer.html'
    template = templateEnv.get_template(tfile)
    section = 'footer'
    htext = template.render(section=section)
    print(htext, file=rf)
    rf.close()

def create_gflow_report(cluster,searchobj,searchq):
    cs = {3:['Green','xl687995','xl667995'],1:['Red','xl697995','xl707995'],2:['Blue','xl717995','xl727995'],0:['Yellow','xl767995','xl777995']}
    fulljson = f"{searchobj}-{searchq}-full.json"
    fullgml = f"{searchobj}-{searchq}-full.gml"
    SITE_ROOT = os.path.realpath(os.path.dirname('.'))
    apath_json = os.path.join(SITE_ROOT, 'data', fulljson)
    gpath = os.path.join(SITE_ROOT, 'data', fullgml)
    tpath = os.path.join(SITE_ROOT, 'templates/CH')
    #jf = open(apath_json, 'r')
    #chl = json.load(jf)
    hpath = re.sub('.gml','.html',gpath)
    rf = open(hpath,'w')
    print('\nCreate HTML Report - ',hpath)
    templateLoader = jinja2.FileSystemLoader(tpath)
    templateEnv = jinja2.Environment(loader=templateLoader)
    tfile = 'base.html'
    template = templateEnv.get_template(tfile)
    template.globals['now'] = datetime.datetime.utcnow
    section = 'header'
    htext = template.render(section=section, cluster=cluster.upper())
    print(htext, file=rf)
    section = 'rows'
    tfile = 'callflow.html'
    template = templateEnv.get_template(tfile)
    count = 0
    for count,ch in enumerate(chl):
        if ch['IsPrimary'] == 'false': # IsPrimary = true for Users not Callhandlers
            ch = parse_ch_rpt_data(ch)
            print(ch['DisplayName'],' +')
            htext = template.render(section=section,ch=ch,chm = ch['MenuEntry'],chg = ch['Greeting'],cht = ch['TransferOption'],chs = ch['ScheduleDetail'],cs1=cs[count % 4][1],cs2=cs[count % 4][2])
            print(htext,file=rf)
    tfile = 'footer.html'
    template = templateEnv.get_template(tfile)
    section = 'footer'
    htext = template.render(section=section)
    print(htext, file=rf)
    rf.close()

def cmd_line_parser():
    cfg = argparse.ArgumentParser(
        description="Uses json files to create reports "
    )
    cfg.add_argument("-o", "--output-type", default="brief",help="Type of the query results brief,json,html")
    cfg.add_argument("-j", "--json-file",default="", help="json input file")
    cfg.add_argument("searchobj", help="Item type to find, user, callhandler, dirn",default="")
    return cfg

def main():
    cfg = cmd_line_parser()
    opt = cfg.parse_args()
    opt.output_type = opt.output_type.lstrip()
    opt.json_file = opt.json_file.lstrip()
    opt.searchobj = opt.searchobj.lstrip()

    qfound = 1
    if opt.searchobj == 'callhandler':
        create_ch_report(opt.cluster,opt.searchobj, opt.searchquery)
    elif opt.searchobj == 'user':
        create_us_report(opt.cluster,opt.searchobj, opt.searchquery)
    elif opt.searchobj == 'dirn':
        create_ch_report(opt.cluster,opt.searchobj, opt.searchquery)
    elif opt.searchobj == 'ctirp':
        create_ctirp_report(opt.cluster,opt.searchobj, opt.searchquery)
    elif opt.searchobj == 'huntp':
        create_ctirp_report(opt.cluster,opt.searchobj, opt.searchquery)
    elif opt.searchobj == 'glfow':
        create_gflow_report(opt.cluster,opt.searchobj, opt.searchquery)
    elif opt.searchobj == 'csqs':
        create_csq_report(opt.json_file,opt.output_type, opt.searchobj)
    elif opt.searchobj == 'skills':
        create_skill_report(opt.json_file,opt.output_type, opt.searchobj)
    elif opt.searchobj == 'resources':
        create_resource_report(opt.json_file,opt.output_type, opt.searchobj)
    elif opt.searchobj == 'teams':
        create_team_report(opt.json_file,opt.output_type, opt.searchobj)
    elif opt.searchobj == 'applications':
        create_application_report(opt.json_file,opt.output_type, opt.searchobj)
    elif opt.searchobj == 'triggers':
        create_trigger_report(opt.json_file,opt.output_type, opt.searchobj)               


if __name__ == '__main__':
    main()