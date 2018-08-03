import base64
import os
import sys
import re
import traceback
import hashlib
import time
import datetime
import json

import requests
import unidecode
import dns.resolver
from win32com.client import constants
from win32com.client.gencache import EnsureDispatch as Dispatch
import yara

import ExtractMsg

CONFIG = {"foldertree": ["Spam", "Inbox"],
          "api_vt": "<vtapikeygoeshere>",
          "rulesdir": "rules",
          "scan_interval": 600,
          "ioc_file": "indicators.csv",
          "ioc_column": 1
          }

UA = "Mozilla/5.0 (Windows NT 10.0; Trident/7.0; rv:11.0) like Gecko"


def yara_filelist(fspath):
    '''
    docstring
    '''
    yara_list = []
    uniq = set()
    for root, subdir, file in os.walk(fspath):
        for fn in file:
            if fn.lower().endswith(".yar") or fn.lower().endswith(".yara"):
                fpath = root+"\\"+fn
                sha256 = ''
                with open(fpath) as f:
                    sha256 = hashlib.sha256(f.read()).hexdigest()
                if not sha256 in uniq:
                    uniq.add(sha256)
                else:
                    continue
                yara_list.append(fpath)
    return yara_list


def load_yara(rule_path):
    '''
    docstring
    '''
    err_yara = []
    good_yara = []
    yarafiles = list(set(yara_filelist(rule_path)))
    for yf in yarafiles:
        try:
            loads = yara.compile(yf)
            good_yara.append(yf)
        except Exception as e:
            print(e)
            err_yara.append(yf)
            try:
                os.remove(yf)
            except Exception:
                pass
            continue
    with open("yara_rules_index.yar", "w+") as yara_index:
        for yf in good_yara:
            #print("Adding yara rule to index: "+yf)
            yara_index.write('include '+'"'+yf+'"\n')
    rules = None
    good_yara = list(set(good_yara))
    recover = False
    recover_remove = ''
    while True:
        try:
            good_dict = {}
            for yf in good_yara:
                good_dict[yf] = yf
            rules = yara.compile(filepaths=good_dict, includes=False)
            break
        except Exception as e:
            estr = str(e)
            yaramatch = re.match(r"("+rule_path+r".*\.yara?).*", estr)
            if not None is yaramatch and not None is yaramatch.group(1):
                yara_file = yaramatch.group(1).strip().replace("/", "\\")
                if "duplicated identifier" in estr:
                    dupm = re.match('.* duplicated identifier \"(.*)\"', estr)
                    dupe = ''
                    if not None is dupm and not None is dupm.group(1):
                        dupe = dupm.group(1)
                        undupe = str(os.urandom(10)).encode(
                            'hex')[:10]+'_'+dupe
                        undupe = undupe[:32]
                        print("Attempting to fix duplicated identifer for " +
                              yara_file+" , identifier "+dupe+". new identifier "+undupe)
                        yara_rule = ''
                        with open(yara_file) as yfl:
                            yara_rule = yfl.read()
                        yara_rule = re.sub(
                            "[^_]"+dupe, " "+undupe, yara_rule, flags=re.MULTILINE | re.UNICODE)
                        # print(yara_rule)
                        with open(yara_file, "w+") as yfl:
                            yfl.write(yara_rule)
                        continue
                else:
                    print(estr)
                    try:
                        with open(yara_file, "w+") as f:
                            f.write("/* removed */\n")
                    except Exception as e:
                        print(e)
                        print("Unable to remove file:"+yara_file)
                    removed = False
                    yfm = yara_file.split(
                        "\\")[len(yara_file.split("\\"))-1].strip()
                    recover_remove = yara_file
                    for yf in good_yara:
                        yf_base = yf.split("\\")[len(yf.split("\\"))-1].strip()
                        #print yf_base+"<>"+yfm
                        if yfm.lower() in yf.lower():
                            print("!!!!! "+yf)
                        if yf_base.endswith(yfm):
                            good_yara.remove(yf)
                            removed = True
                            print("Removed error rule: "+yf+"\t<"+estr+">")
                        if not removed:
                            print("Error:Could not find: "+yfm +
                                  " in index, could not remove it")
                            break
                    rules = None
                    continue
            else:
                print("Could not find a matching yara file in exception:"+estr)
                recover = False
            continue
    return rules


def yara_meta(meta, ruleName, results):
    '''
    docstring
    '''
    to = []
    for k in meta:
        if k.startswith("alert_email"):
            to.append(meta[k])
            break
    if to:
        send_email(to, "Yara rule '"+ruleName+"' matched", results)


def print_yara(matches, context="", msg="", showstrings=True):
    '''
    docstring
    '''
    out = (("*"*40)+context+("*"*40))+"\n"
    out += (msg)+"\n"
    found = False
    for m in matches:
        yara_meta(m.meta, str(m.rule), out)
        found = True
        out += ("Namespace:"+str(m.namespace))+"\n"
        out += ("Meta:"+str(m.meta))+"\n"
        out += ("Rule:"+str(m.rule))+"\n"
        out += ("Tags:"+str(m.tags))+"\n"
        if showstrings:
            out += ("Strings:"+str(m.strings))+"\n"
    out += ("*"*80)+"\n"

    print(out)
    return out


def send_email(to, subject, body):
    '''
    docstring
    '''
    try:
        outlook = Dispatch("Outlook.Application")
        mailitem = outlook.CreateItem(0)
        tostr = ''
        if type(to) is list:
            tostr = ';'.join(to).strip(';').strip().strip(';')
        else:
            tostr = to.strip()
        mailitem.To = tostr
        mailitem.Subject = subject
        mailitem.Body = body
        mailitem.Send()
    except Exception as e:
        print(e)
        traceback.print_exc()


def iocscrape(content):
    '''
    docstring
    '''
    content = content.replace("[.]", ".").replace("hxxp", "http")
    matches = {}
    rex = {}
    rex["URL"] = re.compile(r"(\w+://[\w\.-]{4,}.*)")
    rex["Email"] = re.compile(r"([\w\d\.-]+@[\w\d\.-]+\.[\w\d\.-]+)")
    rex["IP"] = re.compile(r"([0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3})")
    rex["MD5"] = re.compile(r"([a-fA-F\d]{32})")
    rex["SHA1"] = re.compile(r"([a-fA-F\d]{40})")
    rex["SHA224"] = re.compile(r"([a-fA-F\d]{56})")
    rex["SHA384"] = re.compile(r"([a-fA-F\d]{96})")
    rex["SHA256"] = re.compile(r"([a-fA-F\d]{64})")
    rex["SHA512"] = re.compile(r"([a-fA-F\d]{128})")
    rex["Domain"] = re.compile(r"([\w\d\.-]*[\w\d\.-]+\.[a-zA-Z]{2,})")
    for r in sorted(rex):
        m = re.findall(rex[r], content)
    #       print(m)
        if not None is m and m:
            for ioc in m:
                if not r in matches:
                    matches[r] = set()
                matches[r].add(ioc)
    return matches


def vtlookup(filehash):
    '''
    docstring
    '''
    response_code = ""
    try:
        params = {'apikey': CONFIG["api_vt"],
                  'resource': str(filehash).strip()}
        headers = {
            "User-Agent": UA
        }
        response = requests.get('https://www.virustotal.com/vtapi/v2/file/report',
                                params=params, headers=headers)
        if response.status_code != 200 and response.status_code != 204:
            response_code = str(response.status_code)
            raise(Exception("Response code not 200:"+str(response.status_code)))
        elif response.status_code == 204:
            for i in range(0, 30):
                print(
                    "Exceeded api request limit, sleeping for 10 seconds, hash:"+str(filehash))
                time.sleep(10)
                response = requests.get('https://www.virustotal.com/vtapi/v2/file/report',
                                        params=params, headers=headers)
                if response.status_code == 200:
                    break
            if response.status_code != 200:
                print("Giving up due to api timeouts (30 attempts) for hash:" +
                      filehash+" , Final status code:"+str(response.status_code))
                return None
        r = response.json()
        for k in r:
            if not k == "scans":
                print(str(k)+": "+str(r[k]))
            else:
                for scan in r["scans"]:
                    for kv in r["scans"][scan]:
                        if r["scans"][scan]["detected"]:
                            print("\t["+str(scan)+"]"+kv+": " +
                                  str(r["scans"][scan][kv]))
        return r
    except Exception as e:
        print("VTLOOKUP error")
        print(e)
       # traceback.print_exc()
        return None


class IOClist:
    ioclist = []


def external_scans(content, vt=True, binary=False):
    '''
    docstring
    '''
    if None is content:
        return '', False
    matchfound = False
    if not binary:
        content = unidecode.unidecode(content)
    vtdomain = "https://www.virustotal.com/#/domain/"
    vtip = "https://www.virustotal.com/#/ip-address/"
    vthash = "https://www.virustotal.com/#/file/"
    res = ''
    res += "<h1>External api scan results</h1></br>\n"
    sha256 = hashlib.sha256(content).hexdigest()
    if not binary and not vt:
        iocs = iocscrape(content)
        res += "<h2>IOC Listing (Scraped)</h2></br> \n"
        res += "SHA256: <a href='"+vthash+sha256+"'>"+sha256+"</a></br>\n\n"
        for itype in iocs:
            if itype in ["Domain", "IP"]:
                res += "<h3>"+itype+":</h3><ul></br> \n"
                for ioc in list(iocs[itype]):
                    if itype == "Domain":
                        res += "\t<li> <a href='"+vtdomain+ioc+"'>"+ioc+"</a></li>\n"
                    elif itype == "IP":
                        res += "\t<li> <a href='"+vtip+ioc+"'>"+ioc+"</a></li>\n"
                    if IOClist.ioclist:
                        for ln in IOClist.ioclist:
                            if ioc.lower().strip().replace("\"", "") == ln.lower().strip().replace("\"", ""):
                                matchfound = True
                                res += "\t</br>IOC Matched:"+ioc+" >Against>\t"+ln+"</br>"
                res += "</ul></br>\n"
            res += "\n"
    if vt:
        res += "<h2>VirusTotal scan results</h2></br>\n"
        vtres = vtlookup(sha256)
        if not None is vtres:
            res += json.dumps(vtres, indent=4, sort_keys=True)
            if vtres["result"]:
                matchfound = True
        else:
            res += "<h2>VirusTotal lookup error.</h2>"
        # res+=("-"*80)+"\n"
    res += "<hr>\n"
    return res, matchfound


def scan_inbox():
    '''
    docstring
    '''
    global CONFIG
    foldertree = CONFIG["foldertree"]
    outlook = Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    cwd = os.getcwd()
    processed = []
    matchfound = False

    class Oli():
        def __init__(self, outlook_object):
            self._obj = outlook_object

        def items(self):
            array_size = self._obj.Count
            for item_index in xrange(1, array_size+1):
                yield (item_index, self._obj[item_index])

        def prop(self):
            return sorted(self._obj._prop_map_get_.keys())
    rules = None

    def loadrules():
        '''
    docstring
    '''
        rules = load_yara("rules")
        rulecount = 0
        for r in rules:
            rulecount += 1
        print("Loaded "+str(rulecount)+" YARA rules.")
        folderindex = 0

    loadrules()
    loadioc()
    #this needs fixing :/
    outlookfolder = mapi.Folders
    for inx, folder in Oli(outlookfolder).items():
        if folder.Name == foldertree[0]:
            outlookfolder = folder
            print(">"+folder.Name+":")
            break
    for currentfolder in foldertree[1:]:
        for inx, folder in Oli(outlookfolder.Folders).items():
            if folder.Name == currentfolder:
                print("\t>> "+folder.Name)
                outlookfolder = folder
                folderindex = inx
               # break

    try:
        os.mkdir(cwd+"\\workdir")
    except Exception:
        pass
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._mailitem?view = outlook-pia
    for msg in outlookfolder.Items:
        try:
            for attachment in msg.Attachments:
                if attachment.FileName.startswith("Scan_results") and hashlib.sha256(msg.Body).hexdigest() not in processed:
                    print("Removed Scan_results")
                    msg.Attachments.Remove(attachment.Index)
                    msg.Save()
        except Exception as e:
            print(e)
    #raw_input("Press any key to start scanning, Outlook will ask you for grant permission to access the Inbox...")
    while True:
        print("["+datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S') +
              "] Inbox Scan started for "+"/".join(foldertree).strip("/"))
        print("Scanning folder: /".join(foldertree).strip("/"))
        for msg in outlookfolder.Items:
            try:

                msgsha256 = hashlib.sha256(msg.Body).hexdigest()
                if msgsha256 in processed:
                    continue
                else:
                    processed.append(msgsha256)

                yarascan = ''
                external_scan_result = ''
                msgmeta = ''
                yaramatches = None
                senderdomain = sendertxt = senderspf = ''
                try:
                    msgmeta += "SenderEmailAddress:\t" + \
                        unidecode.unidecode(msg.SenderEmailAddress)+"\n"
                    msgmeta += "To:\t"+unidecode.unidecode(msg.To)+"\n"
                    msgmeta += "Subject:\t" + \
                        unidecode.unidecode(msg.Subject)+"\n"
                    msgmeta += "CC:\t"+unidecode.unidecode(msg.CC)+"\n"
                    msgmeta += "Categories:\t" + \
                        unidecode.unidecode(msg.Categories)+"\n"
                except AttributeError:
                    pass

                print("-"*80)
                print(msgmeta)

                if msg.Attachments.Count > 0:
                    #print "."
                    for attachment in msg.Attachments:
                       # print attachment
                        #print attachment.GetTemporaryFilePath()
                        #print attachment.FileName
                        if attachment.FileName:
                            attachment.SaveAsFile(
                                cwd+"\\workdir\\"+attachment.FileName)
                            try:
                                yaramatches = rules.match(
                                    cwd+"\\workdir\\"+attachment.FileName)
                            except Exception as e:
                                pass
                            if yaramatches:
                                matchfound = True
                                yarascan += print_yara(yaramatches, context="Attachment match:"+str(
                                    attachment.FileName)[:40], msg=msgmeta)
                                with open(cwd+"\\workdir\\"+attachment.FileName, "rb") as f:
                                    scan_result, matchfound = external_scans(
                                        f.read(), binary=True)
                                    external_scan_result += scan_result
                            if attachment.FileName.lower().endswith(".msg"):
                                msgdata = ExtractMsg.process_msg(
                                    cwd+"\\workdir\\"+attachment.FileName)
                                m = "MSG found:"+cwd+"\\workdir\\"+attachment.FileName+"\n\tSubject:" + \
                                    str(msgdata["subject"])+"\n\tTo:"+str(msgdata["to"])+"\n\tFrom:"+str(
                                        msgdata["from"])+"\n\tDate:"+str(msgdata["date"])
                                msgmeta += m+"\n"

                                print("MSG attachment:")
                                print(m)
                                # print(msgdata["body"])
                                try:
                                    yaramatches = rules.match(
                                        data=msgdata["body"])
                                except Exception as e:
                                    pass
                                    # traceback.print_exc()
                                if yaramatches:
                                    yarascan += print_yara(
                                        yaramatches, context="MSG attachment body match", msg=msgmeta)
                                    matchfound = True
                                scan_result, matchfound = external_scans(
                                    msgdata["body"], vt=False)
                                external_scan_result += scan_result
                                if not None is msgdata:
                                    for msgattachment in msgdata['attachments']:
                                        # print("Attachment:"+msgattachment["filename"])
                                        with open(cwd+"\\workdir\\__"+attachment.FileName+"__"+msgattachment["filename"], "wb+") as f:
                                            f.write(base64.b64decode(
                                                msgattachment['data']))
                                        try:
                                            yaramatches = rules.match(
                                                cwd+"\\workdir\\__"+attachment.FileName+"__"+msgattachment["filename"])
                                        except Exception as e:
                                            pass
                                        if yaramatches:
                                            matchfound = True
                                            yarascan += print_yara(yaramatches, msg=msgmeta, context="Attachment extracted from MSG attachment matched: "+str(
                                                attachment.FileName+"__"+msgattachment["filename"])[:64])
                                            with open(cwd+"\\workdir\\__"+attachment.FileName+"__"+msgattachment["filename"], "rb") as f:
                                                scan_result, matchfound = external_scans(
                                                    f.read(), binary=True)
                                                external_scan_result += scan_result
                                else:
                                    print("MSGdata is none")
                hbody = unidecode.unidecode(msg.HTMLBody)
                body = unidecode.unidecode(msg.Body)
                for line in hbody.splitlines():
                    if "x-originating-ip" in line.lower():
                        print line

                try:
                    yaramatches = rules.match(data=hbody)
                    yaramatches = rules.match(data=body)
                except Exception as e:
                    pass

                if yaramatches:
                    print '-'*80
                    matchfound = True
                    yarascan += print_yara(yaramatches, context="HTMLBody matched",
                                           msg=msgmeta, showstrings=True)
                    print '-'*80

               # if yaramatches:
                  #  print '-'*80
                    #yarascan += print_yara(yaramatches,context = "Plain body matched",msg = msgmeta,showstrings = True)
                   # print '-'*80

                scan_result, matchfound = external_scans(hbody, vt=False)
                external_scan_result += scan_result
                #print (external_scan_result)
                header = '''
                <html>
                <body>
                '''
                footer = '''
                </body>
                </html>
                '''
                if senderspf:
                    header += "<h1>SPF records</h1></b>"
                    header += "<h3>Domain</h3>:"+senderdomain
                    header += "</br>SPF records:"+senderspf
                if yarascan.strip():
                    yarascan = "<h1>Yara matches</h1></br><pre>"+yarascan+"</pre><hr>"
                scanres = header+yarascan+"</br>\n"+external_scan_result
                resfile = ''
                if matchfound:
                    resfile = "Scan_results_matchfound.htm"
                else:
                    resfile = "Scan_results_nomatches.htm"

                if scanres:
                    with open(cwd+"\\"+resfile, "w+") as f:
                        f.write(scanres)
                    msg.Attachments.Add(cwd+"\\"+resfile, 1, 1, resfile)
                    msg.Save()
            except Exception as e:
                print e
                traceback.print_exc()
        print("["+datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S') +
              "] Done scanning,sleeping for "+str(CONFIG["scan_interval"])+" seconds...")
        time.sleep(int(CONFIG["scan_interval"]))
        loadconfig()
        loadioc()


def loadioc():
    '''
    docstring
    '''
    global CONFIG
    iocfile = ''
    ioccolumn = 1

    if len(sys.argv) > 1:
        iocfile = sys.argv[1]
    elif CONFIG["ioc_file"]:
        iocfile = CONFIG["ioc_file"]
    if CONFIG["ioc_column"]:
        ioccolumn = int(CONFIG["ioc_column"])
    if iocfile:
        try:
            with open(iocfile) as f:
                for ln in f.read().splitlines():
                    columns = ln.split(",")
                    if len(columns) > ioccolumn:
                        ioc = columns[ioccolumn]
                        if not ioc in IOClist.ioclist:
                            IOClist.ioclist.append(ioc)
        except Exception as e:
            print("IOC file "+iocfile+" Could was not loaded.")
    print("Loaded "+str(len(IOClist.ioclist))+" Indicators for scanning.")


def loadconfig():
    '''
    docstring
    '''
    global CONFIG
    with open("config.json") as f:
        CONFIG = json.loads(f.read())


def main():
    '''
    docstring
    '''
    global CONFIG
    loadconfig()
    reload(sys)
    sys.setdefaultencoding("utf-8")
    scan_inbox()


if __name__ == "__main__":
    main()
