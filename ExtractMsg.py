#!/usr/bin/env python
# -*- coding: latin-1 -*-
"""
ExtractMsg:
    Extracts emails and attachments saved in Microsoft Outlook's .msg files

https://github.com/mattgwwalker/msg-extractor
"""

__author__ = "Matthew Walker"
__date__ = "2016-10-09"
__version__ = '0.3'

# --- LICENSE -----------------------------------------------------------------
#
#    Copyright 2013 Matthew Walker
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.

import os
import sys
import glob
import traceback
from email.parser import Parser as EmailParser
import email.utils
import olefile as OleFile
import base64
import unidecode
import re
from imapclient.imapclient import decode_utf7

# This property information was sourced from
# http://www.fileformat.info/format/outlookmsg/index.htm
# on 2013-07-22.
properties = {
    '001A': 'Message class',
    '0037': 'Subject',
    '003D': 'Subject prefix',
    '0040': 'Received by name',
    '0042': 'Sent repr name',
    '0044': 'Rcvd repr name',
    '004D': 'Org author name',
    '0050': 'Reply rcipnt names',
    '005A': 'Org sender name',
    '0064': 'Sent repr adrtype',
    '0065': 'Sent repr email',
    '0070': 'Topic',
    '0075': 'Rcvd by adrtype',
    '0076': 'Rcvd by email',
    '0077': 'Repr adrtype',
    '0078': 'Repr email',
    '007d': 'Message header',
    '0C1A': 'Sender name',
    '0C1E': 'Sender adr type',
    '0C1F': 'Sender email',
    '0E02': 'Display BCC',
    '0E03': 'Display CC',
    '0E04': 'Display To',
    '0E1D': 'Subject (normalized)',
    '0E28': 'Recvd account1 (uncertain)',
    '0E29': 'Recvd account2 (uncertain)',
    '1000': 'Message body',
    '1008': 'RTF sync body tag',
    '1035': 'Message ID (uncertain)',
    '1046': 'Sender email (uncertain)',
    '3001': 'Display name',
    '3002': 'Address type',
    '3003': 'Email address',
    '39FE': '7-bit email (uncertain)',
    '39FF': '7-bit display name',

    # Attachments (37xx)
    '3701': 'Attachment data',
    '3703': 'Attachment extension',
    '3704': 'Attachment short filename',
    '3707': 'Attachment long filename',
    '370E': 'Attachment mime tag',
    '3712': 'Attachment ID (uncertain)',

    # Address book (3Axx):
    '3A00': 'Account',
    '3A02': 'Callback phone no',
    '3A05': 'Generation',
    '3A06': 'Given name',
    '3A08': 'Business phone',
    '3A09': 'Home phone',
    '3A0A': 'Initials',
    '3A0B': 'Keyword',
    '3A0C': 'Language',
    '3A0D': 'Location',
    '3A11': 'Surname',
    '3A15': 'Postal address',
    '3A16': 'Company name',
    '3A17': 'Title',
    '3A18': 'Department',
    '3A19': 'Office location',
    '3A1A': 'Primary phone',
    '3A1B': 'Business phone 2',
    '3A1C': 'Mobile phone',
    '3A1D': 'Radio phone no',
    '3A1E': 'Car phone no',
    '3A1F': 'Other phone',
    '3A20': 'Transmit dispname',
    '3A21': 'Pager',
    '3A22': 'User certificate',
    '3A23': 'Primary Fax',
    '3A24': 'Business Fax',
    '3A25': 'Home Fax',
    '3A26': 'Country',
    '3A27': 'Locality',
    '3A28': 'State/Province',
    '3A29': 'Street address',
    '3A2A': 'Postal Code',
    '3A2B': 'Post Office Box',
    '3A2C': 'Telex',
    '3A2D': 'ISDN',
    '3A2E': 'Assistant phone',
    '3A2F': 'Home phone 2',
    '3A30': 'Assistant',
    '3A44': 'Middle name',
    '3A45': 'Dispname prefix',
    '3A46': 'Profession',
    '3A48': 'Spouse name',
    '3A4B': 'TTYTTD radio phone',
    '3A4C': 'FTP site',
    '3A4E': 'Manager name',
    '3A4F': 'Nickname',
    '3A51': 'Business homepage',
    '3A57': 'Company main phone',
    '3A58': 'Childrens names',
    '3A59': 'Home City',
    '3A5A': 'Home Country',
    '3A5B': 'Home Postal Code',
    '3A5C': 'Home State/Provnce',
    '3A5D': 'Home Street',
    '3A5F': 'Other adr City',
    '3A60': 'Other adr Country',
    '3A61': 'Other adr PostCode',
    '3A62': 'Other adr Province',
    '3A63': 'Other adr Street',
    '3A64': 'Other adr PO box',

    '3FF7': 'Server (uncertain)',
    '3FF8': 'Creator1 (uncertain)',
    '3FFA': 'Creator2 (uncertain)',
    '3FFC': 'To email (uncertain)',
    '403D': 'To adrtype (uncertain)',
    '403E': 'To email (uncertain)',
    '5FF6': 'To (uncertain)'}


def windowsUnicode(string):
    if string is None:
        return None
    if sys.version_info[0] >= 3:  # Python 3
        return str(string, 'utf_16_le')
    else:  # Python 2
        return unicode(string, 'utf_16_le')


class Attachment:
    def __init__(self, msg, dir_):
        try:
            # Get long filename
           ##print("Attachment() called"+str(dir_))
            if not type(dir_) is list:
                self.longFilename = msg._getStringStream([dir_, '__substg1.0_3707'])

                # Get short filename
                self.shortFilename = msg._getStringStream([dir_, '__substg1.0_3704'])

                # Get attachment data
            
                self.data = msg._getStream([dir_,'__substg1.0_37010102'])
            else:
                self.longFilename = msg._getStringStream([dir_[0],dir_[1],dir_[2], '__substg1.0_3707'])

                # Get short filename
                self.shortFilename = msg._getStringStream([dir_[0],dir_[1],dir_[2], '__substg1.0_3704'])

                # Get attachment data            
                self.data=msg._getStream([dir_[0],dir_[1],dir_[2],'__substg1.0_37010102'])
                #if self.data is None:
                #  #print("data is none:"+str([dir_[0],dir_[1],dir_[2],'__substg1.0_37010102']))

        except Exception as e:
              traceback.format_exc()
            

    def save(self):
        # Use long filename as first preference
        filename = self.longFilename
        # Otherwise use the short filename
        if filename is None:
            filename = self.shortFilename
        # Otherwise just make something up!
        if filename is None:
            import random
            import string
            filename = 'UnknownFilename ' + \
                ''.join(random.choice(string.ascii_uppercase + string.digits)
                        for _ in range(5)) + ".bin"
            
        f = open(filename, 'wb')
        f.write(self.data)
        f.close()
        return filename
    
    def dump(self):
        # Use long filename as first preference
        filename = self.longFilename
        # Otherwise use the short filename
        if filename is None:
            filename = self.shortFilename
        # Otherwise just make something up!
        if filename is None:
            import random
            import string
            filename = 'UnknownFilename ' + \
                ''.join(random.choice(string.ascii_uppercase + string.digits)
                        for _ in range(5)) + ".bin"
            
        
        return {"filename":filename,"data":base64.b64encode(self.data)}


class Message(OleFile.OleFileIO):
    def __init__(self, filename):
        OleFile.OleFileIO.__init__(self, filename)

    def _getStream(self, filename):
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _getStringStream(self, filename, prefer='unicode'):
        """Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then the parameter /prefer/ specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = "/".join(filename)

        asciiVersion = self._getStream(filename + '001E')
        unicodeVersion = windowsUnicode(self._getStream(filename + '001F'))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

    @property
    def subject(self):
        return self._getStringStream('__substg1.0_0037')

    @property
    def header(self):
        try:
            return self._header
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
            else:
                self._header = None
            return self._header

    @property
    def date(self):
        # Get the message's header and extract the date
        if self.header is None:
            return None
        else:
            return self.header['date']

    @property
    def parsedDate(self):
        return email.utils.parsedate(self.date)

    @property
    def sender(self):
        try:
            return self._sender
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["from"]
                if headerResult is not None:
                    self._sender = headerResult
                    return headerResult

            # Extract from other fields
            text = self._getStringStream('__substg1.0_0C1A')
            email = self._getStringStream('__substg1.0_0C1F')
            result = None
            if text is None:
                result = email
            else:
                result = text
                if email is not None:
                    result = result + " <" + email + ">"

            self._sender = result
            for dir_ in self.listdir():
                if dir_[0].startswith("__attach") and dir_[1]=='__substg1.0_3701000D' and dir_[2].startswith("__attach"):
                    #'__substg1.0_3701000D'
                    text = self._getStringStream([dir_[0],dir_[1],dir_[2],'__substg1.0_0C1A'])
                    email = self._getStringStream([dir_[0],dir_[1],dir_[2],'__substg1.0_0C1F'])
                    result = None
                    if text is None:
                        result = email
                    else:
                        result = text
                        if email is not None:
                            result = result + " <" + email + ">"
                    
            return result

    @property
    def to(self):
        try:
            embeddisplay=self._getStringStream(['__attach_version1.0_#00000000', '__substg1.0_3701000D', '__substg1.0_002C0102'])            
            if not None is embeddisplay:
               return embeddisplay
            
            return self._to
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["to"]
                if headerResult is not None:
                    self._to = headerResult
                    return headerResult

            # Extract from other fields
            # TODO: This should really extract data from the recip folders,
            # but how do you know which is to/cc/bcc?
            display = self._getStringStream('__substg1.0_0E04')

            self._to = display
            embeddisplay=self._getStringStream(['__attach_version1.0_#00000000', '__substg1.0_3701000D', '__substg1.0_002C0102'])            
            if not None is embeddisplay:
               display= embeddisplay            
            return display

    @property
    def cc(self):
        try:
            return self._cc
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["cc"]
                if headerResult is not None:
                    self._cc = headerResult
                    return headerResult

            # Extract from other fields
            # TODO: This should really extract data from the recip folders,
            # but how do you know which is to/cc/bcc?
            display = self._getStringStream('__substg1.0_0E03')

            self._cc = display
            return display

    @property
    def body(self):
        # Get the message body
        return self._getStringStream('__substg1.0_1000')
    @property
    def embeddedheader(self):
        catg=self._getStream(['__substg1.0_037020102','__substg1.0_3704001F','__substg1.0_37090102','__substg1.0_371D0102'])

        for dir_ in self.listdir(True):
           ##print("------"+str(dir_)+"------")
           # s=self.openstream(dir_)
         # #print(s.read())
          #  md=self.get_metadata()
         #   md.parse_properties(self)
       #     ms=self.openstream(['__attach_version1.0_#00000000', '__substg1.0_3701000D', '__properties_version1.0'])
          ##print(self.get_metadata().category)
          ##print(self.getproperties(dir_))
      
            if dir_ == ['__attach_version1.0_#00000000', '__substg1.0_3701000D', '__substg1.0_007D001F']:
                #print(dir_)
                hdr=self._getStream(dir_)
                if not None is hdr:
                    hdrlines=hdr.replace("\x00","").replace("'","").strip().splitlines()
                    parsedhdrs=self.parse_smtp_header(hdrlines)
                    hdrdict={}
                    for ln in parsedhdrs:
                        if ":" in ln:
                            k=ln[:ln.index(":")].strip()
                            v=ln[len(k)+1:].strip()
                            if k == "Received":
                                if not "Received" in hdrdict:
                                    hdrdict["Received"]=[]
                                hdrdict[k].append(v)
                            else:
                                hdrdict[k]=v
                    return hdrdict
        #hdr=self._getStream(["__attach_version1.0_#00000000", '__substg1.0_3701000D' ,'__substg11.0_007d'])
        #print(hdr)0 _ 8 0 1 7 1 0 1 F")

        return "Notfound"
    
    @property
    def attachments(self):
        try:

            # Get the attachments
            attachmentDirs = []
            for dir_ in self.listdir():
               ##print dir_
                #for d in dir_:
                   # if self.exists([dir_[0],d]):
                    #  #print("!! "+str([dir_[0],d]))
                if dir_[0].startswith('__attach') and dir_[0] not in attachmentDirs:#  and (dir_[1]=='__substg1.0_37010102' or dir_[2]=='__substg1.0_37010102'):
                    attachmentDirs.append(dir_[0])
                   ##print(">> attachment_dir:"+str(dir_))
                if  dir_[0].startswith('__attach') and dir_[1]=='__substg1.0_3701000D' and dir_[2].startswith('__attach') and [dir_[0],'__substg1.0_3701000D',dir_[2]] not in attachmentDirs:
                    attachmentDirs.append([dir_[0],'__substg1.0_3701000D',dir_[2]])
                    
                   ##print("Appended:"+str([dir_[0],'__substg1.0_3701000D',dir_[2]]))
                    

            self._attachments = []

            for attachmentDir in attachmentDirs:
                self._attachments.append(Attachment(self, attachmentDir))
              ##print("$$"+str(attachmentDir))
            return self._attachments
        except Exception as e:
            traceback.format_exc()
          #print(e)
            return []
    def parse_smtp_header(self,lines):
        if not type(lines) is list:
            return None
        
        lnum=0
        out=[]
        tmpline=''
        p=re.compile("(^[\w-]*):(.*)")
        i=0
        for i in range(len(lines)):
            line=lines[i]
            if len(line.strip())<2:
                continue
            m=re.match(p,line.strip())	
            if not  None is m and not None is m.group(1):
                tmpline+=line
                for j in range(1,len(lines)-i):			
                    m2=re.match(p,lines[i+j])
                    if None is m2 or None is m2.group(1):
                        tmpline+=" "+lines[i+j]
                    else:
                    ##print ("^^"+lines[i+j])
                        i=(i+j)
                        break
            
                out.append(tmpline)
                tmpline=''
            #else:
            ##print("###"+line)
        #out.sort()
        return out

    def save(self, toJson=False, useFileName=False, raw=False):
        '''Saves the message body and attachments found in the message.  Setting toJson
        to true will output the message body as JSON-formatted text.  The body and
        attachments are stored in a folder.  Setting useFileName to true will mean that
        the filename is used as the name of the folder; otherwise, the message's date
        and subject are used as the folder name.'''

        if useFileName:
            # strip out the extension
            dirName = filename.split('/').pop().split('.')[0]
        else:
            # Create a directory based on the date and subject of the message
            d = self.parsedDate
            if d is not None:
                dirName = '{0:02d}-{1:02d}-{2:02d}_{3:02d}{4:02d}'.format(*d)
            else:
                dirName = "UnknownDate"

            if self.subject is None:
                subject = "[No subject]"
            else:
                subject = "".join(i for i in self.subject if i not in r'\/:*?"<>|')

            dirName = dirName + " " + subject

        def addNumToDir(dirName):
            # Attempt to create the directory with a '(n)' appended

            for i in range(2, 100):
                try:
                    newDirName = dirName + " (" + str(i) + ")"
                    os.makedirs(newDirName)
                    return newDirName
                except Exception:
                    traceback.format_exc()
                    pass
            return None

        try:
            if not toJson:
                os.makedirs(dirName)
        except Exception:
            newDirName = addNumToDir(dirName)
            if newDirName is not None:
                dirName = newDirName
            else:
                raise Exception(
                    "Failed to create directory '%s'. Does it already exist?" %
                    dirName
                    )

        oldDir = os.getcwd()
        try:
            if not toJson:
                os.chdir(dirName)

                # Save the message body
                fext = 'json' if toJson else 'text'
            
                f = open("message." + fext, "w")
            # From, to , cc, subject, date

            def xstr(s):
                return '' if s is None else str(s)

            attachmentNames = []
            # Save the attachments
            #print(str(self._getStream(["__attach_version1.0_#00000000", '__substg1.0_37010102'])))
            for attachment in self.attachments:
                if not None is attachment and not None is attachment.data:
                    if not toJson:
                        attachmentNames.append(attachment.save())
                    else:
                        attachmentNames.append(attachment.dump())

            if toJson:
                import json
                

                emailObj = {'from': xstr(self.sender),
                            'to': xstr(self.to),
                            'cc': xstr(self.cc),
                            'subject': xstr(self.subject),
                            'date': xstr(self.date),
                            'attachments': attachmentNames,
                            'body': decode_utf7(self.body)}

                # f.write(json.dumps(emailObj, ensure_ascii=True,indent=4))
               ##print(json.dumps(emailObj, ensure_ascii=True,indent=4))
            else:
                f.write("From: " + xstr(self.sender) + "\n")
                f.write("To: " + xstr(self.to) + "\n")
                f.write("CC: " + xstr(self.cc) + "\n")
                f.write("Subject: " + xstr(self.subject) + "\n")
                f.write("Date: " + xstr(self.date) + "\n")
                f.write("-----------------\n\n")
                f.write(self.body)

                f.close()

        except Exception:
            self.saveRaw()
            raise

        finally:
            # Return to previous directory
            os.chdir(oldDir)
    def dump(self, toJson=False, useFileName=False, raw=False):
        '''Saves the message body and attachments found in the message.  Setting toJson
        to true will output the message body as JSON-formatted text.  The body and
        attachments are stored in a folder.  Setting useFileName to true will mean that
        the filename is used as the name of the folder; otherwise, the message's date
        and subject are used as the folder name.'''
        
        if useFileName:
            # strip out the extension
            dirName = filename.split('/').pop().split('.')[0]
        else:
            # Create a directory based on the date and subject of the message
            d = self.parsedDate
            if d is not None:
                dirName = '{0:02d}-{1:02d}-{2:02d}_{3:02d}{4:02d}'.format(*d)
            else:
                dirName = "UnknownDate"

            if self.subject is None:
                subject = "[No subject]"
            else:
                subject = "".join(i for i in self.subject if i not in r'\/:*?"<>|')

            dirName = dirName + " " + subject

        def addNumToDir(dirName):
            # Attempt to create the directory with a '(n)' appended

            for i in range(2, 100):
                try:
                    newDirName = dirName + " (" + str(i) + ")"
                    os.makedirs(newDirName)
                    return newDirName
                except Exception:
                    traceback.format_exc()
            return None

        try:

            
            def xstr(s):
                #return s
                try:
                    return '' if s is None else str(unidecode.unidecode(s))
                except Exception:
                   ##print("<<<<"+s+">>>>")
                    raise
                    return s
                

            attachmentNames = []
            for attachment in self.attachments:
                if not None is attachment and not None is attachment.data:
                    attachmentNames.append(attachment.dump())

            if toJson:
                import json
                
                emailObj = {'from': xstr(self.sender),
                            'to': xstr(self.to),
                            'cc': xstr(self.cc),
                            'subject': xstr(self.subject),
                            'date': xstr(self.date),
                            'attachments': attachmentNames,
                            'body': decode_utf7(self.body),
                            'embeddedheader':self.embeddedheader
                            }
                return emailObj


        except Exception:
            traceback.format_exc()
            raise


    def saveRaw(self):
        # Create a 'raw' folder
        oldDir = os.getcwd()
        try:
            rawDir = "raw"
            os.makedirs(rawDir)
            os.chdir(rawDir)
            sysRawDir = os.getcwd()

            # Loop through all the directories
            for dir_ in self.listdir():
                sysdir = "/".join(dir_)
                code = dir_[-1][-8:-4]
                global properties
                if code in properties:
                    sysdir = sysdir + " - " + properties[code]
                os.makedirs(sysdir)
                os.chdir(sysdir)

                # Generate appropriate filename
                if dir_[-1].endswith("001E"):
                    filename = "contents.txt"
                else:
                    filename = "contents"

                # Save contents of directory
                f = open(filename, 'wb')
                f.write(self._getStream(dir_))
                f.close()

                # Return to base directory
                os.chdir(sysRawDir)

        finally:
            os.chdir(oldDir)

    #def dump(self):
        ## Prints out a summary of the message
        #print('Message')
        #print('Subject:', self.subject)
        #print('Date:', self.date)
        #print('Body:')
        #print(self.body)

    def debug(self):
        for dir_ in self.listdir():
            if dir_[-1].endswith('001E'):  # FIXME: Check for unicode 001F too
              #print("Directory: " + str(dir))
              #print("Contents: " + self._getStream(dir))
              pass

    def save_attachments(self, raw=False):
        """Saves only attachments in the same folder.
        """
        for attachment in self.attachments:
            attachment.save()

def process_msg(filename):
    try:
        msg=Message(filename)
        
    
        msgdata=msg.dump(True,False)
        #del msgdata["attachments"]
        #del msgdata["body"]
        return msgdata
    except Exception:
        traceback.format_exc()
if __name__ == "__main__":
    if len(sys.argv) <= 1:
      #print(__doc__)
        print("""
Launched from command line, this script parses Microsoft Outlook Message files
and save their contents to the current directory.  On error the script will
write out a 'raw' directory will all the details from the file, but in a
less-than-desirable format. To force this mode, the flag '--raw'
can be specified.

Usage:  <file> [file2 ...]
   or:  --raw <file>
   or:  --json

   to name the directory as the .msg file, --use-file-name
""")
        sys.exit()

    writeRaw = False
    toJson = False
    useFileName = False

    for rawFilename in sys.argv[1:]:
        if rawFilename == '--raw':
            writeRaw = True

        if rawFilename == '--json':
            toJson = True

        if rawFilename == '--use-file-name':
            useFileName = True

        for filename in glob.glob(rawFilename):
            msg = Message(filename)
            try:
                if writeRaw:
                    msg.saveRaw()
                else:
                    msg.save(toJson, useFileName)
            except Exception:
                # msg.debug()
              #print("Error with file '" + filename + "': " +
                traceback.format_exc()
