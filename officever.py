#!/usr/bin/env python2.7

# OfficeVer - Get Microsoft Office version of the document supplied
# Coded by Vilius Povilaika, GitHub: https://github.com/viliuspovilaika
# Download the newest version at https://github.com/viliuspovilaika/OfficeVer

import sys
import os
import shutil

version="0.01.2"

def GetOfficeVersion(versionIntString):
	versionIntString = versionIntString[0:versionIntString.index(".") + 2]
	if versionIntString == "1.0":
		return "Office 1.0"
        elif versionIntString == "1.5":
                return "Office 1.5"
        elif versionIntString == "1.6":
                return "Office 1.6"
        elif versionIntString == "3.0":
                return "Office 3.0"
        elif versionIntString == "4.0":
                return "Office 4.0"
        elif versionIntString == "4.3":
                return "Office 4.3"
        elif versionIntString == "4.2":
                return "Office 4.2 for NT"
        elif versionIntString == "7.0":
                return "Office 95"
        elif versionIntString == "8.0":
                return "Office 97"
        elif versionIntString == "8.5":
                return "Office 97 Powered by Word 98"
        elif versionIntString == "9.0":
                return "Office 2000"
        elif versionIntString == "10.0":
                return "Office XP"
        elif versionIntString == "11.0":
                return "Office 2003"
        elif versionIntString == "12.0":
                return "Office 2007"
        elif versionIntString == "14.0":
                return "Office 2010"
        elif versionIntString == "15.0":
                return "Office 2013"
        elif versionIntString == "16.0":
                return "Office 2016"
	else:
		return "ERR" + versionIntString

if "linux" in sys.platform:
	normal_prefix = "\033[39m"
	blue_prefix = "\033[38;5;69m"
	red_prefix = "\033[31m"
	green_prefix = "\033[38;5;40m"
	errorCode = red_prefix + "[" + normal_prefix + "-" + red_prefix + "] " + normal_prefix
	okCode = blue_prefix + "[" + normal_prefix + "*" +  blue_prefix + "] " + normal_prefix
	goodCode = green_prefix + "[" + normal_prefix + "+" + green_prefix + "] " + normal_prefix
else:
	errorCode = "[-] "
	okCode = "[*] "
	goodCode = "[+] "

def ArgumentError(invalidArg = ""):
	print ""
	if invalidArg != "":
		print errorCode + "Invalid argument " + invalidArg
		print ""
	print okCode + "OfficeVer - Get Microsoft Office version of the document supplied"
	print sys.argv[0] + " [arguments] <filename>"
	print "Available arguments: "
	print "\t--help (-h)\tShow this menu and exit"
	print "\t--version (-v)\tShow officever version and exit"
	print "\t--all (-a)\tShow verbose output"
	print ""
	print "Example usage: " + sys.argv[0] + " -v AnnualReport.doc"
	print ""
	print "Coded by Vilius Povilaika, GitHub: https://github.com/viliuspovilaika"
	print ""
	sys.exit(0)

if len(sys.argv) < 2:
	ArgumentError()

verbose = False
undefinedCounter = 0
lastUndefined = ""
documentPath = ""
for argument in sys.argv:
	if argument == sys.argv[0]:
		nothing=""
	elif argument == "--help" or argument == "-h":
		ArgumentError()
	elif argument == "--version" or argument == "-v":
		print ""
		print okCode + "OfficeVer version " + version
		print ""
		sys.exit(0)
	elif argument == "--all" or argument == "-a":
		verbose = True
	else:
		if undefinedCounter == 0:
			documentPath = argument
		undefinedCounter += 1
		lastUndefined = argument
if undefinedCounter < 1:
	ArgumentError()
elif undefinedCounter > 1:
	ArgumentError(lastUndefined)

print ""
print okCode + "OfficeVer version " + version
if verbose:
	print ""
	print okCode + "Importing needed libraries.."

try:
	import zipfile
except ImportError:
	print ""
	print errorCode + "Zipfile library not found. Please install"
	print ""
	sys.exit(1)

# Main operation

if verbose:
	print okCode + "Extracting the document.."

msVersion = ""
if not os.path.isfile(documentPath):
	print ""
	print errorCode + "File not found!"
	print ""
	sys.exit(1)
try:
	zipObj = zipfile.ZipFile(documentPath, 'r')
except Exception:
	if verbose:
		print okCode + "Can't extract the document, trying to read it.."
	with open(documentPath, 'r') as myfile:
		data=myfile.read().replace('\n', '')
		if "Microsoft Word" in data:
			currentPos = 0
			while True: 
				data = data[currentPos + len("Microsoft Word"):]
				currentPos = data.index("Microsoft Word")
				if data[data.index("Microsoft Word"):data.index("Microsoft Word")+len("Microsoft Word 6.0 or later")] == "Microsoft Word 6.0 or later":
					currentPos = data.index("Microsoft Word")
                                elif data[data.index("Microsoft Word"):data.index("Microsoft Word")+len("Microsoft Word versions 6.0 or later")] == "Microsoft Word versions 6.0 or later":
                                        currentPos = data.index("Microsoft Word")
				else:
					msVersBuffer = data[data.index("Microsoft Word"):data.index("Microsoft Word") + 50]
					break
		else:
			print ""
                        print errorCode + "Version info not found!"
                        print ""
                        sys.exit(0)
	try: msVersBuffer = msVersBuffer[0:msVersBuffer.index("Document")]
	except Exception:
		try: msVersBuffer = msVersBuffer[0:msVersBuffer.index("or later") + len("or later")]
		except Exception:
                    try: nothing = int(msVersBuffer[len("Microsoft Word "):len("Microsoft Word ") + 1])
                    except Exception:
                        msVersBuffer = ""
                    msVersBuffer = msVersBuffer[len("Microsoft Word "):len("Microsoft Word ") + 3]
        OfficeVersion = GetOfficeVersion(msVersBuffer)
        if msVersBuffer == "":
        	print ""
        	print errorCode + "Version info not found!"
        	print ""
                sys.exit(0)
        elif OfficeVersion[:3] == "ERR":
                print ""
                print goodCode + "Version found, but is not in our database: " + OfficeVersion[3:]
                print ""
                sys.exit(0)
	else:
        	print ""
                print goodCode + "Version found: " + OfficeVersion
        	print ""
		sys.exit(0)

if not os.path.isdir("officevers_temp"):
	os.makedirs("officevers_temp")
else:
	shutil.rmtree("officevers_temp")
	os.makedirs("officevers_temp")

zipObj.extractall("officevers_temp")
zipObj.close()

method = 0

if verbose:
	print okCode + "Reading the version info.."

with open('officevers_temp/docProps/app.xml', 'r') as myfile:
	data=myfile.read().replace('\n', '')
if "<AppVersion>" in data:
	msVersion = data[data.index("<AppVersion>") + len("<AppVersion>"):data.index("</AppVersion>")]
        method = 1
elif "<Application>" in data:
        if verbose:
            print okCode + "Version info not found, trying to load the product name.."
        msVersion = data[data.index("<Application>") + len("<Application>"):data.index("</Application")]
        method = 2

if msVersion == "":
	print ""
	print errorCode + "Version info not found!"
	print ""
elif method == 1:
	print ""
        OfficeVersion = GetOfficeVersion(msVersion)
        if OfficeVersion[:3] == "ERR":
            print goodCode + "Version found, but is not in our database: " + OfficeVersion[3:]
        else:
	    print goodCode + "Version found: " + OfficeVersion
	print ""
elif method == 2:
        print ""
        print goodCode + "Version info not found, but the name of the application was found instead: " + msVersion
        print ""
shutil.rmtree("officevers_temp")
