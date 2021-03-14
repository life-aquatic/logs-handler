import sys
import win32clipboard
import re
import os
from paramiko import Transport, SFTPClient
import zipfile
from stat import S_ISDIR, S_ISREG
import logging
import urllib.request
import shutil

#todo:
#0) unzip not only ".zip". but also ".ZIP"
#1) download not only files, but folders
#2) if there's something in local folder, do sync
#3) option to forcibly redownload all shit
#4) disable debug paramiko logging, add current operation and percent complete
#5) run in background to switch contenxt between cases
#6): correctly handle cases with two upload locations: SFTP Log Locations
#http://supportattachments.aws.cis.local/ticket/04504159; http://syd.supportattachments.aws.cis.local/ticket/04504159



def DownloadAllFromHttp(logFolderUrl, localFolder):
    with urllib.request.urlopen(logFolderUrl) as response:
        webPageContents = response.read().decode()

    logFileUrls = re.findall(r'href="....+"', webPageContents)
    for i in logFileUrls:
        url = logFolderUrl + "/" + i[6:-1]
        file_name = localFolder + i[6:-1]
        print('downloading ' + file_name)
        with urllib.request.urlopen(url) as response2:
            with open(file_name, 'wb') as out_file:
                shutil.copyfileobj(response2, out_file)
        print('downloaded ' + file_name)

class SftpClient:
    _connection = None

    def __init__(self, host, port, username, password):
        self.host = host
        self.port = port
        self.username = username
        self.password = password

        self.create_connection(self.host, self.port,
                               self.username, self.password)

    @classmethod
    def create_connection(cls, host, port, username, password):

        transport = Transport(sock=(host, port))
        transport.connect(username=username, password=password)
        cls._connection = SFTPClient.from_transport(transport)




    def list_sftp(self, remotedir):
        #TODO: I'm appending "upload" here, it's just ugly
        file_list = []
        for entry in self._connection.listdir_attr(remotedir):
            mode = entry.st_mode
            #if S_ISDIR(mode):
            if S_ISREG(mode):
                file_list.append(("/upload/" + entry.filename, entry.filename))
        return file_list



    @staticmethod
    def uploading_info(uploaded_file_size, total_file_size):

        logging.info('uploaded_file_size : {} total_file_size : {}'.
                     format(uploaded_file_size, total_file_size))

    def upload(self, local_path, remote_path):

        self._connection.put(localpath=local_path,
                             remotepath=remote_path,
                             callback=self.uploading_info,
                             confirm=True)

    def file_exists(self, remote_path):

        try:
            print('remote path : ', remote_path)

            # so "_connection" here is an object of type SFTPClient. stat() method returns info about a remote file
            self._connection.stat(remote_path)
        except FileNotFoundError:
            return False
        else:
            return True

    def download(self, remote_path, local_path):

        if self.file_exists(remote_path):
            self._connection.get(remote_path, local_path,
                                 callback=None)


    def close(self):
        self._connection.close()




# get clipboard data
def get_clipboard():
    try:
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        return str(data)
    except:
        return ""


class Case:
    def __init__(self, case_number):
        self.caseNumber = case_number

    def add_sftp_and_folder(self, raw_clip):
        self.folderPath = str.format(r'G:\keys\{}', self.caseNumber) #this line is duplicate in two different methods
        try:
            self.awsPath2 = re.search(r'http.+supportattachments.+\d{8}', raw_clip).group()
        except:
            self.awsPath2 = ""

        try:
            os.makedirs(self.folderPath)
        except FileExistsError:
            # directory already exists
            pass


    def add_sftp_and_folder2(self, raw_clip):
        parsed = re.search("sftp://(.+):(.+)@(.+)/upload", raw_clip)
        self.sftpLogin = parsed.group(1)
        self.sftpPassw = parsed.group(2)
        self.sftpAddr = parsed.group(3)
        self.folderPath = str.format(r'G:\keys\{}', self.caseNumber)
        try:
            os.makedirs(self.folderPath)
        except FileExistsError:
            # directory already exists
            pass



    def __repr__(self):
        return {"Case Number": self.caseNumber,
                "SFTP URL": self.sftpAddr,
                "SFTP Login": self.sftpLogin,
                "SFTP Password": self.sftpPassw}


if __name__ == '__main__':
    runtime_option = sys.argv[1]

    raw_clip = get_clipboard()
    CaseNum = re.search("04\d{6}", raw_clip).group()
    currentCase = Case(CaseNum)


    if runtime_option == "D":
        currentCase.add_sftp_and_folder2(raw_clip)
        download_local_path = 'G:\\keys\\' + currentCase.caseNumber + '\\'
        print("downloading from sftp")
        port = 22
        print(currentCase.sftpPassw)
        print(currentCase.__repr__())
        client = SftpClient(currentCase.sftpAddr, port, currentCase.sftpLogin, currentCase.sftpPassw)

        dirlist = client.list_sftp("./upload")
        for item in dirlist:
            client.download(item[0], download_local_path + item[1])
            print("downloaded " + item[0])
        client.close()
        extension = ".zip"
        os.chdir(download_local_path)  # change directory from working dir to dir with files
        for item in os.listdir(download_local_path):  # loop through items in dir
            if item.endswith(extension):  # check for ".zip" extension
                file_name = os.path.abspath(item)  # get full path of files
                name_for_new_subfolder = os.path.basename(item)[0:17]

                zip_ref = zipfile.ZipFile(file_name)  # create zipfile object
                zip_ref.extractall(download_local_path + '\\' + name_for_new_subfolder)  # extract file to dir
                print("extracted " + file_name)
                zip_ref.close()  # close file

    if runtime_option == "C":
        log_path_string = "Log upload locations:\n" \
                          "1) Customer portal - {}\n" \
                          "2) SFTP - {}\n" \
                          "   SFTP Login: {}\n" \
                          "   SFTP Password: {}".format(currentCase.awsPath,
                                                     currentCase.sftpLinkFull,
                                                     currentCase.sftpLogin,
                                                     currentCase.sftpPassw)
        win32clipboard.OpenClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, log_path_string)
        win32clipboard.CloseClipboard()




    if runtime_option == "Z":
        currentCase.add_sftp_and_folder(raw_clip)

        download_local_path = 'G:\\keys\\' + currentCase.caseNumber + '\\'
        DownloadAllFromHttp(currentCase.awsPath2, download_local_path)


    #input('Press Enter to Continue...')
