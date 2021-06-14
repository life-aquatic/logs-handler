import sys
import win32clipboard
import re
import os
from paramiko import Transport, SFTPClient
import zipfile
from stat import S_ISDIR, S_ISREG, S_IWUSR
import logging
import urllib.request
import shutil
from selenium import webdriver
import time

#todo:
#1) download not only files, but folders
#2) if there's something in local folder, do sync
#3) option to forcibly redownload all shit
#4) disable debug paramiko logging, add current operation and percent complete
#5) run in background to switch contenxt between cases
#6): correctly handle cases with two upload locations: SFTP Log Locations:
#http://supportattachments.aws.cis.local/ticket/04504159; http://syd.supportattachments.aws.cis.local/ticket/04504159



def download_all_from_http(logFolderUrl, localFolder):

    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": localFolder}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(chrome_options=options)

    driver.get(logFolderUrl)
    all_links = driver.find_elements_by_partial_link_text("")
    logFileUrls = [(driver.current_url + "/" + i.text, i.text) for i in all_links[1:]]

    for i in logFileUrls:
        url = i[0]
        file_name = localFolder + "/" + i[1]
        print('starting download of ' + url)
        driver.get(url)
    print('all files queued for download')
    driver.minimize_window()


    # wait for download complete
    wait = True
    while (wait == True):
        atleastone = False
        for item in os.listdir(localFolder):
            if item.lower().endswith('crdownload'):
                atleastone = True
        if atleastone:
            wait = True
        else: wait = False
        print('downloading files ...')
        time.sleep(5)

    print('finished downloading all files ...')
    driver.close()

def unzip_all_in_folder(local_source_path, local_target_path, extension):
    unprocessed_files = []
    os.chdir(local_source_path)  # change directory from working dir to dir with files
    for item in os.listdir(local_source_path):  # loop through items in dir
        if item.lower().endswith(extension):  # check for ".zip" extension

            file_name = os.path.abspath(item)  # get full path of files
            print("extracting from " + file_name)
            name_for_new_subfolder = os.path.basename(item)[0:-4]
            zip_ref = zipfile.ZipFile(file_name)  # create zipfile object
            zip_ref.extractall(local_target_path + '\\' + name_for_new_subfolder)  # extract file to dir
            print("extracted to " + local_target_path + '\\' + name_for_new_subfolder + '\\' + item)
            zip_ref.close()  # close file
        else:
            unprocessed_files.append(item)
    return unprocessed_files

def case_list_from_clipboard(local_path):
    active_cases = re.findall("\d{8}", get_clipboard())
    downloaded_cases = os.listdir(local_path)
    for i in downloaded_cases:
        if i not in active_cases:
            yield i



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
        try:
            self.awsPath2 = re.search(r'http.+supportattachments.+\d{8}', raw_clip).group()

        except:
            self.awsPath2 = ""
        self.folderPath = str.format(r'D:\keys\{}', self.caseNumber)
        self.downloadPath =  str.format(r'C:\Users\Fedor.Nikitin\Downloads\{}', self.caseNumber)
        try:
            os.makedirs(self.folderPath)
            os.makedirs(self.downloadPath)
        except FileExistsError:
            pass


    def add_sftp_and_folder2(self, raw_clip):
        parsed = re.search("sftp://(.+):(.+)@(.+)/upload", raw_clip)
        self.sftpLogin = parsed.group(1)
        self.sftpPassw = parsed.group(2)
        self.sftpAddr = parsed.group(3)
        self.folderPath = str.format(r'D:\keys\{}', self.caseNumber)
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
    local_path = 'D:\\keys\\'
    raw_clip = get_clipboard()
    CaseNum = re.search("04\d{6}", raw_clip).group()
    currentCase = Case(CaseNum)


    if runtime_option == "D":
        currentCase.add_sftp_and_folder2(raw_clip)
        download_local_path = local_path + currentCase.caseNumber + '\\'
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
        unzip_all_in_folder(download_local_path, ".zip")




    if runtime_option == "Z":
        currentCase.add_sftp_and_folder(raw_clip)
        download_all_from_http(currentCase.awsPath2, currentCase.downloadPath)
        unprocessed = unzip_all_in_folder(currentCase.downloadPath, currentCase.folderPath, ".zip")
        if len(unprocessed) > 0:
            for i in unprocessed:
                print(i)
            input("some files were left unprocessed. press any key to close")

    if runtime_option == "SF":
        todelete = [i for i in case_list_from_clipboard(local_path)]
        if input("now I will delete these, ok? " + ", ".join(todelete) + " y/n\n").lower() == 'y':
            for i in todelete:
                try:
                    shutil.rmtree(local_path + '\\' + i)
                    print("deleted " + local_path + '\\' + i)
                except BaseException as error:
                    print ("unable to delete: {}: {}".format(i, error))
                    pass
        input("press Enter to exit")



