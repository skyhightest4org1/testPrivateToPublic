from __future__ import absolute_import
#import requests
from builtins import bytes
from builtins import str
from builtins import range
from skybot.OF.lib.Utilities import dlp_requests as requests
import xml.etree.ElementTree as ET
import time, datetime
import timeout_decorator
import os
import re
from skybot.lib.logger import logger
from skybot.OF.lib.Utilities.HealthMonitor import trackme
from skybot.OF.lib.Utilities import Utils

from skybot.OF.lib.core.SkyHighDashboard import Interface
from .ServiceBase import CSP
from robot.api.deco import keyword
from robot.libraries.BuiltIn import BuiltIn
import json
import random
import jinja2
from skybot.lib import SHNInterface
from skybot.OF.lib.core.SkyHighDashboard.ShnDlpInterface import ShnDlpUtil
from skybot.AM.resources.locators import O365_locators
from skybot.lib.web_automation.CommonHelper import CommonHelper
from skybot.lib.web_automation.ActionsHelper import ActionsHelper
from skybot.lib.web_automation.LocatorType import LocatorType
from skybot.lib.web_automation.SyncHelper import SyncHelper
from skybot.OF.lib.Utilities.CustomError import AccesstokenError

#requests.packages.urllib3.disable_warnings()

WS = ""
if os.path.exists("OF/"):
    WS = "OF/"

class OneDrive(CSP):
    known_files = {}
    attempt = 0
    URL_PATTERN_TO_FIND='-my.sharepoint'

    def exclude_external_users(func):
        def exclude_users(instance):
            internal_user = []
            users = func(instance)
            for user in users:
                if user is None or user.lower() == 'none':
                    continue
                if instance.user.split('@')[1].lower() == user.split('@')[1].lower():
                    internal_user.append(user)
            logger.debug("Internal users list" +str(internal_user))
            return internal_user
        return exclude_users

    def __init__(self, shutil, qa_env, tenantid, cspid, user=None, instance_id = None):
        logger.info("Initializing "+str(self.__class__.__name__)+" object...")
        super(OneDrive, self).__init__(qa_env)
        self.user = user
        self.password = None
        # self.email_flat , self.domain_url = [None] * 2
        self.root_folder = None
        self.shutil = shutil
        self.cspid = cspid
        self.instance_id = instance_id
        self.tenantid = tenantid
        #logger.console("Inside onedrive: " + self.access_token)
        ##decrypt_token is set as true to figure out the type of office365 application being installed###
        self.access_token = shutil.get_access_token(cspid, self.instance_id)
        self.access_token_graph = shutil.get_access_token(cspid,self.instance_id, resource="https://graph.microsoft.com", decrypt_token=True)
        #xlogger.console("Inside onedrive: " + self.access_token_graph)
        self.endpoint_GetFolderByServerRelativeUrl = None
        self.endpoint_GetFileByServerRelativeUrl = None
        self.endpoint_users = None
        self.endpoint_contextinfo = None
        self.domain_name = None
        self.list_library_name = self.get_list_name()
        if self._get_domain_name():
            self._get_endpoints()
        if self.access_token is not None:
            self.headers = {"Authorization": "Bearer " + self.access_token, "Accept": "application/json"}
        self.permission_object = None
        self.collaborated_object = None
        self.members_to_collaborate = []
        self.multiple_collaborators = None
        self.external_users = []
        self.same_domain_external_users = []
        self.watchtower_url = SHNInterface.myenv.get_watchtower_url()
        self.stuck_timeout = 45  # timeout value used esp. for methods that are getting stuck
        # (self.request_digest, self.cookies) = self._get_request_digest()
        # Creating tmp directory in root folder to dump the user level XMl files
        if not os.path.exists(WS+"tmp"):
            os.mkdir(WS+"tmp")

        if user is not None:
            self.as_user(user)
        self.list_guid = None
        self.all_collaborators = []
        self.remaining_collaborators =[]
        self.isFlexiLink = None
        self.site=None
        self.response_get_all_users=None


    def as_user(self, user):
        self.user = user
        #logger.console("Running onedrive with user: "+ self.user)
        self.user_flat , self.domain_url = [None] * 2
        if self.domain_name:
            self._get_endpoints()
            (self.request_digest, self.cookies) = self._get_request_digest()
        # self._get_endpoints()
        # (self.request_digest, self.cookies) = self._get_request_digest()

    def _refresh_token(self):
        logger.debug("Refreshing access token ...")
        self.access_token = self.shutil.get_access_token(self.cspid,self.instance_id)
        # logger.console("Inside onedrive: " + self.access_token)
        self.access_token_graph = self.shutil.get_access_token(self.cspid, self.instance_id,resource="https://graph.microsoft.com")
        if self.access_token is not None:
            self.headers = {"Authorization": "Bearer " + self.access_token, "Accept": "application/json"}
        self.user_flat , self.domain_url = [None] * 2
        self._get_endpoints()
        (self.request_digest, self.cookies) = self._get_request_digest()

    # Files method
    def upload_file(self, filename, parent_id = 0, overwrite=True):
        """
        To upload the file to Onedrive.

        Args:
            parent_id: Id of the folder where file required to be uploaded
            filename: filename with location to be uploaded
            overwrite (Optional) : Should we update file if it exists. Default is True

        Returns:
            Id: File ID of the uploaded file.

        Raises:
            None
        """
        super(OneDrive, self).upload_file(filename)
        file_id = None
        retry = 3
        filename = str(filename)
        if self.testdata not in filename:
            filename = self.testdata + "/" + filename
        if not overwrite:
            file, ext = filename.split('/')[-1].split('.')
            new_file = file + str(time.time()).split('.')[0]+ '.' + ext
            filename = '/'.join(filename.split('/')[:-1]) + '/' + new_file
        # headers = {"X-RequestDigest": self.request_digest, "Content-Type": self.get_mime_type(filename=os.path.basename(filename)),
        #             "Accept": "application/json"}
        parent_id, self.mostrecentfolder = [0 if self.mostrecentfolder is None else self.mostrecentfolder] * 2
        if parent_id:
            endpoint_GetFolderByServerRelativeUrl = re.sub('GetFolderByServerRelativeUrl(.*)',
                                                                'GetFolderByServerRelativeUrl(\'' + parent_id + '\')',
                                                                    self.endpoint_GetFolderByServerRelativeUrl)
        else:
            endpoint_GetFolderByServerRelativeUrl = self.endpoint_GetFolderByServerRelativeUrl
        upload_url = endpoint_GetFolderByServerRelativeUrl + '/Files/add(url=\'' + os.path.basename(filename) \
                        + '\', overwrite=true)'
        logger.info("URL to upload is " + upload_url)

        with open(filename, "rb") as fp:
            if filename not in OneDrive.known_files:
                OneDrive.known_files[filename] = fp.read()

        # if filename not in OneDrive.known_files:
        #     OneDrive.known_files[filename] = open(filename, 'rb')#"".join(file_handle.readlines())
        #     data_to_upload = OneDrive.known_files[filename].read()

        #logger.console("loading to memory")
        #file_handle = open(filename, "rb")

        for i in range(retry):
            headers = {"X-RequestDigest": self.request_digest,"Content-Type": self.get_mime_type(filename=os.path.basename(filename)),"Accept": "application/json"}

            logger.debug("Size of the data to uplaod >>>>"+ str(len(OneDrive.known_files[filename])))
            response_upload_file = requests.post(upload_url, headers=headers, cookies=self.cookies, data=OneDrive.known_files[filename])
            self.response = response_upload_file

            if (response_upload_file.status_code == 200):
                # file_size_before_upload = os.path.getsize(OneDrive.known_files[filename].name)
                # logger.debug(">>>>[Before Upload] file {0} with file size {1}".format(OneDrive.known_files[filename].name,file_size_before_upload))

                file_size_after_upload = json.loads(response_upload_file.text)['Length']
                logger.debug(">>>>[After Upload] file file size is {0}".format(file_size_after_upload))
                if (int(file_size_after_upload) >0):
                    logger.info("Response for upload file is " + str(response_upload_file.text))
                    break
                else:
                    file_id=response_upload_file.json()["ServerRelativeUrl"]
                    self.delete_file(file_id)
                    logger.warn("File size is 0 byte so retrying .. ")

            elif response_upload_file.status_code in [401, 403]:
                logger.debug("Got a 403 error refreshing access token...")
                self._refresh_token()

            else:
                logger.console("status code is not matched {0}".format(response_upload_file.status_code))
                break


            #response_upload_file = requests.post(upload_url, headers=headers, cookies=self.cookies, data=OneDrive.known_files[filename], timeout=1)
        if isinstance(response_upload_file, dict):
            if response_upload_file.get('status_code') == 504:
                logger.warn('request timed out for {0} after {1} sec, however we will return True '
                                'assuming, post request is successful!!'.format("upload_file", self.stuck_timeout))
                assumed_file_id = str(parent_id + '/' + str(filename).split("/")[-1])
            self.lastuploadedfiles.append(
                {
                    "fileid": assumed_file_id,
                    "filename": str(filename).split("/")[-1],
                    "folderid": parent_id,
                    "quarantineref": '/personal/' + self.user.replace("@", '_').replace('.', '_') + ':' + str(assumed_file_id),
                    "permissions_object": {"id": self.mostrecentfolder, "permissions_list": None}
                }
            )
            test_name = BuiltIn().replace_variables('${TEST_NAME}')
            BuiltIn().set_suite_metadata(test_name + "_" + str(self.instance_id) + "_lastuploadedfiles",
                                         self.lastuploadedfiles)
            return assumed_file_id

        logger.info("Response post upload file is: " + response_upload_file.text)
        if "ServerRelativeUrl" in response_upload_file.json():
            logger.info("File " + os.path.basename(filename) + " is successfully uploaded")
            file_id = response_upload_file.json()["ServerRelativeUrl"]
        self.lastuploadedfiles.append(
                                        {
                                            "fileid":str(file_id),
                                            "filename":str(filename).split("/")[-1],
                                            "folderid": parent_id,
                                            "quarantineref":'/personal/' + self.user.replace("@",'_').replace('.', '_') + ':' + str(file_id),
                                            "permissions_object": {"id": self.mostrecentfolder, "permissions_list": None}

                                        }
                                     )
        logger.info("Files uploaded thus far: " + str(self.lastuploadedfiles))
        test_name = BuiltIn().replace_variables('${TEST_NAME}')
        BuiltIn().set_suite_metadata(test_name + "_" + str(self.instance_id) + "_lastuploadedfiles", self.lastuploadedfiles)
        return file_id

    def update_file(self, file_id, filename):
        """
        To update the file when we know file id

        Args:
            file_id: Id of the file to be updated
            filename: filename with location to be updated with

        Returns:
            Id: File ID of the updated file.

        Raises:
            None
        """
        return NotImplementedError

    @keyword("download last uploaded file in")
    def download_file(self, file_id=None):
        """
        To download the file when we know file id or name

        Args:
            file_to_check: structured information about the file to download filename, fileid and folderid
            length: The number of bytes to receive, default or -1 is receive all data
        Returns:
            file: content of the file. as byte array

        """
        headers = {"X-RequestDigest": self.request_digest}
        if file_id is None:
            file_id = self.lastuploadedfiles[-1].get('fileid')

        logger.debug("Going to download file with file_id: " + str(file_id))

        endpoint_GetFileByServerRelativeUrl = re.sub('GetFileByServerRelativeUrl(.*)',
                                                     'GetFileByServerRelativeUrl(\'' + file_id + '\')',
                                                     self.endpoint_GetFileByServerRelativeUrl)
        get_file_data_url = endpoint_GetFileByServerRelativeUrl + "/$value"
        request_headers = self.headers.copy()  # create a copy of the headers dict for use only in this function
        r = requests.get(url=get_file_data_url, headers=request_headers, cookies=self.cookies, stream=True)
        if r.status_code in [401, 403]:
            logger.warn("Got a 40X error with content: " + r.text)
            self._refresh_token()
            return False
        if r.status_code in [200, 206]:
            logger.info("file downloaded successfully: " + str(len(r.content)) + " bytes")
            self.lastdownloadedfilecontents = bytes(r.content)
            return True
        else:
            logger.debug("error while downloading file")
            return False

    @keyword("move last uploaded file to")
    def move_file(self, file_id=None):
        """

        :param file_id:
        :return:
        """
        headers = {"X-RequestDigest": self.request_digest}
        filename = ""

        if file_id is None:
            file_id = self.lastuploadedfiles[-1].get('fileid')
            filename = self.lastuploadedfiles[-1].get('filename')

        move_filename = filename
        logger.debug("Going to move file with file_id: " + str(file_id))
        move_folder_id = self.create_folder("random")
        endpoint_GetFileByServerRelativeUrl = re.sub('GetFileByServerRelativeUrl(.*)','GetFileByServerRelativeUrl(\'' + file_id + '\')',self.endpoint_GetFileByServerRelativeUrl)
        file_move_url = endpoint_GetFileByServerRelativeUrl + "/moveto" + "(newurl='" + str(move_folder_id) + "/" + move_filename + "')"
        response_move_file = requests.post(url=file_move_url, headers=headers, cookies=self.cookies)
        if response_move_file.status_code == 200:
            logger.info("Response for upload file is " + str(response_move_file.text))
            self.lastuploadedfiles.append(
                                        {
                                            "fileid":str(file_id),
                                            "filename":str(move_filename).split("/")[-1],
                                            "folderid": move_folder_id,
                                            "quarantineref":'/personal/' + self.user.replace("@",'_').replace('.', '_') + ':' + str(file_id),
                                            "permissions_object": {"id": self.mostrecentfolder, "permissions_list": None}

                                        }
                                     )
        elif response_move_file.status_code in [401,403]:
            logger.debug("Got a 403 error refreshing access token...")
            self._refresh_token()
            raise Exception
        logger.debug("Response post move file is: " + response_move_file.text)

        return file_id

    @keyword("delete last uploaded file in")
    def delete_file(self, file_id=None):
        """
        To delete the file using file id

        Args:
            file_id: Id of the file to be deleted.
                    e.g. "/personal/admin_ak5_onmicrosoft_com/Documents/ShareIt/DeleteMe.txt"

        Returns:
            file: True if able to delete else False

        Raises:
            None
        """
        if file_id is None:
            file_id = self.lastuploadedfiles[-1].get('fileid')

        logger.debug('Going to delete filename with file_id  '+str(file_id))
        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
        is_file_deleted = False

        delete_url = re.sub('GetFileByServerRelativeUrl(.*)',
                                                                'GetFileByServerRelativeUrl(\'' + file_id + '\')',
                                                                    self.endpoint_GetFileByServerRelativeUrl)
        logger.info("Url to delete the file " + file_id + " is " + delete_url)
        response_delete_file = requests.delete(delete_url, headers=headers, cookies=self.cookies)
        if response_delete_file:
            is_file_deleted = True
        return is_file_deleted

    def get_file_info(self, file_id, list_item_all_fields=None):
        """
        get info of the file. will return json of the file property given it's file id
        :param file_id: file id
        :return: JSON
        """
        logger.debug("Inside get_file_info")
        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
        endpoint_GetFileByServerRelativeUrl = re.sub('GetFileByServerRelativeUrl(.*)',
                                                                'GetFileByServerRelativeUrl(\'' + file_id + '\')',
                                                                    self.endpoint_GetFileByServerRelativeUrl)
        if self.site:
            endpoint_GetFileByServerRelativeUrl = re.sub('(.*)\/_api', self.domain_url + '/sites/' + self.site + '/_api', \
                                                           endpoint_GetFileByServerRelativeUrl)
        get_file_info_url = endpoint_GetFileByServerRelativeUrl + list_item_all_fields
        response_file_info = requests.get(url=get_file_info_url, headers=headers, cookies=self.cookies)
        if response_file_info.status_code == 200:
            logger.info("Response for getting file info is " + str(response_file_info.text))

        elif response_file_info.status_code in [401,403]:
            logger.debug("Got a 403 error..")
            self._refresh_token()
            if site.self:
                req_dig = self.req_digest(self.domain_url + "/sites/" + self.site)
                headers["X-RequestDigest"] = req_dig
            raise Exception

        logger.debug("Response from requesting file info for " + str(file_id) + " + is " + str(response_file_info.text))
        return response_file_info.json()

    # Folders method

    def create_folder(self, folder_name, parent_id=None,site=None):
        """
        To create folder given parent id and folder name

        Args:
            parent_id: Id of the folder where folder needs to be created e.g. ( folder1/folder2 )
            folder_name: Name of folder to be created

        Returns:
            folder id

        Raises:
            None
        """
        super(OneDrive, self).create_folder(folder_name, parent_id)
        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
        logger.info("Request digest is " + str(self.request_digest))
        logger.info("Cookies is: " + str(self.cookies))
        folder_id = None
        if folder_name is None:
            logger.error("Folder name is None")
            return False
        if str(folder_name) == "random":
            folder_name = str(time.time())
        logger.info("Inside create folder")
        if parent_id:
            endpoint_GetFolderByServerRelativeUrl = re.sub('GetFolderByServerRelativeUrl(.*)',
                                                                'GetFolderByServerRelativeUrl(\'' + parent_id + '\')',
                                                                    self.endpoint_GetFolderByServerRelativeUrl)
        else:
            endpoint_GetFolderByServerRelativeUrl = self.endpoint_GetFolderByServerRelativeUrl

        if site:
            self.site = site[type(self).__name__]
            endpoint_GetFolderByServerRelativeUrl = re.sub('(.*)\/_api',
                                                           self.domain_url + '/sites/' + self.site + '/_api',
                                                           endpoint_GetFolderByServerRelativeUrl)
            req_dig = self.req_digest(self.domain_url + "/sites/" + self.site)
            headers["X-RequestDigest"] = req_dig
        else:
            self.site = None


        create_folder_url = endpoint_GetFolderByServerRelativeUrl + '/Folders/add(url=\'' + folder_name + '\')'
        logger.info("Create Folder url is " + create_folder_url)
        response_create_folder = requests.post(create_folder_url, headers=headers, cookies=self.cookies)
        self.response = response_create_folder
        if response_create_folder.status_code == 200:
            logger.info("Response for creating folder is " + str(response_create_folder.text))

        elif response_create_folder.status_code in [401,403]:
            logger.debug("Got a 403 error..")
            self._refresh_token()
            if response_create_folder.status_code==403 and OneDrive.attempt<3:
                OneDrive.attempt = OneDrive.attempt + 1
                self.create_folder(folder_name, parent_id)
            raise Exception

        if "ServerRelativeUrl" in response_create_folder.json():
            logger.info("Folder is created successfully with ServerRelativeUrl as " + response_create_folder.json()["ServerRelativeUrl"])
            folder_id = response_create_folder.json()["ServerRelativeUrl"]
        self.mostrecentfolder = folder_id
        self.mostrecentfoldername = folder_name
        test_name = BuiltIn().replace_variables('${TEST_NAME}')
        BuiltIn().set_suite_metadata(test_name + "_" + str(self.instance_id) + "_mostrecentfolder", self.mostrecentfolder)
        BuiltIn().set_suite_metadata(test_name + "_" + str(self.instance_id) + "_mostrecentfoldername", self.mostrecentfoldername)
        return folder_id

    def delete_folder(self, folder_id):
        """
        Delete the folder given folder id. Yet to be implemented

        Args:
            folder_id: Id of the folder to be deleted

        Returns:
            Boolean

        Raises:
            None
        """
        logger.debug("Going to delete the folder with id: " + folder_id)
        if not folder_id:
            folder_id = self.mostrecentfolder
        if folder_id:
            endpoint_GetFolderByServerRelativeUrl = re.sub('GetFolderByServerRelativeUrl(.*)',
                                                                'GetFolderByServerRelativeUrl(\'' + folder_id + '\')',
                                                                    self.endpoint_GetFolderByServerRelativeUrl)
        else:
            endpoint_GetFolderByServerRelativeUrl = self.endpoint_GetFolderByServerRelativeUrl

        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
        response_folder_delete = requests.delete(url=endpoint_GetFolderByServerRelativeUrl, headers=headers, cookies=self.cookies)
        # response_folder_delete = requests.delete(url=folder_id, headers=headers, cookies=self.cookies)
        if response_folder_delete.status_code == 200:
            logger.info("Folder Deleted Success => " + str(folder_id))

        elif response_folder_delete.status_code in [401,403]:
            logger.debug("Got a 403 error..")
            self._refresh_token()
            raise Exception

    def get_folder_info(self, folder_id=None, params=""):
        """
        Get info about the folder

        Args:
            folder_id: Id of the folder to be deleted

        Returns:
            List containing JSON of file properties

        Raises:
            None
        """
        logger.debug("inside get_folder_info")
        if not folder_id:
            folder_id = self.mostrecentfolder
        if folder_id:
            endpoint_GetFolderByServerRelativeUrl = re.sub('GetFolderByServerRelativeUrl(.*)',
                                                                'GetFolderByServerRelativeUrl(\'' + folder_id + '\')',
                                                                    self.endpoint_GetFolderByServerRelativeUrl)
        else:
            endpoint_GetFolderByServerRelativeUrl = self.endpoint_GetFolderByServerRelativeUrl
        logger.debug("Inside get_folder_info with params: folder_id" + str(folder_id))
        get_folder_url = endpoint_GetFolderByServerRelativeUrl + params
        for attempt in range(1, 3):
            headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
            response_folder_info = requests.get(url=get_folder_url, headers=headers, cookies=self.cookies)
            logger.debug(response_folder_info.json())
            if response_folder_info.status_code == 200:
                logger.debug(response_folder_info.json())
                return response_folder_info.json()
            elif response_folder_info.status_code in [401, 403]:
                logger.info("Token expired, refreshing and trying again")
                self._refresh_token()
            else:
                logger.warn("Failed to get folder info, will retry after 10 sec: " + str(attempt) + "/3")
                time.sleep(10)
                continue
        logger.error("Failed to get folder info post attempt 3 times: " + str(folder_id))
        raise Exception

    # Permission methods

    def list_permissions(self, object_id):
        """
        List permissions of an object . Currently implemented for folder only ( yet to implement )

        Args:
            object_id: Id of the object for which permissions required to be listed

        Returns:
            JSON of all permissions been part of the object e.g ["email": "user1@ak5.onmicrosoft.com", "role": "editor"]
            RoleType mappings reference:
                https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.roletype%28v=office.14%29.aspx

        Raises:
            None
        """


        if str(object_id) == "None":
            object_id = self.collaborated_object

        logger.debug("Inside list_permissions to fetch permission for " + str(object_id))
        permissions_list = []
        permissions_dict = {6: "none", 1: "viewer", 2: "editor", 3: "guest", 5: "owner"}
        principaltype_dict = {0 :"none", 1 :"user", 2 :"distributionlist", 4 :"securitygroup", 8 :"sharepointgroup", 15 :"all"}
        headers = self.headers.copy()
        headers.update({"Accept": "application/json"})

        # headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}

        if object_id:
            if self.lastuploadedfiles and self.lastuploadedfiles[-1].get('filename') == object_id.split('/')[-1]:

                endpoint_GetByServerRelativeUrl = re.sub('GetFileByServerRelativeUrl(.*)',
                                                         'GetFileByServerRelativeUrl(\'' + object_id + '\')',
                                                         self.endpoint_GetFileByServerRelativeUrl)
            else:
                endpoint_GetByServerRelativeUrl = re.sub('GetFolderByServerRelativeUrl(.*)',
                                                         'GetFolderByServerRelativeUrl(\'' + object_id + '\')',
                                                         self.endpoint_GetFolderByServerRelativeUrl)
        else:
            if self.lastuploadedfiles and self.lastuploadedfiles[-1].get('filename') == object_id.split('/')[-1]:

                endpoint_GetByServerRelativeUrl = self.endpoint_GetFileByServerRelativeUrl
            else:
                endpoint_GetByServerRelativeUrl = self.endpoint_GetFolderByServerRelativeUrl

        self.get_permissions_url = endpoint_GetByServerRelativeUrl + \
                                   "?$expand=ListItemAllFields/ParentList,ListItemAllFields/RoleAssignments/Member," + \
                                   "ListItemAllFields/RoleAssignments/RoleDefinitionBindings," + \
                                   "ListItemAllFields/RoleAssignments/Member/Users"

        response_get_permissions = requests.get(self.get_permissions_url, headers=headers, cookies=self.cookies)

        if response_get_permissions.status_code == 200:
            #logger.info("Response for getting permission is " + str(response_get_permissions.text))
            logger.info("Response for getting permission is successful")

        elif response_get_permissions.status_code in [401,403]:
            logger.debug("Got a 403 error..")
            logger.debug("Response for getting permission is " + str(response_get_permissions.text))
            self._refresh_token()

        # logger.debug(json.dumps(response_get_permissions.json().get("ListItemAllFields").get("RoleAssignments")))
        for item in response_get_permissions.json()["ListItemAllFields"]["RoleAssignments"]:
            if "Email" in item["Member"] and not item["Member"]["IsSiteAdmin"]:
                #logger.debug("item_Member_email is: " + item["Member"]["Email"])
                #logger.debug("self.user is:  " + self.user)
                if item["Member"]["Email"] and item["Member"]["Email"] not in self.user:
                    member = {"email":"", "role": "", "type": ""}
                    member["email"] = str(item["Member"]["Email"])
                    # permissions_list.append({"email": "", "role": ""})
                    logger.trace("User found with EmailID: " + str(item["Member"]["Email"]))
                    # permissions_list[-1]["email"] = str(item["Member"]["Email"])
                    role_type = []
                    for role_list in item["RoleDefinitionBindings"]:
                        logger.trace("Role associated with the user is:" + str(role_list["RoleTypeKind"]))
                        role_type.append(role_list["RoleTypeKind"])
                    logger.trace("Role type list contains " + str(role_type))
                    # permissions_list[-1]["role"] = permissions_dict[max(role_type)]
                    member["role"] = permissions_dict[max(role_type)]
                    if member["role"] is not "guest":
                        permissions_list.append(member)
                    member["type"] = principaltype_dict[int(item["Member"]["PrincipalType"])]
        logger.trace("Response of retrieving all collab is " + str(permissions_list))
        return permissions_list

    def list_permission(self, permission_id):
        """
        List permission given the permission ID
        :param permission_id: Permission Id of an object for which info is required
        :return: JSON representing the different attributes of a permission e.g. username, role
        """
        return NotImplementedError

#    @timeout_decorator.timeout(90, use_signals=False)
    def add_permission(self, object_id, user_attr, reset=True, file_collaboration=False,collaborator=None,add_to_all_collaborators=True):
        """
        Add a new permission given the object, and user attributes been provided
        Role details : 1 = View, 2 = Edit, 3 = Owner, 0 = None

        Args:
            object_id: Id of the object for which permissions required to be listed
            user_attr: List of user e.g. {"email":"skyhigh.blore@gmail.com", "role": 2}
                       If email is "default" , we get current user info and add +1 to the user name
                       e.g. skyhigh.blore@gmail.com will become skyhigh.blore+1@gmail.com
            reset: Optional, if true it will re apply the same permission if found

        Returns:
            True : If permission is added successfully

        Raises:
            None
        """

        if (self.user not in self.remaining_collaborators) and (self.__class__.__name__ == "SharePoint"):
            self.remaining_collaborators.append(self.user)
        try:
            modified_permision = BuiltIn().replace_variables('${permission_type}')
        except:
            modified_permision = ''

        user = None
        user_attr_dict = {"viewer": 1, "editor": 2, "owner": 3}
        result = False
        variables = BuiltIn().get_variables()
        if "${multiple_collaborators}" in variables:
            self.multiple_collaborators = BuiltIn().replace_variables('${multiple_collaborators}')
        if "${flexilink}" in variables:
            self.isFlexiLink = BuiltIn().replace_variables('${flexilink}')
        else:
            self.isFlexiLink = False

        if not object_id:
            if file_collaboration:
                if self.mostrecentfolder and self.mostrecentfolder!=0:
                    self.collaborated_object = self.mostrecentfolder + "/" + self.lastuploadedfiles[-1].get('filename')
                else:
                    self.collaborated_object = self.lastuploadedfiles[-1].get('fileid')
            else:
                self.collaborated_object = self.mostrecentfolder
        else:
            self.collaborated_object = object_id

        test_name = BuiltIn().replace_variables('${TEST_NAME}')
        BuiltIn().set_suite_metadata(test_name + "_" + str(self.instance_id) + "_collaborated_object",
                                     self.collaborated_object)

        logger.debug("Going to add permission for object: " + str(self.collaborated_object))
       # logger.debug("last created group onedrive" + str(self.lastcreatedgroup))
        user_attr = eval(str(user_attr))
        user_attr["role"] = user_attr_dict[user_attr["role"].lower()]
        logger.debug("user_attr contain " + str(user_attr))
        collaborators = list()

        if user_attr["email"] == "*":
            user = self._get_another_user_for_collab()
            user_attr["email"] = user
            logger.debug("User to Collaborate with: " + user)
        elif user_attr["email"] == "external collaborator":
             user = self.externalcollaborator
             user_attr["email"] = user
        elif user_attr["email"] == "group":
             if self.lastcreatedgroup[-1]["email"] != None:
                 user = self.lastcreatedgroup[-1]["email"]
             else:
                 user = self.lastcreatedgroup[-1]["groupname"]
             logger.debug("User to Colloborate with: " + user)
             if not self.isFlexiLink:
                 user_attr["role"] = "role:1073741830" if user_attr["role"] == 2 else "role:1073741826"
        elif user_attr["email"] == 'unaccepted invite collaborator':
            collaborators = self.unacceptedinvitecollaborator.split(",")
            user = collaborators[0]
        else:
            if isinstance(user_attr["email"].split('@'), list) and \
                            user_attr["email"].split('@')[1] != self.user.split('@')[1]:
                user = user_attr["email"]
                user_attr["email"] = 'unaccepted invite collaborator'
                collaborators = [user]
            else:
                user = user_attr["email"]

        #Next set of if else conditions are to invoke collaboration methods as per the collaboration types
        if user_attr["email"] == 'unaccepted invite collaborator' or (user_attr["email"] == user and self.isFlexiLink):
            '''
            role:1073741830(previously 1073741827) is for edit role and role:1073741826 is for view
            '''
            if (user_attr["email"] == user and self.isFlexiLink):
                collaborators = [user]
            collaboration_details = {
                "collaborators": collaborators,
                "add_to_all_collaborators": add_to_all_collaborators,
                "modified_permision": modified_permision,
            }
            result = self.flexiLinkUserCollaboration(user_attr, collaboration_details)
        elif user_attr["email"] == "group" and self.isFlexiLink:
            user_attr["role"] = "2"
            if add_to_all_collaborators:
                self.all_collaborators.append(self.lastcreatedgroup[-1]["groupname"])
            if add_to_all_collaborators and modified_permision.lower() == 'viewer' or add_to_all_collaborators is False:
                self.remaining_collaborators.append(self.lastcreatedgroup[-1]["groupname"])
            result = self.flexiLinkGroupCollaboration(user_attr,file_collaboration)
        elif user_attr["email"] == "group":
            collaboration_details = {
                "collaborator": collaborator,
                "add_to_all_collaborators": add_to_all_collaborators,
                "modified_permision": modified_permision,
                "file_collaboration": file_collaboration
            }
            result = self.directAccessGroupCollaboration(collaboration_details, user_attr)
        else:
            if add_to_all_collaborators:
                self.all_collaborators.append(user_attr["email"])
            if add_to_all_collaborators and modified_permision.lower() == 'viewer' or add_to_all_collaborators is False:
                self.remaining_collaborators.append(user_attr["email"])
            result = self.directAccessUserCollaboration(user_attr, file_collaboration)

        BuiltIn().set_suite_metadata(test_name + "_" + str(self.instance_id) + "_permission_object",
                                     self.permission_object)
        return result

    def get_listid_itemId(self,item,file=None):
        """
                get info listId and itemId of the file/folder.
                :param file/folder
                :return: tuple
        """
        retry  = 3
        logger.debug("Inside get_listid_itemId")
        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}

        endpoint_GetListByServerRelativeUrl  = self.endpoint_GetFileListByServerRelativePathUrl if file \
            else self.endpoint_GetFolderListByServerRelativePathUrl


        endpoint_GetListByServerRelativeUrl = endpoint_GetListByServerRelativeUrl % item

        while retry > 0:
            response = requests.get(url=endpoint_GetListByServerRelativeUrl, headers=headers, cookies=self.cookies)
            if response.status_code == 200:
                logger.debug("Response for ListAllItems: " + str(response.text))
                m = re.search("guid'(.*)'\)/Items\((.*)\)", response.json()["odata.editLink"])
                id_tuple = m.groups()
                break
            elif response.status_code in [401, 403]:
                logger.debug("Got a 403 error..")
                self._refresh_token()
                retry  = retry - 1

        return id_tuple

    def directAccessGroupCollaboration(self, collaboration_details, user_attr):
        newApi = None
        try:
            newApi = BuiltIn().replace_variables('${listApi}')
        except:
            pass

        if newApi is not None and newApi:
            listId, itemId = self.get_listid_itemId(self.collaborated_object, collaboration_details["file_collaboration"])
            endpoint_sharing = self.endpoint_DirectAccessSharing_listId % (listId,itemId)
        else:
            endpoint_sharing = self.endpoint_File_DirectAccessSharing if collaboration_details["file_collaboration"] \
                                                                      else self.endpoint_Folder_DirectAccessSharing
            endpoint_sharing = endpoint_sharing % self.collaborated_object
        logger.debug("URL for sharing is " + endpoint_sharing)

        # Create the list based on a specific collaborator or all
        if collaboration_details["collaborator"]:
            for group in self.lastcreatedgroup:
                if group["groupname"] == collaboration_details["collaborator"]:
                    if collaboration_details["add_to_all_collaborators"]:
                        self.all_collaborators.append(group["groupname"])
                    if collaboration_details["add_to_all_collaborators"] and collaboration_details["modified_permision"].lower() == 'viewer' \
                            or collaboration_details["add_to_all_collaborators"] == False:
                        self.remaining_collaborators.append(group["groupname"])
                    break
        else:
            if collaboration_details["add_to_all_collaborators"]:
                self.all_collaborators.append(group["groupname"] for group in self.lastcreatedgroup)
            if collaboration_details["add_to_all_collaborators"] and collaboration_details["modified_permision"].lower() == 'viewer' \
                    or collaboration_details["add_to_all_collaborators"] == False:
                self.remaining_collaborators.append(group["groupname"] for group in self.lastcreatedgroup)

        picker_input = list()
        collaborated_group_emails = list()
        group_attributes = list()
        group_attibute = {'role': "", 'email': ""}

        # Create the permission request payload based on a specific collaborator or all
        for group in self.lastcreatedgroup:
            if collaboration_details["collaborator"] and collaboration_details["collaborator"] == group["email"]:
                if newApi is not  None and newApi:
                    picker_key = "{\"Key\": \"" + group["fid"] + "\"," + \
                                 "\"DisplayText\": \"" + group['apiDisplayText'] + "\"," + \
                                 "\"IsResolved\": true," + \
                                 "\"Description\":\"" + group["groupname"] + "\"," + \
                                 "\"EntityType\":\"SecGroup\"," + \
                                 "\"EntityData\":{\"Email\":\"" + group["email"] + "\"," + \
                                 "\"DisplayName\":\"" + group["apiDisplayName"] + "\"}," + \
                                 "\"MultipleMatches\":[]," + \
                                 "\"ProviderName\":\"FederatedDirectoryClaimProvider\"," + \
                                 "\"ProviderDisplayName\":\"Federated Directory\"}"
                else:
                    picker_key = "{\"Key\": \"" + group["email"] + "\"}"
                picker_input.append(picker_key)
                collaborated_group_emails.append(group["email"])
                group_attibute['role'] = user_attr['role']
                group_attibute['email'] = group['email']
                group_attributes.append(group_attibute)
                break
            else:
                if newApi is not  None and newApi:
                    picker_key = "{\"Key\": \"" + group["fid"] + "\"," + \
                                   "\"DisplayText\": \"" + group['apiDisplayText'] + "\"," + \
                                   "\"IsResolved\": true," + \
                                   "\"Description\":\"" + group["groupname"] + "\"," + \
                                   "\"EntityType\":\"SecGroup\"," + \
                                   "\"EntityData\":{\"Email\":\"" + group["email"] + "\"," + \
                                                    "\"DisplayName\":\"" + group["apiDisplayName"] + "\"}," + \
                                   "\"MultipleMatches\":[]," + \
                                   "\"ProviderName\":\"FederatedDirectoryClaimProvider\","+ \
                                   "\"ProviderDisplayName\":\"Federated Directory\"}"

                else:
                    picker_key = "{\"Key\": \"" + group["email"] + "\"}"
                picker_input.append(picker_key)
                collaborated_group_emails.append(group["email"])
                group_attibute['role'] = user_attr['role']
                group_attibute['email'] = group['email']
                group_attributes.append(group_attibute)

        peoplePicker = '[' + ', '.join(picker_input) + ']'
        data = {
            "emailBody": None,
            "includeAnonymousLinkInEmail": False,
            "sendEmail": True,
            "propagateAcl": True,
            "useSimplifiedRoles": True,
            "roleValue": user_attr["role"],
            "peoplePickerInput": peoplePicker
        }
        logger.debug("Users to collaborate with :" + str(self.all_collaborators))
        logger.debug("Data is " + str(data))
        result = list()
        for n in range(1, 4):
            logger.info("Trying to add permission to the file... iteration %s of 3" % n)
            headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json",
                       "Content-Type": "application/json"}
            response_add_permission = requests.post(url=endpoint_sharing, headers=headers,
                                                    data=json.dumps(data), cookies=self.cookies,
                                                    timeout=self.stuck_timeout)
            logger.debug("The response action is ",response_add_permission)
            if response_add_permission.status_code == 200:
                logger.debug("Response of setting permission for folder is " + str(response_add_permission.text))
                invitedUsers = response_add_permission.json()["UniquelyPermissionedUsers"]
                for invitedUser in invitedUsers:
                    logger.debug(invitedUser["Email"])
                    if invitedUser["Email"] in collaborated_group_emails or invitedUser['User'] in collaborated_group_emails :
                        result.append(True)
                    else:
                        result.append(False)
                break

            elif response_add_permission.status_code in [401, 403, 504]:
                logger.debug("Got not success code: %s so retrying" % response_add_permission.status_code)
                logger.debug("Response for getting permission is " + str(response_add_permission.text))
                self._refresh_token()

        if self.lastuploadedfiles:
            if "object" not in self.lastuploadedfiles[-1]:
                self.lastuploadedfiles[-1]["permissions_object"]["id"] = self.collaborated_object
                self.lastuploadedfiles[-1]["permissions_object"]["permissions_list"] = [group_attributes]
            else:
                self.lastuploadedfiles[-1]["permissions_object"]["permissions_list"] = [group_attributes]
        else:
            self.permission_object = {"permissions_object": {"id": self.collaborated_object, "permissions_list": [group_attributes]}}
            logger.debug("permission object is %s" %(self.permission_object))

        return False if len(result) == 0 else all(result)

    def flexiLinkGroupCollaboration(self, user_attr,file_collaboration=None):
        newApi = None
        try:
            newApi = BuiltIn().replace_variables('${listApi}')
        except:
            pass

        if newApi is not  None and newApi:
            listId, itemId = self.get_listid_itemId(self.collaborated_object, file_collaboration)
            endpoint_sharing = self.endpoint_Flexilink_bylistid % (listId,itemId)
        else:
            endpoint_sharing = self.endpoint_Flexilink.format(self.collaborated_object)

        #endpoint_sharing = self.endpoint_Flexilink.format(self.collaborated_object)
        logger.debug("URL for sharing is " + endpoint_sharing)
        picker_input = list()
        collaborated_group_emails = list()
        group_attributes = list()
        group_attibute = {'role': "", 'email': ""}
        if newApi is not None and newApi:
            for group in self.lastcreatedgroup:
                picker_key = "{\"Key\": \"" + group["fid"] + "\"," + \
                             "\"DisplayText\": \"" + group['apiDisplayText'] + "\"," + \
                             "\"IsResolved\": true," + \
                             "\"Description\":\"" + group["groupname"] + "\"," + \
                             "\"EntityType\":\"SecGroup\"," + \
                             "\"EntityData\":{\"Email\":\"" + group["email"] + "\"," + \
                             "\"DisplayName\":\"" + group["apiDisplayName"] + "\"}," + \
                             "\"MultipleMatches\":[]," + \
                             "\"ProviderName\":\"FederatedDirectoryClaimProvider\"," + \
                             "\"ProviderDisplayName\":\"Federated Directory\"}"
                picker_input.append(picker_key)
                collaborated_group_emails.append(group["email"])
                group_attibute['role'] = user_attr['role']
                group_attibute['email'] = group['email']
                group_attributes.append(group_attibute)

            peoplePicker = '[' + ', '.join(picker_input) + ']'
            data = {
                "request": {
                    "createLink": True,
                    "settings": {
                        "linkKind": 6,
                        "expiration": None,
                        "role": user_attr['role'],
                        "restrictShareMembership": True,
                        "updatePassword": False,
                        "password": ""
                    },
                    "peoplePickerInput": peoplePicker,
                }
            }
        else:
            for group in self.lastcreatedgroup:
                picker_key = "{\"Key\": \"" + group["email"] + "\"}"
                picker_input.append(picker_key)
                collaborated_group_emails.append(group["email"])
                group_attibute['role'] = user_attr['role']
                group_attibute['email'] = group['email']
                group_attributes.append(group_attibute)

            peoplePicker = '[' + ', '.join(picker_input) + ']'
            data = {
                "request": {
                    "createLink": True,
                    "settings": {
                        "linkKind": 6,
                        "expiration": None,
                        "role": user_attr['role'],
                        "restrictShareMembership": True,
                        "updatePassword": False,
                        "password": ""
                    },
                    "peoplePickerInput": peoplePicker,
                    "emailData": {
                        "body": "",
                        "subject": ""
                    }
                }
            }
        logger.debug("Users to collaborate with :" + str(self.all_collaborators))
        logger.debug("Data is " + str(data))

        result = list()
        for n in range(1, 4):
            logger.info("Trying to add permission to the file... iteration %s of 3" % n)
            headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json",
                       "Content-Type": "application/json"}
            response_add_permission = requests.post(url=endpoint_sharing, headers=headers, data=json.dumps(data),
                                                    cookies=self.cookies, timeout=self.stuck_timeout)
            if response_add_permission.status_code == 200:
                logger.debug("Response of setting permission for resource is " + str(response_add_permission.text))
                invitations = response_add_permission.json()['sharingLinkInfo']['Invitations']
                for invitation in invitations:
                    if invitation['invitee']['email'] in collaborated_group_emails:
                        result.append(True)
                    else:
                        result.append(False)
                break
            elif response_add_permission.status_code in [401, 403, 504]:
                logger.debug("Got not success code: %s so retrying" % response_add_permission.status_code)
                logger.debug("Response for getting permission is " + str(response_add_permission.text))
                self._refresh_token()

        if self.lastuploadedfiles:
            if "object" not in self.lastuploadedfiles[-1]:
                self.lastuploadedfiles[-1]["permissions_object"]["id"] = self.collaborated_object
                self.lastuploadedfiles[-1]["permissions_object"]["permissions_list"] = [group_attributes]
            else:
                self.lastuploadedfiles[-1]["permissions_object"]["permissions_list"] = [group_attributes]
        else:
            self.permission_object = {"permissions_object": {"id": self.collaborated_object, "permissions_list": [group_attributes]}}
            logger.debug("permission object is %s" %(self.permission_object))

        return False if len(result) == 0 else all(result)

    def flexiLinkUserCollaboration(self, user_attr, collaboration_details):

        endpoint_sharing = self.endpoint_Flexilink.format(self.collaborated_object)
        logger.debug("URL for sharing is " + endpoint_sharing)
        role = "role:1073741830" if user_attr["role"] == 2 else "role:1073741826"
        collaborators = collaboration_details['collaborators']
        if self.multiple_collaborators:
            picker_input = list()
            for i in range(len(collaborators)):
                picker_key = "{\"Key\": \"" + collaborators[i] + "\"}"
                picker_input.append(picker_key)
            peoplePicker = '[' + ', '.join(picker_input) + ']'
            self.all_collaborators.extend(collaborators)
            user=collaborators
        else:
            user = collaborators[0]
            peoplePicker = "[{\"Key\": \"" + user + "\"}]"
            if collaboration_details['add_to_all_collaborators']:
                self.all_collaborators.append(user)

            if collaboration_details['add_to_all_collaborators'] and collaboration_details['modified_permision'].lower() == 'viewer' \
                    or collaboration_details['add_to_all_collaborators'] == False:
                self.remaining_collaborators.append(user)
        data = {
            "request": {
                "createLink": True,
                "settings": {
                    "linkKind": 6,
                    "expiration": None,
                    "role": user_attr['role'],
                    "restrictShareMembership": True,
                    "updatePassword": False,
                    "password": ""
                },
                "peoplePickerInput": peoplePicker,
                "emailData": {
                    "body": "",
                    "subject": ""
                }
            }
        }
        logger.debug("Users to collaborate with :" + str(self.all_collaborators))
        logger.debug("Data is " + str(data))
        user_attr["email"] = user
        result = list()
        for n in range(1, 4):
            logger.info("Trying to add permission to the file... iteration %s of 3" % n)
            headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json",
                       "Content-Type": "application/json"}
            response_add_permission = requests.post(url=endpoint_sharing, headers=headers, data=json.dumps(data),
                                                    cookies=self.cookies, timeout=self.stuck_timeout)
            if response_add_permission.status_code == 200:
                logger.debug("Response of setting permission for resource is " + str(response_add_permission.text))
                invitations = response_add_permission.json()['sharingLinkInfo']['Invitations']
                for invitation in invitations:
                    if invitation['invitee']['email'] in collaborators:
                        result.append(True)
                    else:
                        result.append(False)
                break
            elif response_add_permission.status_code in [401, 403, 504]:
                logger.debug("Got not success code: %s so retrying" % response_add_permission.status_code)
                logger.debug("Response for getting permission is " + str(response_add_permission.text))
                self._refresh_token()

        if self.lastuploadedfiles:
            if "object" not in self.lastuploadedfiles[-1]:
                self.lastuploadedfiles[-1]["permissions_object"]["id"] = self.collaborated_object
                self.lastuploadedfiles[-1]["permissions_object"]["permissions_list"] = [user_attr]
            else:
                self.lastuploadedfiles[-1]["permissions_object"]["permissions_list"] = [user_attr]
        else:
            self.permission_object = {"permissions_object": {"id": self.collaborated_object, "permissions_list": [user_attr]}}
            logger.debug("permission object is %s" %(self.permission_object))

        return False if len(result) == 0 else all(result)

    @keyword("In ${SERVICE} add specific members ${users_ids} to ${group_id}")
    def add_members(self, usersIds, group_id):
        ids_to_add_to_group = []
        ids_to_add_to_group.extend(usersIds)
        logger.debug(ids_to_add_to_group)
        headers = {"Authorization": "Bearer " + self.access_token_graph, "Content-type": "application/json"}
        logger.debug(group_id)
        endpoint_group_member_url = self.endpoint_groups_url + str(group_id)
        members_list = []
        for i in range(len(ids_to_add_to_group)):
            members_list.append("https://graph.microsoft.com/v1.0/directoryObjects/" + str(ids_to_add_to_group[i]))
        data = {"members@odata.bind": members_list}
        try:
            response_add_members = requests.patch(url=endpoint_group_member_url, headers=headers,
                                                  data=json.dumps(data))
            logger.debug(response_add_members.content)
            return True
        except Exception as e:
            logger.error(" ======member not added====== " + str(e))
            raise Exception
            return False

    @keyword("In ${SERVICE} get ${domain_name} domain ${count} users")
    def get_same_domain_external_users(self, domain, number):
        self.same_domain_external_users = []
        headers = {"Authorization": "Bearer " + self.access_token_graph, "Accept": "application/json"}
        self.response_get_all_users = requests.get(url=self.endpoint_users, headers=headers)
        if self.response_get_all_users.status_code == 200:
            for user in self.response_get_all_users.json().get("value"):
                logger.debug(user)
                # if "ext" in user.get("userPrincipalName").lower():
                if len(self.same_domain_external_users) < int(number):
                    if domain in user.get("userPrincipalName").lower():
                        self.same_domain_external_users.append(user.get("mail"))
                    else:
                        continue
        member_ids_to_add = self._get_user_ids(self.same_domain_external_users)
        return member_ids_to_add

    @keyword("In ${SERVICE} get o365 group ${value}")
    def get_members_from_group(self, value):
        logger.debug(self.lastcreatedgroup)
        return self.lastcreatedgroup[0][value]

    def directAccessUserCollaboration(self, user_attr, file_collaboration):

        endpoint_sharing = self.endpoint_File_DirectAccessSharing if file_collaboration \
                                                        else self.endpoint_Folder_DirectAccessSharing
        endpoint_sharing = endpoint_sharing % self.collaborated_object
        logger.debug("URL for sharing is " + endpoint_sharing)

        user_attr["role"] = "role:1073741827" if user_attr["role"] == 2 else "role:1073741826"
        peoplePicker = '[{"Key": "' + user_attr["email"] + '"}]'
        data = {
            "emailBody": None,
            "includeAnonymousLinkInEmail": False,
            "sendEmail": True,
            "propagateAcl": True,
            "useSimplifiedRoles": True,
            "roleValue": user_attr["role"],
            "peoplePickerInput": peoplePicker
        }
        logger.debug("Users to collaborate with :" + user_attr["email"])
        logger.debug("Data is " + str(data))

        result = list()
        for n in range(1, 4):
            logger.info("Trying to add permission to the file... iteration %s of 3" % n)
            headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json",
                       "Content-Type": "application/json"}
            response_add_permission = requests.post(url=endpoint_sharing, headers=headers,
                                                    data=json.dumps(data), cookies=self.cookies,
                                                    timeout=self.stuck_timeout)
            if response_add_permission.status_code == 200:
                logger.debug("Response of setting permission for folder is " + str(response_add_permission.text))
                invitedUsers = response_add_permission.json()["UniquelyPermissionedUsers"]
                for invitedUser in invitedUsers:
                    if invitedUser["Email"] in [user_attr["email"]]:
                        result.append(True)
                    else:
                        result.append(False)
                break

            elif response_add_permission.status_code in [401, 403, 504]:
                logger.debug("Got not success code: %s so retrying" % response_add_permission.status_code)
                logger.debug("Response for getting permission is " + str(response_add_permission.text))
                self._refresh_token()

        if self.lastuploadedfiles:
            if "object" not in self.lastuploadedfiles[-1]:
                self.lastuploadedfiles[-1]["permissions_object"]["id"] = self.collaborated_object
                self.lastuploadedfiles[-1]["permissions_object"]["permissions_list"] = [user_attr]
            else:
                self.lastuploadedfiles[-1]["permissions_object"]["permissions_list"] = [user_attr]
        else:
            self.permission_object = {"permissions_object": {"id": self.collaborated_object, "permissions_list": [user_attr]}}
            logger.debug("permission object is %s" %(self.permission_object))

        return False if len(result) == 0 else all(result)


    @keyword("Compare permissions in ${SERVICE} ${before} and ${after}")
    def compare_permissions(self,permissions_before,permissions_after):
        logger.debug ("permissions before " + str(permissions_before))
        logger.debug("permissions after " + str(permissions_after))
        from skybot.lib.ObjectPathHelper import object_path_helper
        collaborators_before = object_path_helper.get_object_path_value(permissions_before, "$..*.email")
        collaborators_after = object_path_helper.get_object_path_value(permissions_after, "$..*.email")
        logger.debug(collaborators_after,collaborators_before)
        if collaborators_before == collaborators_after:
            logger.debug("Collaborators have been restored")
            return True
        else:
            logger.error("Collaborators not restored successfully")
            return False

    def edit_permission(self, permission_id, role):
        """
        Edit the permission by recognizing it with permission ID . Can only update role
        :param permission_id: permission ID to edit
        :param role: role of the email to change to . Available options are :-
                    # editor, viewer, previewer, uploader, previewer uploader, viewer uploader, co-owner, or owner
        :return: Boolean
        """
        return NotImplementedError

    def delete_permission(self, permission_id):
        """
        Delete permission using Permission Id

        Args:
            permission_id: Permission Id

        Returns:
            Boolean

        Raises:
            None
        """
        return NotImplementedError

    # Link methods
    @keyword("generate link in ${SERVICE} for last uploaded ${object}")
    @timeout_decorator.timeout(90, use_signals=False)
    def create_link(self, object=object):
        """
        Create the link for an object. yet to be implemented

        Args:
            object: File or Folder
            object_id: Object id for which link to be created
            password (Optional): password to be set
            expiration (Optional): expiration to be set
            direct ( Optional): Boolean specifies it is a direct link or not
            link_type ( Optional): edit or view  # specific to sharepoint or onedrive
        Returns:
            Link Id or link that gets generated

        Raises:
            None
        """
        return self._create_link(link_type='external',object=object,role='editor',allowFileDiscovery=None)

    @keyword("For ${service} generate ${link_type} link for ${object} having ${role} with ${allowFileDiscovery}")
    def internally_shared_link_file(self,link_type,object,role,allowFileDiscovery):
        logger.debug("1. Creating a %s shared link for last uploaded object %s with role %s" % (link_type,object,role))
        return self._create_link(link_type,object,role,allowFileDiscovery)

    def _create_link(self,link_type,object,role,allowFileDiscovery, object_id=None):

        link_url = None
        link_endpoint=''
        data = {}
        if object_id is None:
            if object == "file":
                object_id = self.lastuploadedfiles[-1]["fileid"]
            elif object == "folder":
                object_id = self.mostrecentfolder
        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
        headers.update({"Content-Type": "application/json"})
        if link_type=='internal':
            link_endpoint = 'SP.Web.CreateOrganizationSharingLink'
        elif link_type=='external':
            link_endpoint= 'SP.Web.CreateAnonymousLink'
        endpoint_create_link = self.endpoint_create_link + link_endpoint
        if role == 'editor' or role == 'any':
            data = {"url": self.domain_url + object_id, "isEditLink": True}
        elif role == 'viewer':
            data = {"url": self.domain_url + object_id, "isEditLink": False}
        response_create_link = requests.post(url=endpoint_create_link, headers=headers, data=json.dumps(data), cookies=self.cookies, timeout=self.stuck_timeout)
        if isinstance(response_create_link, dict):
            if response_create_link.get('status_code') == 504:
                logger.warn('request timed out for {0} after {1} sec, however we will return True '
                            'assuming, post request is successful!!'.format("create_link", self.stuck_timeout))
            return True
        else:
            logger.debug("The response is %s" % (response_create_link.json()))
            if response_create_link.status_code == 200:
                logger.debug("Response from generating Link is " + response_create_link.text)

            elif response_create_link.status_code in [401,403]:
                logger.debug("Got a 403 error..")
                self._refresh_token()
                raise Exception
            elif (re.search("Blocked by policy response",str(response_create_link.json()["odata.error"]["message"]["value"]))):
                logger.console("Permission is not set due to block permissions rule")
                result=True

            elif (re.search("Please review the file for sensitive or confidential content",str(response_create_link.json()["odata.error"]["message"]["value"]))):
                logger.console("Permission is not set due to block permissions rule")
                result=True

            if "value" in response_create_link.json():
                link_url = response_create_link.json()["value"]
                logger.debug("Link is generated with the link_url: " + str(link_url))

        return link_url

    def update_link(self, link_id, active, expiration, password):
        """
        Update the link with link id. yet to be implemented

        Args:
            link_id: link to be updated
            active (Optional): Boolean
            expiration (Optional): expiration to be set
            password (Optional): password to be set

        Returns:
            Link Id or link that gets generated

        Raises:
            None
        """
        return NotImplementedError

    def retrieve_link(self, link_id, object_id):
        """
        Retrieve the link by link id. yet to be implemented

        Args:
            link_id: link id of link to be retrieved

        Returns:
            Link URl

        Raises:
            None
        """
        return NotImplementedError

    def delete_link(self, link_id):
        """
        Delete the link by link id. yet to be implemented

        Args:
            link_id: link id to be deleted

        Returns:
            Boolean

        Raises:
            None
        """
        return NotImplementedError

    def retry_handler(func):
        def wrapper(instance, *args, **kwargs):
            result = None
            try:
                result = func(instance, *args, **kwargs)
                if result == False:
                    instance._refresh_token()
                    result = func(instance, *args, **kwargs)
            except:
                instance._refresh_token()
                result = func(instance, *args, **kwargs)

            return result

        return wrapper

    def get_all_links(self, object_id, type="file", link=None):
        """
        Get all the links for an object
        :param object_id:
        :param type: file or folder
        :return: link details
        LinkKind - 1 - restricted link 2 - view link with sign in 3
            edit link with sign in 4 - view link with no sign in 5 - edit link with no sign in
        """
        linkKind_dict = {1: "restricted_link", 2: "view_link_with_sign_in", 3: "edit_link_with_sign_in",
                            4: "view_link_with_no_sign_in", 5: "edit_link_with_no_sign_in"}
        logger.debug("Going to retrieve all links associated with object: " + str(object_id))
        # headers = self.headers.copy()
        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}
        headers.update({"Content-Type": "application/json"})
        response_object_json = None
        if type == "file":
            response_object_json = self.get_file_info(object_id, list_item_all_fields="/ListItemAllFields")
        elif type == "folder" or type == "nested folder":
            response_object_json = self.get_folder_info(object_id, params="/ListItemAllFields")
        editlink = response_object_json["odata.editLink"]
        editlink_re = re.match(r'Web/Lists\(guid\'(.*)\'\)/Items\((.*)\)', editlink)
        parent_guid = editlink_re.group(1)
        file_id = editlink_re.group(2)
        logger.debug("parent GUID is: " + parent_guid + "& file ID is: " + file_id)
        sharing_info_file_handler = open(WS+ "cfg/Office365/GetListItemSharingInfo.xml")
        sharing_info_file_content = str(sharing_info_file_handler.read())
        sharing_info_file_content = sharing_info_file_content.replace("{Guid}","{"+parent_guid+"}")
        sharing_info_file_content = sharing_info_file_content.replace("{Int32}",file_id)

        if link == 'flexilink':
            url = self.domain_url + object_id
            retrieve_flexilink_url = eval('self.endpoint_retrieve_flexilink % url')
            response_get_all_links = requests.post(retrieve_flexilink_url, headers=headers, cookies=self.cookies)
        else:
            response_get_all_links = requests.post(url=self.endpoint_retrieve_links, data=str(sharing_info_file_content), headers=headers, cookies=self.cookies)

        if response_get_all_links.status_code == 200:
            logger.info("Response for retrieve links is " + str(response_get_all_links.text))

        elif response_get_all_links.status_code in [401, 403]:
            logger.debug("Got a 403 error..")
            self._refresh_token()
        else:
            logger.debug("Failed to get response of get all links: " + str(response_get_all_links.json()))
            raise Exception

        logger.debug("Response of get all links is: " + str(response_get_all_links.json()))
        return response_get_all_links.json()

    @keyword("In ${SERVICE} get ${domain_name} domain ${count} users")
    def get_same_domain_external_users(self, domain, number):
        self.same_domain_external_users = []
        headers = {"Authorization": "Bearer " + self.access_token_graph, "Accept": "application/json"}
        self.response_get_all_users = requests.get(url=self.endpoint_users, headers=headers)
        if self.response_get_all_users.status_code == 200:
            for user in self.response_get_all_users.json().get("value"):
                logger.debug(user)
                #if "ext" in user.get("userPrincipalName").lower():
                if len(self.same_domain_external_users) < int(number):
                    if domain in user.get("userPrincipalName").lower():
                        self.same_domain_external_users.append(user.get("mail"))
                    else:
                        continue
        member_ids_to_add = self._get_user_ids(self.same_domain_external_users)
        return member_ids_to_add

    # User related methods
    @retry_handler
    @exclude_external_users
    def get_all_users(self):
        """
        Get list of current users info whose access token is fetched from Skyhigh console

        Args:
            None

        Returns:
            List of users

        Raises:
            None
        """
        logger.debug("Inside get all users")
        users = []
        self.external_users = []
        headers = {"Authorization": "Bearer " + self.access_token_graph, "Accept": "application/json"}
        self.response_get_all_users = requests.get(url=self.endpoint_users, headers=headers)
        if self.response_get_all_users.status_code == 200:
            for user in self.response_get_all_users.json().get("value"):
                if user is None or user == 'None':
                    logger.debug("User is none.Not appending in the list of user ")
                elif "ext" in user.get("userPrincipalName").lower():
                    # user is external
                    self.external_users.append(user.get("mail"))
                    users.append(str(user.get("mail")))
                else:
                    # user is internal
                    users.append(str(user.get("mail")))
        elif self.response_get_all_users.status_code in [403,401]:
                logger.debug("Got a %s error.." %(self.response_get_all_users.status_code))
                self._refresh_token()
                raise Exception
        else:
            logger.debug("Error in getting members info " + self.response_get_all_users.text)
        return users

    # Miscelleneous methods
    @keyword("check if file is tombstoned in")
    @trackme
    def isFileTombstoned(self, filename=None, template=None, HTMLWrap=False):
        """
        Check whether file is tombstoned or not by specifying filename. ( Current logic required to be enhanced )
        we cannot use file id because DLP updates the tombstoned file and we don't have a reference of file_id
        To verify if file is tombstoned, we are checking AppEditor field. Currently, we are not able to distinguish
        if file is Delete tombstoned or Quarantine tombstoned. yet to implement
        Args:
            filename: original filename for which required to be checked if tombstone file is generated ?
            template: check whether the placeholder file has template of delete or qurantine
        Returns:
            Boolean
        Raises:
            None
        """
        logger.debug("Inside OneDrive.isFileTombstoned: " + str(self.lastuploadedfiles))
        if template.lower() not in ['drm_protected', 'seclore_drm_protected']:
            ext = ".pdf"
        elif template.lower() == "seclore_drm_protected" and HTMLWrap:
            ext = ".html"
        else:
            ext = ""
        folder_id = self.mostrecentfolder
        file_to_check = None
        result = False
        if filename is None:
            file_to_check = self.lastuploadedfiles[-1]
        else:
            for each in self.lastuploadedfiles:
                id, name = each.get("fileid"), each.get("filename")
                if name == filename:
                    file_to_check = each

        if file_to_check:
            fileid, filename = file_to_check.get("fileid"), file_to_check.get("filename")

            logger.debug("Inside IsFileTombstoned, checking for " + filename +" " + str(template))


            if (fileid is None) or (filename is None) or (folder_id is None):
                logger.error("Found a None value", str(fileid)+str(filename)+str(folder_id))
                return False
        if not filename:
            logger.error("File is not known to be uploaded ", filename, self.lastuploadedfiles)
            return False
        for n in range(1, 20):
            logger.info("Trying to find the tombstone file")
            file_id = self.find_file_inside_folder_by_name(folder_id, str(filename) + str(ext))
            # if file_id is None:
            #     logger.debug("Error in getting the folder contents")
            if file_id:
                logger.info("Found the file, id is: " + str(file_id))
                file_info = self.get_file_info(file_id, list_item_all_fields="/ListItemAllFields/FieldValuesAsText?$select=AppEditor")
                logger.debug("App editor " + str(file_info))

                if template.lower() in ["quarantine", "delete"]:
                    app_editor_info = file_info.get("AppEditor")
                    if app_editor_info:
                        logger.debug("Verification is completed")
                        result = True
                        break
                elif template.lower() in ["drm_protected"]:
                    result = self._validate_drm_action(file_id, filename)
                    if result:
                        logger.debug("Found file as DRM Protected")
                        break
                    else:
                        logger.debug("File is not yet DRM protected...will re-verify after 10 secs")
                        time.sleep(10)
                elif template.lower() in ["seclore_drm_protected"]:
                    result = self.isFileSecloreDRMProtected(file_id, filename)
                    if result:
                        logger.debug("Found file as Seclore DRM Protected")
                        break
                    else:
                        logger.debug("File is not yet Seclore DRM protected...will re-verify after 10 secs")
                        time.sleep(10)

            else:
                logger.info("Did not find the file yet, sleeping for 10 sec - " + str(n) + "/20")
                time.sleep(15)
        if not result:
            logger.error("File "+str(filename)+" is not Tombstoned in O365")
            return False
        return result

    @keyword("verify changed permission for folder in ${Service} as role ${permission_type}")
    def verify_folder_permissions(self, permission_type, wait_time=60):
        """

        :param permission_type:
        :param wait_time:
        :return:
        """
        result = False
        # folder_id = self.mostrecentfolder
        logger.debug("Inside verify_folder_permissions for OneDrive")
        if self.lastuploadedfiles:
            if self.lastuploadedfiles[-1].get("permissions_object").get("permissions_list"):
                expected_permissions_list = self.lastuploadedfiles[-1].get("permissions_object").get("permissions_list")
            # logger.console("Inside Expected Permissions List:" +str(expected_permissions_list))
            else:
                expected_permissions_list = self.list_permissions(self.permission_object.get("permissions_object").get("id"))
        else:
            expected_permissions_list = self.list_permissions(self.permission_object.get("permissions_object").get("id"))

        logger.debug("Expected permission List is " + str(expected_permissions_list))
        for n in range(1, 20):
            current_permissions_list = self.list_permissions(self.collaborated_object)
            if current_permissions_list is None:
                break
            logger.debug("Current permission list is " + str(current_permissions_list))
            if not permission_type:
                logger.debug("Going to match None as Permission Type")
                if len(current_permissions_list)==1:
                    logger.debug("Going to find if only one user is present")
                    creds = self.shutil.services
                    if current_permissions_list[0]['role'] == 'owner' and (current_permissions_list[0]['email']).split('@')[1] == (creds['allservices']['SharePoint']['users']).split('@')[1]:
                        logger.debug("internal user only. hence revoke collaboration for external users is verified")
                        return True
                        break
                if not current_permissions_list:
                    logger.debug("Did not find any collaborated user, hence will return True")
                    result = True
                    break
            else:
                #logger.debug("Going to match for Permission Type: " + permission_type)
                for list in expected_permissions_list:
                    list["role"] = permission_type
                logger.debug("Expected permission list is:" + str(expected_permissions_list))
                if sorted(current_permissions_list, key=lambda i: i['email']) == \
                        sorted(expected_permissions_list, key=lambda i: i['email']):
                    logger.debug("Current and Expected permission list is matched!! , returning True")
                    result = True
                    break
            logger.debug("Verify permissions is not succeeded yet, sleeping for 10 sec - " + str(n) + "/20")
            time.sleep(10)
        if result is False:
            logger.error("Verifying Permission fails for object: " + str(self.collaborated_object))
        return result

    @keyword("Check ${Service} if ${type} ${link} expired")
    def check_link(self,type='', link=''):

        result=self.verify_link_expiry(type=type,link=link)
        return result

    @keyword("verify whether in ${Service} the ${type} link is expired or not")
    def verify_link_expiry(self, type='', link='', expired=True):
        link_kind_dict = {  'internal_view':2,
                            'internal_edit':3,
                            'external_view':4,
                            'external_edit':5,
                            'internal_any':3,
                            'external_any':5
                         }
        object_id = []

        logger.debug("Object_id passed is " + type)
        if str(type) == "file":
            logger.debug("Matched")
            logger.debug(self.lastuploadedfiles)
            object_id.append(self.lastuploadedfiles[-1]["fileid"])
        elif str(type) == "folder":
            object_id.append(self.mostrecentfolder)
        elif str(type) == "nested folder":
            object_id = self.nestedfolders
        for obj in object_id:
            logger.debug("Going to verify if link is expired on a object: " + str(obj))
            for n in range(1, 20):
                '''
                In case of collaboration with the external user who is not part of O365 users, O365 creates
                the collaboration as a shared link. So I have added the code to verify link expiry in this method
                instead of adding in verify_folder_permissions method
                '''
                if str(link) == 'flexilink':
                    all_links_json = self.get_all_links(obj, type, link)
                else:
                    all_links_json = self.get_all_links(obj, type)
                result=[]
                if str(link) in list(link_kind_dict.keys()):
                    link_kind_id=link_kind_dict.get(link)

                    for linktype in all_links_json[-1]["SharingLinks"]:
                        if linktype["LinkKind"] == link_kind_id:
                            if linktype["IsActive"] == False:
                                logger.debug("%s type link has expired" %(link))
                                result.append(True)
                            else:
                                result.append(False)
                                logger.debug("%s type link has not expired" %(link))
                else:
                    try:
                        for linktype in all_links_json[-1]["SharingLinks"]:
                            if linktype["LinkKind"] == 4 or linktype["LinkKind"] == 5:
                                logger.debug("Matched" + str(linktype))
                                if linktype["IsActive"] == False:
                                    result.append(True)
                                else:
                                    result.append(False)
                    except:
                        logger.debug("This is for flexilink sharing")
                        for linktype in all_links_json["value"]:
                            '''flexi link has linkkind 6 so checking for exact type'''
                            # if 4 <= linktype["linkDetails"]["LinkK ind"] <= 6:
                            if linktype["linkDetails"]["LinkKind"] == 6:
                                logger.debug("Matched" + str(linktype["linkDetails"]))
                                # if linktype["linkDetails"]["IsActive"] == False:
                                #     result.append(True)
                                # else:
                                #     result.append(False)
                                invitees = [invitation['invitee']['email'] for invitation in linktype["linkDetails"]["Invitations"]]
                                for user in self.all_collaborators:
                                    result.append(True) if user not in invitees else result.append(False)

                logger.debug("Result of verifying links is: " + str(result))
                if all(result) == True:
                    break
                else:
                    logger.debug("Verify links is not succeeded yet, sleeping for 10 sec - " + str(n) + "/20")
                    time.sleep(10)
        return all(result)

    def _get_domain_name(self):
        domain = re.match(r'.*@(.*)\.onmicrosoft\.com', self.user)
        if domain:
            self.domain_name = domain.group(1)
            logger.debug("Sharepoint domain name is: " + self.domain_name)
        elif self.access_token_graph is not None:
            logger.debug("Cannot get Sharepoint domain name from username: " + self.user)
            logger.info("Getting the Sharepoint domain name from MS Graph")
            header = {"Authorization": "Bearer " + self.access_token_graph}
            url = "https://graph.microsoft.com/v1.0/users/" + self.user + "/drive/root"
            try:
                r = requests.get(url, headers=header)
                Utils.request_trace(r)
                onedriveurl = r.json().get("webUrl")
                logger.debug("webUrl from Graph: " + str(onedriveurl))
                domain = re.match(r'https://(.+)-my\.sharepoint\.com', onedriveurl)
                self.domain_name = domain.group(1)
            except Exception as e:
                logger.error("Exception while extracting domain name from Graph response " + str(e))

        if self.domain_name is None:
            logger.error("Could not get Sharepoint domain name")
            return False

        logger.debug("Sharepoint and OneDrive domain name is " + self.domain_name)
        return True

    def _get_endpoints(self):
        """
        Going to build endpoints based on the email/domain
        :return: None
        """
        logger.debug("Building O365 endpoints...")

        self.user_flat = self.user.replace("@", ".").replace(".", "_")
        self.domain_url = "https://" + self.domain_name + "-my.sharepoint.com"
        self.domain_admin_url = "https://" + self.domain_name + "-admin.sharepoint.com"
        self.root_folder = "/personal/" + self.user_flat + "/" + self.list_library_name
        self.endpoint_GetFolderByServerRelativeUrl = self.domain_url + "/personal/" + self.user_flat + \
                                                        "/_api/Web/GetFolderByServerRelativeUrl(\'" + self.root_folder + "\')"
        self.endpoint_GetFileByServerRelativeUrl = self.domain_url + "/personal/" + self.user_flat + \
                                                    "/_api/Web/GetFileByServerRelativeUrl(\'" + self.root_folder + "\')"
        self.endpoint_GetFileListByServerRelativePathUrl = self.domain_url + "/personal/" + self.user_flat + \
                                                           "/_api/web/GetFileByServerRelativePath(decodedurl=@relativeUrl)" + \
                                              "/ListItemAllFields?@relativeUrl='%s'"
        self.endpoint_GetFolderListByServerRelativePathUrl = self.domain_url + "/personal/" + self.user_flat + \
                                                             "/_api/web/GetFolderByServerRelativePath(decodedurl=@relativeUrl)" + \
                                                      "/ListItemAllFields?@relativeUrl='%s'"
        self.endpoint_users = "https://graph.microsoft.com/v1.0/" + self.domain_name + ".onmicrosoft.com" + "/users"
        self.endpoint_retrieve_links = "https://" + self.domain_name + "-my.sharepoint.com" + "/personal/" + \
                                          self.user_flat + "/_vti_bin/client.svc/ProcessQuery"
        self.endpoint_contextinfo = self.domain_url + '/personal/' + self.user_flat + '/_api/contextinfo'
        self.endpoint_create_field = self.domain_url + '/personal/' + self.user_flat + \
                                      '/_api/web/lists/getbytitle(\'Documents\')/Fields'
        # self.endpoint_update_field = self.domain_url + '/personal/' + self.user_flat + \
        #                              '/_api/web/lists/getbytitle(\'Documents\')/items(2)'
        self.endpoint_groups = "https://graph.microsoft.com/v1.0/" + self.domain_name + ".onmicrosoft.com" + "/groups"
        self.endpoint_groups_url = "https://graph.microsoft.com/v1.0/groups/"
        self.endpoint_create_link = self.domain_url + "/personal/" + self.user_flat + "/_api/"
        self.endpoint_retrieve_flexilink = self.domain_url + "/personal/" + self.user_flat + \
                            "/_api/web/getlistitem(@url)/getsharinginformation/permissionsInformation/links?@url='%s'"
        self.endpoint_host_web_url = self.domain_url + "/personal/" + self.user_flat
        self.endpoint_list = self.domain_url + "/_api/SP.AppContextSite(@target)/web/Lists"
        self.endpoint_create_list = self.endpoint_list + "?@target='" + self.endpoint_host_web_url + "'"
        self.endpoint_GetFileByServerRelativePath = self.domain_url + '/personal/' + self.user_flat + \
                                    "/_api/web/GetFileByServerRelativePath(decodedurl=@relativeUrl)/$value?@relativeUrl=" + \
                                    "'/personal/" + self.user_flat
        self.default_root_folder = "/Documents/"
        self.endpoint_DirectAccessSharing = self.domain_url + "/personal/" + self.user_flat + \
                           "/_api/SP.Sharing.DocumentSharingManager.UpdateDocumentSharingInfo"
        self.endpoint_Flexilink = self.domain_url + '/personal/' + self.user_flat + \
                                  "/_api/web/GetListItemUsingPath(decodedurl=@u)/ShareLink?@u='{0}'"
        self.endpoint_Folder_DirectAccessSharing = self.domain_url + "/personal/" + self.user_flat + \
                    "/_api/web/GetFolderByServerRelativeUrl(@relativeUrl)/ListItemAllFields/ShareObject?@relativeUrl='%s'"
        self.endpoint_File_DirectAccessSharing = self.domain_url + "/personal/" + self.user_flat + \
                    "/_api/web/GetFileByServerRelativeUrl(@relativeUrl)/ListItemAllFields/ShareObject?@relativeUrl='%s'"
        self.endpoint_DirectAccessSharing_listId = self.domain_url +  "/personal/" + self.user_flat + "/_api/web/Lists(@a1)/GetItemById(@a2)/ShareObject?@a1='{%s}'&@a2='%s'"
        self.endpoint_Flexilink_bylistid = self.domain_url + '/personal/' + self.user_flat + \
                                  "/_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink?@a1='{%s}'&@a2='%s'"

    #@trackme
    def _get_request_digest(self):
        """
        This method is required to get request digest to be used by upload file method. Request digest is fetched by a series of steps
        a) Build and copy xml based on the prototype xml saved in cfg/GetOfficeTokenCommand.xml
        b) Get Office token and cookies
        :return: Tuple of Request Digest and Cookies
        """

        # this function requires the password for ofice 365, this was hardcoded before,
        # now read it from environment variable office365_password or the cfg file
        # and default to "Abcd_1234" when not found in cfg file

        if not self.shutil.current_service:
            servicename = "OneDrive"
        else:
            servicename=self.shutil.current_service

        #servicename = "onedrive"
        #default value in case of email verification
        for group,services in list(self.shutil.services.items()):
                if group == 'outageservices':
                    continue
                for service in list(services.keys()):
                    if self.instance_id == self.shutil.services.get(group).get(service).get("instanceid"):
                        servicename=service
                        break

        default_password = "Abcd_1234"
        #o365servicename = self.__class__.__name__
        o365servicename=servicename
        if os.environ.get("office365_password"):
            password = os.environ.get("office365_password")
            logger.debug("Took "+o365servicename+" password from environment variable office365_password")
        else:
            try:
                if o365servicename.lower() == "exchange":
                    password = self.shutil.services.get("mailservices").get(o365servicename).get("password", default_password)
                else:
                    password = self.shutil.services.get("allservices").get(o365servicename).get("password", default_password)
            except Exception as e:
                logger.warn("could not get "+o365servicename+" password from config file, using hardcoded default")
                password = default_password
        self.password = password

        tree = ET.parse(WS+"cfg/Office365/GetOfficeTokenCommand.xml")
        office_token = None
        cookies = None
        request_digest = None
        for value in tree.find('.//{http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd}UsernameToken'):
            if(value.text) == "OneDriveUserName":
                value.text = self.user
            if(value.text) == "OneDrivePassword":
                value.text = password
        for value in tree.find('.//{http://www.w3.org/2005/08/addressing}EndpointReference'):
             if(value.text) == "OneDriveBaseUrl":
                 value.text = self.domain_url
        tree.write(WS+"tmp/" + self.user_flat + "_"+o365servicename+".xml")
        f = open(WS+"tmp/" + self.user_flat + "_"+o365servicename+".xml", "r")
        r = requests.post("https://login.microsoftonline.com/extSTS.srf", data=f.read())
        Utils.request_trace(r)
        logger.info("Response from extSTS.srf for "+o365servicename+" is: " + r.text)

        tree = ET.fromstring(r.content)

        for value in tree.findall('.//{http://schemas.xmlsoap.org/ws/2005/02/trust}RequestedSecurityToken'):
            token = value.find('.//{http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd}BinarySecurityToken')
            office_token = token.text
        logger.info(o365servicename + " token is " + str(office_token))

        r = requests.post(self.domain_url + '/_forms/default.aspx?wa=wsignin1.0', data=office_token)
        Utils.request_trace(r)
        cookies = r.cookies
        logger.info("Cookies for "+o365servicename+" is " + str(cookies))
        r = requests.post(self.endpoint_contextinfo, cookies=cookies)
        Utils.request_trace(r)
        if r.status_code != 200:
            logger.error("Error authenticating to " + o365servicename + " with status: " + str(
                r.status_code) + " message: " + r.text)
            return (None, None)
        tree = ET.fromstring(r.content)
        for value in tree.findall('.//{http://schemas.microsoft.com/ado/2007/08/dataservices}FormDigestValue'):
            request_digest = value.text
        logger.info("Request Digest for "+ o365servicename +" is " + str(request_digest))
        return (request_digest, cookies)

    def _get_another_user_for_collab(self):
        users = self.get_all_users()
        users.remove(str(self.user))
        users.remove(str(self.admin))
        return random.choice(users)

    def get_users(self, users=None):
        logger.debug("inside get users")
        self.members_to_collaborate = self.get_users_for_O365Group(users)
        """
        # if not self.members_to_collaborate:
        internal_users_set = set(self.get_all_users()) - set(self.external_users)
        external_users_set = self.external_users
        logger.debug("Internal users set is: " + str(internal_users_set))
        logger.debug("External users set is: " + str(external_users_set))

        if "*" in users:
            try:
                if "i" not in users or "e" not in users:
                    logger.error("Please use correct syntax to select users . e.g. i1*e1*, due to incorrect synatx "
                                 "this will use by default *")
                internal_groups_choice, external_groups_choice = re.match(r'i(\d)\*e(\d)\*', users, re.M|re.I).groups()
                internal_users = random.sample(internal_users_set, int(internal_groups_choice))
                external_users = random.sample(external_users_set, int(external_groups_choice))
            except Exception as e:
                logger.warn("Error in randomly picking Users" + str(e))

            self.members_to_collaborate = ",".join(internal_users+external_users)
            self.members_to_collaborate=self.members_to_collaborate.split(",")
            logger.console("users to collaborate:" + str(self.members_to_collaborate))
        # else:
        #     logger.debug("Members to collaborate already exists.. hence using " + str(self.members_to_collaborate))
        """

    def get_mime_type(self, filename):
        """

        :param filename: filename for which mime type requires to be returned
        :return:
        """
        ext = os.path.splitext(filename)[1]
        if "doc" in ext:
            return "application/msword"
        if "docx" in ext:
            return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        if "ppt" in ext:
            return "application/vnd.ms-powerpoint"
        if "pptx" in ext:
            return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        if "xls" in ext:
            return "application/vnd.ms-excel"
        if "xlsx" in ext:
            return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if "txt" in ext:
            return "text/plain"

    @keyword("check if ${filename} is present in folder")
    def find_file_inside_folder_by_name(self, parent_id, file_name):
        """
        Private method
        Get file id by specifying file name and parent id

        Args:
            parent_id: Parent Id of the folder
            file_name: Name of the file for which file id to be found

        Returns:
            Boolean

        Raises:
            None
        """

        file_name = str(file_name)

        if parent_id is None:
            parent_id = self.mostrecentfolder

        logger.debug("Going to find " + str(file_name) + " inside folder ID: " + str(parent_id))
        file_id = None
        list_files = self.get_folder_info(parent_id, params="/Files").get("value")
        if list_files is None:
            return None
        if list_files:
            for file_properties in list_files:
                if file_properties["Name"] == file_name:
                    file_id = file_properties["ServerRelativeUrl"]
                    return file_id
        if file_id is None:
            logger.debug("Unable to find the file: "+ str(file_name))

        return None

    def cleanup(self):
        """
        going to clean onedrive account . Delete recursively all files and folders in account
        :return:
        """
        logger.debug("Going to delete folder for user: " + self.user)
        folder_info_json = self.get_folder_info(self.root_folder, params="/Folders").get("value")
        for item in folder_info_json:
            # if item.get("ServerRelativeUrl", None) is None:
            #     if "Forms" in item.get("odata.id"):
            #         continue
            #     self.delete_folder(item.get("odata.id"))
            # else:
            #     logger.debug(item.get("ServerRelativeUrl"))
            #     if "Forms" in item.get("ServerRelativeUrl"):
            #         continue
            #     self.delete_folder(item.get("ServerRelativeUrl"))
            logger.debug(item.get("ServerRelativeUrl"))
            if "Forms" in item.get("ServerRelativeUrl"):
                continue
            self.delete_folder(item.get("ServerRelativeUrl"))
        file_info_json = self.get_folder_info(self.root_folder, params="/Files").get("value")
        for item in file_info_json:
            logger.debug(item.get("ServerRelativeUrl"))
            self.delete_file(item.get("ServerRelativeUrl"))

    def verify_permission_on_files_inside_folder(self, permission_type, folder_id=None):
        """
        To verify permission on the files present in the the folder
        Args:
            permission_type: Changed permission to be verified. Ex. Editor, Viewer, None
            folder_id: Id of the folder

        Returns:
            Bool. Return True if verification is successfull for all the files

        """
        results = []

        if folder_id is None:
            folder_id = self.mostrecentfolder

        files_list = self.list_all_files_inside_folder(folder_id=folder_id)
        files_list = files_list["value"]

        if files_list:
            for files in files_list:
                self.collaborated_object = files['ServerRelativeUrl']
                result = self.verify_folder_permissions(permission_type)
                results.append(result)

            return all(results)

        else:
            logger.info("No files found inside folder:: %s" % folder_id)

    def list_all_files_inside_folder(self, folder_id):
        """
        To get all the files present in the folder
        Args:
            folder_id: Id of the folder

        Returns:
            List of all the files present inside folder

        """

        files_list = self.get_folder_info(folder_id, params="/Files")

        return files_list

    def add_metadata_tag(self, name, value, parent_id=0):
        """
        To add metadata tag to last updated file
        Args:
            name: name of the tag to set on the file
            value: tag value
        """

        file_id = self.lastuploadedfiles[-1]["fileid"]
        if file_id is None:
            return
        endpoint_GetFileByServerRelativeUrl = re.sub('GetFileByServerRelativeUrl(.*)',
                                                     'GetFileByServerRelativeUrl(\'' + file_id + '\')',
                                                     self.endpoint_GetFileByServerRelativeUrl)

        endpoint_set_metadata = endpoint_GetFileByServerRelativeUrl + '/ListItemAllFields'

        create_field_headers = {"X-RequestDigest": self.request_digest, "Content-Type": 'application/json',
                   "Accept": "application/json"}

        create_field = { 'Title': name,
                 'FieldTypeKind': 2,
                 'Required': 'false',
                 'EnforceUniqueValues': 'false',
                 'StaticName': name
                 }

        update_headers = {"X-RequestDigest": self.request_digest, "Content-Type": 'application/json',
                          "Accept": "application/json", "If-Match": "*"}

        tag_internal_name = self._get_fields_by_title(name, headers=create_field_headers)

        if tag_internal_name is None:
            res_create_field = requests.post(self.endpoint_create_field, headers=create_field_headers, cookies=self.cookies, json=create_field)
            if res_create_field.status_code == 201:
                logger.info("Response for create field is " + str(res_create_field.text))
            else:
                raise Exception("Failed to create metadata tag")
            tag_internal_name = self._get_fields_by_title(name, headers=create_field_headers)
            update_data = {tag_internal_name: value}
        else:
            update_data = {tag_internal_name: value}

        res_update_field = requests.patch(endpoint_set_metadata, headers=update_headers, cookies=self.cookies, json=update_data)
        if res_update_field.status_code == 204:
            logger.info("Response for update data is " + str(res_update_field.text))
        else:
            raise Exception("Failed to add tag to the metadata field")
        return res_update_field.status_code

    def validate_metadata_tag(self, name, value, parent_id=0):
        """
        To add metadata tag to last updated file
        Args:
            name: name of the tag to set on the file
            value: tag value
        """


        file_id = self.lastuploadedfiles[-1]["fileid"]
        if file_id is None:
            return
        endpoint_GetFileByServerRelativeUrl = re.sub('GetFileByServerRelativeUrl(.*)',
                                                     'GetFileByServerRelativeUrl(\'' + file_id + '\')',
                                                     self.endpoint_GetFileByServerRelativeUrl)

        endpoint_set_metadata = endpoint_GetFileByServerRelativeUrl + '/ListItemAllFields'

        update_headers = {'Authorization': "Bearer " + str(self.access_token) , 'Content-Type': 'application/json',
                          'Accept': "application/json"}

        #update_data = {name: value}

        res_update_field = requests.get(endpoint_set_metadata, headers=update_headers)
        res_update_field_json = res_update_field.json()
        logger.debug("response is %s and response type is %s" %(res_update_field,type(res_update_field_json)))

        if res_update_field.status_code in (204,200):
            logger.info("Response for fetching metadata is " + str(res_update_field.text))
        elif res_update_field.status_code in [401,403]:
            logger.debug("Token expired, refreshing and trying again")
            self._refresh_token()
            update_headers['Authorization'] = "Bearer " + str(self.access_token)
            res_update_field = requests.get(endpoint_set_metadata, headers=update_headers)
            res_update_field_json = res_update_field.json()
        else:
            raise Exception("Failed to fetch metadata field" + str(res_update_field))

        for key in list(res_update_field_json.keys()):
            if res_update_field_json[key]==value and key==name:
                logger.debug("found the metadata tag %s and value %s" %(name, value))
                return True
        logger.console("Metadata not found")

        return False

    @keyword("In ${SERVICE} create o365 group with ${visibility} ${name}")
    def create_o365_group(self, visibility, name, group_type=None):
        logger.debug("Inside create_o365_group" + str(name))
        result = True
        #name = name+".onedrive"
        if not group_type:
            group_type = ["Unified"]
            logger.debug("type of group is === " + str(group_type))
        else:
            logger.debug("type of group in else is === " + str(group_type))
        existing_groups = self.get_o365_groups()
        logger.debug("existing groups === " + str(existing_groups))

        for group in existing_groups["value"]:
            logger.debug("group is ==" + str(group))
            if group.get("displayName") == name:
                logger.debug("O365 groups named %s already exists " % (name))
                self.lastcreatedgroup.append(
                                        {
                                            "groupid":str(group.get("id")),
                                            "groupname":str(group.get("displayName")),
                                            "email":str(group.get("mail")),
                                            "id": str(group.get("id")),
                                            'fid': "c:0o.c|federateddirectoryclaimprovider|" + str(group.get("id")),
                                            'apiDisplayText': str(group.get("displayName")) + " Members",
                                            'apiDisplayName': str(group.get("displayName")) + " Members"

                                        }
                                     )
                logger.debug("last created group when group already exists == " + str(self.lastcreatedgroup))
                return result

        owner_id = self._get_user_ids([self.admin])
        logger.debug("===Group not found, creating new one===")
        mail_nick_name=name.replace(" ","")
        headers = {"Authorization": "Bearer " + self.access_token_graph, "content-type": "application/json"}
        data={
              "groupTypes": group_type,
              "displayName": name,
              "mailNickname": mail_nick_name,
              "mailEnabled": "true",
              "securityEnabled": "false",
              "visibility":visibility,
              "owners@odata.bind": ["https://graph.microsoft.com/v1.0/users/" + owner_id[0]]
        }
        logger.debug("endpoint is %s, headers are %s. data is %s " % (self.endpoint_groups,headers,data))
        #logger.debug("Data type is" + str(type(data)) )
        #logger.debug("headers type is" + str(type(headers)) )
        response_create_group = requests.post(url=self.endpoint_groups, headers=headers, data=json.dumps(data))
        new_group=response_create_group.json()
        logger.debug("response is " + str(new_group))
        if response_create_group.status_code in (200,201):

            self.lastcreatedgroup.append(
                                            {
                                                "groupid":str(new_group.get("id")),
                                                "groupname":str(new_group.get("displayName")),
                                                "email":str(new_group.get("mail")),
                                                "id": str(new_group.get("id")),
                                                'fid': "c:0o.c|federateddirectoryclaimprovider|" + str(new_group.get("id")),
                                                'apiDisplayText': str(new_group.get("displayName")) + " Members",
                                                'apiDisplayName': str(new_group.get("displayName")) + " Members"

                                            }
                                         )
            logger.debug("last created group when new group is created  == " + str(self.lastcreatedgroup))
        else:
            logger.error("Group creation failed due to " + str(response_create_group._content))
            result=False
        return result

    @retry_handler
    def get_o365_groups(self):
        headers = {"Authorization": "Bearer " + self.access_token_graph, "Accept": "application/json"}
        response_get_all_groups = None
        response_get_all_groups = requests.get(url=self.endpoint_groups, headers=headers)

        if response_get_all_groups.status_code in [200,201]:
            return response_get_all_groups.json()
        elif response_get_all_groups.status_code in [401,403]:
            logger.debug("Refresh token and retry again")
            raise Exception
        else:
            logger.debug("Failed to get the token")


    @retry_handler
    def _get_user_ids(self,mail):

        logger.debug("Inside _get_user_ids and mail value is %s" % (mail))
        user_id=[]
        headers = {"Authorization": "Bearer " + self.access_token_graph, "Accept": "application/json"}
        response_get_all_users = requests.get(url=self.endpoint_users, headers=headers)
        if response_get_all_users.status_code == 200:
            for user in response_get_all_users.json().get("value"):
                # logger.debug("user is %s" % (user))
                if user.get("mail") in mail:
                    user_id.append(user.get('id'))
                    # logger.debug("user id list is %s ==== " % (user_id))

        elif response_get_all_users.status_code in [403,401]:
            logger.debug("Got a %s error.." %(response_get_all_users.status_code))
            self._refresh_token()
            raise Exception
        else:
            logger.debug("Error in getting member id " + response_get_all_users.text)
            return None
        return user_id

    @keyword("In ${SERVICE} add o365 group members")
    def add_members_to_o365_group(self):
        result = False
        members_to_add_to_group = []
        group_id = self.lastcreatedgroup[-1]["groupid"]
        endpoint_group_member_url = self.endpoint_groups_url + str(group_id) + "/members/$ref"
        o365_users = BuiltIn().replace_variables('${o365_users}')
        group_users = self.get_users_for_O365Group(o365_users)
        # get all users
        all_users = self.get_all_users()

        # finding diff of users from what is already selected to be part of policy
        # setdiff = set(all_users)-set(self.members_to_collaborate)
        # setdiff = set(all_users) - set(group_users)
        # randomly chosing another user to collaborate
        # int_mem = random.choice(list(setdiff))

        # adding to the variable members_added
        # members_to_add_to_group.append(int_mem)
        members_to_add_to_group.extend(group_users)
        self.lastcreatedgroup[-1]["members"]=members_to_add_to_group

        # for member in self.members_to_collaborate:
        #     members_to_add_to_group.append(member)

        member_ids_to_add = self._get_user_ids(members_to_add_to_group)
        logger.debug("adding members now %s======" % (member_ids_to_add))
        headers = {"Authorization": "Bearer " + self.access_token_graph, "content-type": "application/json"}
        for member_id in member_ids_to_add:
            data = {
                "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/"+str(member_id)
            }
            retry = 3
            try:
                if self._isMemberPresentInGroup(group_id, member_id):
                    result = True
                    continue
                response_add_members = requests.post(url=endpoint_group_member_url , headers=headers, data=json.dumps(data))
            except Exception as e:
                logger.error(" ======member not added====== " + str(e) )
                raise Exception

            logger.debug("response for adding member id %s is %s" % (member_id,response_add_members))
            if response_add_members.status_code == 204:
                logger.debug("member added successfully")
                result = True
            else:
                if response_add_members.status_code == 403:
                    self.access_token_graph = shutil.get_access_token(self.cspid, self.instance_id,
                                                                  resource="https://graph.microsoft.com",
                                                                  decrypt_token=True)
                logger.debug("Member not added  due to error, so retrying")
                #logger.error("Member not added due to " + str(response_add_members._content))
                while retry > 0:
                    time.sleep(30)
                    response_add_members = requests.post(url=endpoint_group_member_url, headers=headers,
                                                         data=json.dumps(data))
                    if response_add_members.status_code == 204:
                        logger.debug("member added successfully")
                        result = True
                        break
                    else:
                        result = False
                    retry = retry - 1
        if result:
            return result
        else:
            raise Exception

    @keyword("In ${SERVICE} delete latest O365 group")
    def delete_o365_group(self):
        result = []
        for group in self.lastcreatedgroup:
            headers = {"Authorization": "Bearer " + self.access_token_graph, "content-type": "application/json"}
            group_delete_endpoint = self.endpoint_groups + "/" + str(group.get("groupid"))
            group_delete_response = requests.delete(url=group_delete_endpoint , headers=headers)
            if group_delete_response.status_code == 204:
                logger.debug("Deleted group " + str(group.get("groupname")))
                result.append(True)
            else:
                logger.error("Could not delete group %s due to following error %s " % (group.get("name"),group_delete_response._content))
                result.append(False)
        logger.debug("No group to delete")
        return all(result)

    @keyword("In ${SERVICE} ${enable} ${feature} feature")
    def enable_shared_link_feature(self, enable, feature):

        self.shared_link_url = self.watchtower_url + "/v1/" + str(self.shutil.gettenantid()) + "/config/" +str(feature)


        if enable == 'enable':
            logger.debug("Going to enable %s feature with url %s"  % (feature, self.shared_link_url))
            try:
                o365_enable_shared_link = requests.request("POST", self.shared_link_url)
            except Exception as e:
                logger.debug(e)
                return False

            if o365_enable_shared_link.status_code == 200:
                logger.debug("The feature %s is enabled %s:- " % (feature,o365_enable_shared_link.text))

                return True
            else:
                return False
        elif enable == 'disable':
            logger.debug("Going to disable %s feature with url %s" % (feature, self.shared_link_url))
            try:
                o365_enable_shared_link = requests.request("DELETE", self.shared_link_url)
            except Exception as e:
                logger.debug(e)
                return False

            if o365_enable_shared_link.status_code == 200:
                logger.debug("The %s Feature disabled %s :- " % (feature, o365_enable_shared_link.text))
                return True
            else:
                return False


    @keyword("fetch event stats from")
    def fetch_event_stats(self,Sleep_Enable=False):
        ''' Fetch the Event stats for the CSP provided '''
        logger.console("Fetch event counts for all the CSPs of a tenant")
        zeus_ip = self.shutil.get_zeus_ip()
        if Sleep_Enable:
            logger.console("Waiting for 10 mins for EventStats to get updated")
            time.sleep(600)
        url = 'http://'+zeus_ip+':9000/v1/tenant-health/'+str(self.tenantid)
        response = requests.get(url)
        if response.status_code == 200:
            logger.console("Able to get tenant health")
        else:
            logger.console("Failed to get the tenant health with status code"+ response.status_code)
            return False
        try:
            api_data = json.loads(response.text)
        except Exception:
            logger.console("Failed to parse json for api fetch event counts")
            return False
        event_stats = []
        for i in api_data:
            if int(i['instance_id'] != int(self.instanceid)):
                continue
            else:
                event_stats = i['event_stat_list']
                break
        csp_event_stats = {}
        for item in event_stats:
            name = item['name']
            csp_event_stats[name] = item
        return csp_event_stats

    def _get_fields_by_title(self, tag_name, headers):

        response = requests.get(self.endpoint_create_field, headers=headers, cookies=self.cookies, verify=False)
        if response.status_code == 200:
            logger.console("Able to get fileds by title")
        else:
            logger.console("Failed to get the fields by title"+ response.status_code)
            return None

        json_res = json.loads(response.text)
        fields = [item for item in json_res['value'] if item['Title'] == tag_name]
        if len(fields) == 0:
            logger.console("Did not find tag :%s on the root folder" % tag_name)
            return None

        property_name = fields[0].get('EntityPropertyName', None)
        return property_name

    @keyword("In ${SERVICE} ${action} external sharing in O365")
    def external_sharing(self, action="enable"):
        """
        Enable Or Disable External Sharing For The SharePoint Site

        Args:
            action: enable or disable

        Returns:
            List containing JSON of file properties

        Raises:
            None
        """
        sharesettings_file_path = self.shutil.env_info.get("TEMPLATES_FOLDER")
        sharesettings_file = 'ShareSetting.xml'
        setting = 2 if action == "disable" else 3
        kwargs = {
            'sharesetting': int(setting),
            'domain_url':str(self.domain_url),
        }

        env = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath=sharesettings_file_path),
                                 trim_blocks=True,
                                 lstrip_blocks=True
                                 )

        template = env.get_template(sharesettings_file)
        template_after_render = template.render(kwargs)
        sharesettings_data = template_after_render
        cspid = "16131" # tenant administration in OneDrive can be done only with SharePoint token
        from skybot.lib.GetInstances import get_instance
        instance_id = get_instance.get_instance_id(cspid)
        try:
            self.access_token_admin = self.shutil.get_access_token(cspid, resource=self.domain_admin_url, instance_id=instance_id)
        except Exception as e:
            logger.error(e)
            raise Exception
        endpoint_sharing = self.domain_admin_url + "/_vti_bin/client.svc/ProcessQuery"
        headers = {"Authorization": "Bearer " + self.access_token_admin, "Accept": "application/xml"}

        response_sharing = requests.post(url=endpoint_sharing, headers=headers, data=str(sharesettings_data))
        logger.debug("The response is %s" % (response_sharing.json()))

        if response_sharing.status_code in [401,403]:
            logger.debug("Got a %s error.." %(response_sharing.status_code))
            raise Exception
        elif response_sharing.status_code == 200:
            logger.console("Successfully set the sharing on the site")
            ext_share = response_sharing.json()[4]["DefaultSharingLinkType"]
            if ext_share == 3:
                logger.debug("External Sharing is enabled on the site")
                return True
            elif ext_share == 2:
                logger.debug("Internal Sharing is enabled on the site")
                return True
        return False

    @keyword("In ${SERVICE} update last uploaded file with data ${data}")
    def update_file(self,data='NULL'):
        """
        To overwrite the last uploaded file with input data

        Args:
            data : Data that needs to be updated in to the document

        Returns:
            Returns the file_id of last uploaded file

        Raises:
            Exception in case of any token refresh issues
        """
        filename = self.lastuploadedfiles[-1].get('filename')
        filecontent = data
        file_id = self.lastuploadedfiles[-1]["fileid"]
        if file_id is None:
            return
        headers = {"X-RequestDigest": self.request_digest, "Content-Type": self.get_mime_type(filename=os.path.basename(filename)),
                   "Accept": "application/json"}
        parent_id, self.mostrecentfolder = [0 if self.mostrecentfolder is None else self.mostrecentfolder] * 2
        if parent_id:
            endpoint_GetFolderByServerRelativeUrl = re.sub('GetFolderByServerRelativeUrl(.*)',
                                                           'GetFolderByServerRelativeUrl(\'' + parent_id + '\')',
                                                           self.endpoint_GetFolderByServerRelativeUrl)
        else:
            endpoint_GetFolderByServerRelativeUrl = self.endpoint_GetFolderByServerRelativeUrl
        update_url = endpoint_GetFolderByServerRelativeUrl + '/Files/add(url=\'' + os.path.basename(filename) \
                     + '\', overwrite=true)'
        logger.info("URL to update is " + update_url)

        response_update_file = requests.post(update_url, headers=headers, cookies=self.cookies, data=filecontent)

        if response_update_file.status_code == 200:
            logger.info("Response for upload file is " + str(response_update_file.text))

        elif response_update_file.status_code in [401,403]:
            logger.debug("Got a 403 error refreshing access token...")
            self._refresh_token()
            raise Exception

        logger.info("Response post upload file is: " + response_update_file.text)
        return file_id

    @keyword("get version count of last uploaded file")
    def get_version_details(self,filename=None,site=None):
        """
        To get the version count for last uploaded file

        Args:
            None

        Returns:
            version count: returns the version count of the last uploaded file

        Raises:
            Exception in case of any token refresh issues
        """
        retry = 3
        if not filename:
            file_id = self.lastuploadedfiles[-1]["fileid"]
        else:
            file_id = os.path.join(self.mostrecentfolder,filename)
        headers = {"X-RequestDigest": self.request_digest, "Accept": "application/json"}

        endpoint_GetFileByServerRelativeUrl = re.sub('GetFileByServerRelativeUrl(.*)','GetFileByServerRelativeUrl(\'' + file_id + '\')',self.endpoint_GetFileByServerRelativeUrl)
        if site:
            req_dig = self.req_digest(self.domain_url + "/sites/" + site)
            headers["X-RequestDigest"] = req_dig
            endpoint_GetFileByServerRelativeUrl  = re.sub('(.*)\/_api', self.domain_url + '/sites/' + site + '/_api', endpoint_GetFileByServerRelativeUrl )
        get_versions_url = endpoint_GetFileByServerRelativeUrl + '?$expand=Versions'

        for i in range(retry):
            response_get_versions = requests.get(get_versions_url, headers=headers, cookies=self.cookies)
            logger.debug(response_get_versions.json())
            if response_get_versions.status_code == 200:
                logger.info("Response for getting version count is " + str(response_get_versions.text))
                version_dict = {}
                version_dict['latest_version'] = int(eval(response_get_versions.json()["UIVersionLabel"]))
                version_history = [float(x['VersionLabel']) for x in response_get_versions.json()['Versions']]
                version_dict['version_history'] = version_history
                version_dict['version_count'] = len(version_history) + 1
                # version_count = int(eval(response_get_versions.json()["UIVersionLabel"]))
                return version_dict
            elif response_get_versions.status_code in [401,403]:
                logger.debug("Got a 403 error refreshing access token...")
                self._refresh_token()
                if site:
                    req_dig = self.req_digest(self.domain_url + "/sites/" + site)
                    headers["X-RequestDigest"] = req_dig
                if retry == 3:
                    raise AccesstokenError
            else:
                raise Exception("Failed to retrieve version of the file")

    def _isMemberPresentInGroup(self, groupid, member):
        """
        Check if member is present in the group or not
        :param groupid: O365 group id
        :param member: User id
        :return: Boolean
        """
        retry = 3
        result = False
        headers = {"Authorization": "Bearer " + self.access_token_graph, "content-type": "application/json"}
        endpoint_group_member_url = self.endpoint_groups_url + str(groupid) + "/members/$ref"
        while retry > 0:
            response = requests.get(url=endpoint_group_member_url , headers=headers)
            retry = retry - 1
            if response.status_code == 200:
                try:
                    members_list = json.loads(response.text)["value"]
                    if isinstance(members_list, list) and len(members_list) == 0:
                        return False
                    member_ids = [item['@odata.id'].split("/")[-2] for item in members_list]
                    return True if member in member_ids else False
                except KeyError:
                    logger.error("Response does not contain value filed")
                    return False
            else:
                logger.debug("Get group members call did not go through .. retrying")
                time.sleep(30)

        if not result:
            raise Exception

    def _validate_drm_action(self, file_id, filename):
        """
        drm_action validation is three step process:
        1. Verify file size before and after applying drm. File size should increase after drm
        2. Verify checksume before and after applying drm. It should not be same
        3. After applying drm action, list the file tags and it should have 'DRM_PROTECTED' as one of the tags
        :param file_id:
        :return: Boolean
        """
        result = False

        self.download_file(file_id)
        download_content_to = WS + 'tmp/' + str(time.time()) + '.' + filename.split('.')[-1]
        with open(download_content_to, 'wb') as f:
            f.write(self.lastdownloadedfilecontents)
        from tika import parser
        try:
            content = parser.from_file(download_content_to)
        except Exception:
            if os.path.exists(download_content_to):
                os.remove(download_content_to)
                return False

        if 'PROTECTED BY IONIC SECURITY' in content['content']:
            logger.debug("File is DRM protected")
            result=True

        if os.path.exists(download_content_to):
            os.remove(download_content_to)

        return result

    @keyword("In ${SERVICE} create ${listType} in o365 with name ${listName}")
    def create_list(self, listType, listName):
        """
        Create generic list or document library in O365
        :param listType:Generic List = 0; Document Library = 1;
        :param listName: Name of the list or doclibrary
        """
        if listType.lower() == "document library":
            BaseType = 1
            BaseTemplate = 101
        elif listType.lower() == "generic list":
            BaseType = 0
            BaseTemplate = 100
        else:
            logger.debug("list type is not supported")
            raise Exception
        logger.debug("Going to Create: " + str(listType))
        headers = {"Authorization": "Bearer " + self.access_token, "Accept": "application/json;odata=verbose"}
        headers.update({"Content-Type": "application/json;odata=verbose"})
        payload = {
            "__metadata": {"type": "SP.List"},
            "AllowContentTypes": 'true',
            "BaseType": BaseType,
            "BaseTemplate": BaseTemplate,
            "ContentTypesEnabled": 'true',
            "Description": "Creating list via API",
            "Title": listName
        }

        response_create_list = requests.post(self.endpoint_create_list, headers=headers, json=payload)
        if response_create_list.status_code == 500 and \
            "title already exists" in response_create_list.json()['error']['message']['value']:
            logger.info("List or Doc Lib already exists " + str(response_create_list.json()))
            self._get_list_guid(listName)
        elif response_create_list.status_code in [200,201]:
            self.list_guid = response_create_list.json()["d"]["Id"]
            logger.debug("Created the List: " + listName + "& GUID is: " + self.list_guid)
        else:
            logger.error("Failed to Create the List")
            raise Exception
        return response_create_list.json()

    @keyword("In ${SERVICE} delete o365 list ${listName}")
    def delete_list(self, listName=None):
        """
        Delete generic list or document library in O365
        :param listName: Name of the list or doclibrary
        """
        if listName is not None:
            logger.debug("Going to get GUID of list: " + str(listName))
            self._get_list_guid(listName)
        if self.list_guid is not None:
            logger.debug("Deleting the List with GUID: " + self.list_guid)
            headers = {"Authorization": "Bearer " + self.access_token, "X-HTTP-Method": "DELETE", "IF-MATCH": "*"}
            self.endpoint_delete_list = self.endpoint_list + "(guid'" + self.list_guid + "')?@target='" + self.endpoint_host_web_url + "'"
            response_delete_list = requests.post(self.endpoint_delete_list, headers=headers)
            logger.debug("Deleted the list with GUID: " + self.list_guid)
            return True
        else:
            logger.warn("No existing list to be Deleted")
            return None

    def _get_list_guid(self, listName):
        """
        Gets the Guid of generic list or document library in O365
        :param listName: Name of the list or doclibrary
        """
        headers = {"Authorization": "Bearer " + self.access_token, "Accept": "application/json;odata=verbose"}
        headers.update({"Content-Type": "application/json;odata=verbose"})
        self.endpoint_get_list = self.endpoint_list + "/getbytitle('" + listName + "')?@target='" + self.endpoint_host_web_url + "'"
        response_get_list = requests.get(self.endpoint_get_list, headers=headers)
        if response_get_list.status_code != 200 and \
                        "does not exist" in response_get_list.json()['error']['message']['value']:
            logger.error("Error:" + str(response_get_list.json()['error']['message']['value']))
            raise Exception
        self.list_guid = response_get_list.json()["d"]["Id"]
        return self.list_guid

    # @keyword("In ${SERVICE} validate aip label ${label_info}")
    # def validate_aip_label(self, label_info):
    #     """
    #     To verify the applied AIP label name on the given file
    #     Args:
    #         file_path: path of the file
    #         label_info: Colon separated AIP instance id, name, type of label, label id, name (ex- 2067:AIP-shneuqa:62e5eb5c-da1f-412a-ba31-04818b4c53d7:passport-label)
    #     """
    #     label = str(label_info).split(':')
    #     label_id = label[2]
    #     label_name = label[3]
    #     root_file_path = self.default_root_folder + self.mostrecentfoldername + '/' + self.lastuploadedfiles[-1]["filename"]
    #     aip_label_details = "AIP Label ID: " + label_id + ", AIP Label Name: " + label_name
    #     logger.debug("File Path: " + root_file_path + ", " + aip_label_details)
    #     file_download_url = self.endpoint_GetFileByServerRelativePath + root_file_path + "'"
    #     request_headers = {'Authorization': "Bearer " + str(self.access_token), 'Content-Type': 'application/json',
    #                        'Accept': "application/json"}
    #     response = requests.get(file_download_url, headers=request_headers)
    #     if response.status_code in (204, 200):
    #         logger.info("Successfully retrieved the file content to verify AIP labels")
    #     elif response.status_code in [401, 403]:
    #         logger.debug("Token expired, refreshing and trying again")
    #         self._refresh_token()
    #         raise Exception
    #     else:
    #         raise Exception("Failed to fetch file content " + str(response.content))

        # if str(response.content).find(str(label_id)) and str(response.content).find(str(label_name)):
        #     logger.debug("Got " + aip_label_details + " in the file " + root_file_path)
        #     return True
        # logger.debug("Didn't find " + aip_label_details + " in the file " + root_file_path)
        # return False

    @keyword("Log into ${service} as ${user}")
    def login_to_service_ui(self, user):
        app_name = "OneDrive"
        super(OneDrive, self).login_to_service_ui(user)
        self.driver.set_window_size(1600, 1200)
        if user == 'user':
            username = self.user
            password = self.password
        elif user == 'external user':
            username = self.externalcollaborator
            password = self.externalcollaborator_password
        try:
            serviceurl = "https://login.microsoftonline.com"
            CommonHelper.go_to_url(self.driver, serviceurl)
            CommonHelper.wait_for_seconds(3)
            CommonHelper.wait_for_element_and_sendkeys(self.driver, LocatorType.XPATH, O365_locators['login_username'],
                                                       username)
            CommonHelper.wait_for_element_and_click(self.driver, LocatorType.XPATH, O365_locators['login_next_button'])
            CommonHelper.wait_for_seconds(5)
            CommonHelper.wait_for_element_and_sendkeys(self.driver, LocatorType.XPATH, O365_locators['login_password'],
                                                       password)
            CommonHelper.wait_for_element_and_click(self.driver, LocatorType.XPATH, O365_locators['login_signIn_button'])
            CommonHelper.wait_for_seconds(4)

            keep_sign_ele = CommonHelper.get_elements_from_locator_type(self.driver, LocatorType.XPATH,
                                                                        O365_locators['login_keep_signedIn'])
            if len(keep_sign_ele) > 0:
                CommonHelper.wait_for_element_and_click(self.driver, LocatorType.XPATH,
                                                        O365_locators['login_keep_signedIn'])
            CommonHelper.wait_for_seconds(15)

            if CommonHelper.is_element_displayed(self.driver, LocatorType.XPATH, str(O365_locators['homepage_login']), 120):
                logger.console("Succesfully logged into o365 account")
                return True
            else:
                return False

        except Exception as e:
            ss_name = time.time()
            logger.console("Login to O365 failed, taking screenshot")
            self.driver.save_screenshot("%s.png" % ss_name)
            return False

    def enable_api_access(self, params, driver_obj, api, wait, EC, By):
        time.sleep(5)
        wait.until(EC.visibility_of_element_located((By.XPATH, api.page_elements_dict["common"]["preReqCheck"])))
        driver_obj.find_element_by_xpath(api.page_elements_dict['common']['preReqCheck']).click()
        logger.debug("Clicked Prerequisites")
        time.sleep(3)
        wait.until(EC.element_to_be_clickable((By.XPATH, api.page_elements_dict["common"]["nextButton"])))
        driver_obj.find_element_by_xpath(api.page_elements_dict['common']['nextButton']).click()
        logger.debug("Clicked Next")
        time.sleep(5)
        driver_obj.find_element_by_xpath(api.page_elements_dict['common']['credsButton']).click()
        logger.debug("Clicked Provide Credentials")
        time.sleep(5)
        handles=driver_obj.window_handles
        current=driver_obj.current_window_handle
        driver_obj.switch_to.window(handles[1])

        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['oneDriveEmail']).send_keys(str(params['email']))
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['nextButton']).click()
        time.sleep(10)
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['password']).send_keys(str(params['password']))
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['nextButton']).click()
        time.sleep(10)
        driver_obj.find_element_by_xpath(api.page_elements_dict['OneDrive']['acceptButton']).click()
        time.sleep(10)
        driver_obj.switch_to.window(current)
        return True

    def click_url(self, url):
        url = [x for sublist in url for x in sublist]
        url = list(dict.fromkeys(url))
        urls= [value for value in url if value is not False]
        try:
            link_to_click = self.select_url(urls, OneDrive.URL_PATTERN_TO_FIND)[0]
        except Exception as e:
            logger.console("Link received is None")
            return False

        if link_to_click:
            logger.console("Link to click in OneDrive= " + str(link_to_click))
            CommonHelper.go_to_url(self.driver, link_to_click)
            #waiting for the page to load
            CommonHelper.wait_for_seconds(8)
            if CommonHelper.is_element_displayed(self.driver, LocatorType.XPATH,
                                                 O365_locators['onedrive_item_removed_page']):
                logger.console("OneDrive link is expired- The user doesnot have permission to access this file")
                return True
            else:
                logger.console("OneDrive link accessible")
                return False
        else:
            logger.error("No OneDrive link received")
            return False

    def get_users_for_O365Group(self, users):
        internal_users_set = set(self.get_all_users()) - set(self.external_users)
        external_users_list = self.external_users
        internal_users_list = [i_user for i_user in list(internal_users_set) if i_user.lower() not in (self.user.lower(), self.email.lower())]
        #internal_users_list = [i_user for i_user in list(internal_users_set)]
        logger.debug("Internal users list is: " + str( ))
        logger.debug("External users list is: " + str(external_users_list))
        o365_group_users = []

        if "*" in users:
            try:
                if "i" not in users or "e" not in users:
                    logger.error("Please use correct syntax to select users . e.g. i1*e1*, due to incorrect synatx "
                                 "this will use by default *")
                internal_groups_choice, external_groups_choice = re.match(r'i(\d)\*e(\d)\*', users, re.M|re.I).groups()
                internal_users = sorted(internal_users_list)[:int(internal_groups_choice)]
                external_users = sorted(external_users_list)[:int(external_groups_choice)]
                # internal_users = random.sample(internal_users_set, int(internal_groups_choice))
                # external_users = random.sample(external_users_set, int(external_groups_choice))
                o365_group_users.extend(internal_users)
                o365_group_users.extend(external_users)
                logger.debug("O365 group users:" + str(o365_group_users))
            except Exception as e:
                logger.warn("Error in randomly picking Users" + str(e))
        return o365_group_users

    def get_list_name(self):
        if getattr(self, 'library', None) is None:
            title = "Documents"
        else:
            title = self.library
        return title


if __name__ == '__main__':
    #from skybot.OF.lib.core.SkyHighDashboard.ShnDlpInterface import ShnDlpUtil
    #shutil = ShnDlpUtil("qaautoregression", 80560, "Welcome2dlp#", "dlpp1qa@gmail.com", None, None, use_token=True)
    from skybot.lib import SHNInterface
    SHNInterface.myenv.project = "OF"
    SHNInterface.myenv.zeus_admin_uname = "admin@spop.com"
    SHNInterface.myenv.zeus_admin_pwd = "admin"
    # SHNInterface.myenv = SHNInterface.Util('qaautoregression', 80560, 'Welcome2dlp#', 'dlpp1qa@gmail.com', None, None, use_token=True)
    from skybot.OF.lib.core.SkyHighDashboard.ShnDlpInterface import ShnDlpUtil
    shutil = ShnDlpUtil('qaautoregression', 79164, 'Skyhigh_1234', 'vidisha_aggarwal@mcafee.com', None, None, use_token=True)
    # SHNInterface.myenv = shutil
    Od = OneDrive(shutil, "qaautoregression", 79164, 3210, "user2@skyhightest2.onmicrosoft.com", instance_id=10844)
    Od.as_user("user1@skyhightest3.onmicrosoft.com")
    Od.cleanup()
    # Od.delete_file("7BE27B6A9B-F2BF-46AC-85D7-35C091A2E81E")
    #Od.create_folder('Test7')

