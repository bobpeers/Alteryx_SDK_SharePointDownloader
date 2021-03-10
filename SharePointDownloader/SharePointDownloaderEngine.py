import AlteryxPythonSDK as Sdk
import xml.etree.ElementTree as Et
import os
import fnmatch
from urllib.parse import urlparse
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
from requests_ntlm import HttpNtlmAuth

class AyxPlugin:
    def __init__(self, n_tool_id: int, alteryx_engine: object, output_anchor_mgr: object):
        # Default properties
        self.n_tool_id: int = n_tool_id
        self.alteryx_engine: Sdk.AlteryxEngine = alteryx_engine
        self.output_anchor_mgr: Sdk.OutputAnchorManager = output_anchor_mgr
        
        # Custom properties
        self.url: str = None
        self.site: str = None
        self.docs: str = None
        self.version: str = None
        self.username: str = None
        self.password: str = None
        self.filespec: str = None
        self.savepath: str = None
        self.is_initialized = True
        pass

    def pi_init(self, str_xml: str):
        self.output_anchor = self.output_anchor_mgr.get_output_anchor('Output')
        self.site = Et.fromstring(str_xml).find('site').text  if 'site' in str_xml else None
        self.docs = Et.fromstring(str_xml).find('docs').text  if 'docs' in str_xml else None
        self.version = Et.fromstring(str_xml).find('version').text  if 'version' in str_xml else None       
        self.username = Et.fromstring(str_xml).find('username').text if 'username' in str_xml else None
        self.password = Et.fromstring(str_xml).find('password').text if 'password' in str_xml else None
        self.filespec = Et.fromstring(str_xml).find('filespec').text if 'filespec' in str_xml else None
        self.savepath = Et.fromstring(str_xml).find('save_path').text if 'save_path' in str_xml else None

        if self.password is not None:
            self.password = self.alteryx_engine.decrypt_password(self.password, 0)

        # Validity checks.
        if self.site is None:
            self.display_error_msg('Site URL field cannot be empty.')
            
        elif self.docs is None:
            self.display_error_msg('Documents folder field cannot be empty.')

        elif self.version is None:
            self.display_error_msg('Version field cannot be empty.')            
        
        elif self.username is None:
            self.display_error_msg('Username field cannot be empty.')
            
        elif self.password is None:
            self.display_error_msg('Password field cannot be empty.')
        
        elif self.filespec is None:
            self.display_error_msg('File specification field cannot be empty.')
            
        elif self.savepath is None:
            self.display_error_msg('Save location field cannot be empty.')

        elif not os.path.exists(self.savepath):
            self.display_error_msg('Save location does not exist. Create the folder first.')

    def pi_add_incoming_connection(self, str_type: str, str_name: str) -> object:
        return self

    def pi_add_outgoing_connection(self, str_name: str) -> bool:
        return True

    def build_record_info_out(self):
        """
        A non-interface helper for pi_push_all_records() responsible for creating the outgoing record layout.
        :param file_reader: The name for csv file reader.
        :return: The outgoing record layout, otherwise nothing.
        """

        record_info_out = Sdk.RecordInfo(self.alteryx_engine)  # A fresh record info object for outgoing records.
        #We are returning a single column and a single row. 
        
        record_info_out.add_field('FilePath', Sdk.FieldType.string, 100)
        return record_info_out

    def download(self) -> list:

        dl_files = []

        # parse url
        parsed = urlparse(self.site)
        scheme = 'https' if parsed.scheme == '' else parsed.scheme
        version = 365

        if self.version == '365':
            sp_version = Version.v365
        else:
            sp_version = Version.v2007
        
        try:
            if sp_version == Version.v365:
                authcookie = Office365(f'{scheme}://{parsed.netloc}', username=self.username, password=self.password).GetCookies()
            else:
                cred = HttpNtlmAuth(self.username, self.password)
        except:
            raise Exception(f'Unable to authenticate using supplied user name and password.')  
        else:
            self.display_info('Sucessfully authnticated')             

        try:
            if sp_version == Version.v365:
                site = Site(self.site, version=sp_version, authcookie=authcookie)
            else:
                site = Site(self.site, version=sp_version, auth=cred)
        except:
            raise Exception(f'{self.site} is not a valid site')
        else:
            self.display_info(f'Sucessfully accessed site {self.site}')  
                                
        # build path to document folder
        doc_path = os.path.join(parsed.path, self.docs)

        try:
            folder = site.Folder(doc_path)
            for f in folder.files:
                fname = f['Name']
                if fnmatch.fnmatch(fname, self.filespec):
                    dest = os.path.join(self.savepath, fname)
                    with open(dest, mode='wb') as file:
                        file.write(folder.get_file(fname))
                        dl_files.append(dest)
        except:
            raise Exception(f'Unable to download files from {self.docs}')
        
        return dl_files

    def pi_push_all_records(self, n_record_limit: int) -> bool:

        if not self.is_initialized:
            return False

        record_info_out = self.build_record_info_out()  # Building out the outgoing record layout.
        self.output_anchor.init(record_info_out)  # Lets the downstream tools know of the outgoing record metadata.
        record_creator = record_info_out.construct_record_creator()  # Creating a new record_creator for the new data.

        if self.alteryx_engine.get_init_var(self.n_tool_id, 'UpdateOnly') == 'True':
            return False

        try:
            filelist = self.download()
        except Exception as e:
            self.display_error_msg(str(e))
            return False
        
        for f in filelist:
            record_info_out[0].set_from_string(record_creator, f)
        
            #record_info_out[0].set_from_string(record_creator, self.password)
            out_record = record_creator.finalize_record()
            self.output_anchor.push_record(out_record, False)  # False: completed connections will automatically close.
            record_creator.reset()  # Resets the variable length data to 0 bytes (default) to prevent unexpected results.

        if len(filelist) > 0:
            self.display_info(f'Downloaded {len(filelist)} files to {self.savepath}')
        else:
            self.display_info(f'No files matched file specification {self.filespec}')
        self.output_anchor.close()  # Close outgoing connections.
        return True

    def pi_close(self, b_has_errors: bool):
        self.output_anchor.assert_close()  # Checks whether connections were properly closed.

    def display_error_msg(self, msg_string: str):
        self.alteryx_engine.output_message(self.n_tool_id, Sdk.EngineMessageType.error, msg_string)
        self.is_initialized = False

    def display_info(self, msg_string: str):
        self.alteryx_engine.output_message(self.n_tool_id, Sdk.EngineMessageType.info, msg_string)


class IncomingInterface:
    def __init__(self, parent: AyxPlugin):
        pass

    def ii_init(self, record_info_in: Sdk.RecordInfo) -> bool:
        pass

   
    def ii_push_record(self, in_record: Sdk.RecordRef) -> bool:
        pass

    def ii_update_progress(self, d_percent: float):
        # Inform the Alteryx engine of the tool's progress.
        pass


    def ii_close(self):
        pass
