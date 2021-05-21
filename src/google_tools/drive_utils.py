class GoogleDriver:

    def _ExtractAll(self, d, q):
        page_token = None
        while True:
            response = self.service.files().list(q=q,
                                                 spaces='drive',
                                                 fields='nextPageToken, files(id, name)',
                                                 pageToken=page_token).execute()
            files = response['files']
            d.update({file['name']:file['id'] for file in files})
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break

    def _UpdateDriver(self):
        self._ExtractAll(d=self.gdrive_folders,
                         q="mimeType='application/vnd.google-apps.folder'")
        self._ExtractAll(d=self.gdrive_files,
                         q="mimeType!='application/vnd.google-apps.folder'")

    def __init__(self, service):
        self.service = service
        self.gdrive_folders = {}
        self.gdrive_files = {}
        self._ExtractAll(d=self.gdrive_folders,
                         q="mimeType='application/vnd.google-apps.folder'")
        self._ExtractAll(d=self.gdrive_files,
                         q="mimeType!='application/vnd.google-apps.folder'")

    def GetFolders(self):
        return self.gdrive_folders

    def GetFiles(self):
        return self.gdrive_files

    def SearchFolder(self, folder_name):
        assert folder_name in self.gdrive_folders, 'Folder not found'
        _id = self.gdrive_folders[folder_name]
        folder_contents = self.service.files().list(q=f"'{_id}' in parents", fields="*").execute()
        return folder_contents

    def SearchFile(self, file_name):
        file_contents = self.service.files().list(q=f"name = '{file_name}'", fields="*").execute()
        return file_contents

    def MakeFolder(self, folder_name):
        body = {"name":folder_name, "mimeType":"application/vnd.google-apps.folder"}
        self.service.files().create(body=body).execute()
        self._ExtractAll(d=self.gdrive_folders,
                         q="mimeType='application/vnd.google-apps.folder'")

    def MoveToFolder(self, name, folder_name):
        combined_dict = {**self.gdrive_folders, **self.gdrive_files}
        file_id = combined_dict[name]
        folder_id = self.gdrive_folders[folder_name]
        file = self.service.files().get(fileId=file_id,
                                        fields='parents').execute()
        previous_parents = ', '.join(file.get('parents'))
        file = self.service.files().update(fileId=file_id,
                                           removeParents=previous_parents,
                                           addParents=folder_id,
                                           fields='id, parents').execute()
        return file
