import argparse
import logging
import os
import re
from urllib.parse import urlparse

from sharepy.session import connect


logger = logging.getLogger("sharepy.upload")


class SharePointUploadError(Exception):
    """Error uploading to SharePoint"""


class SharePointUpload:
    """Work with SharePoint files and folders"""
    # https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest

    def __init__(self, site_url, session):
        self.site_url = site_url
        if not self.site_url.endswith("/"):
            self.site_url += "/"
        parsed_url = urlparse(self.site_url)
        self.base_path = parsed_url.path
        self.session = session

    @classmethod
    def create(cls, site_url, username, password):
        """Create a SharePointUpload for site"""
        logger.info("Connecting to SharePoint url=%r, username=%r", site_url, username)
        session = connect(site_url, username=username, password=password)
        return cls(site_url, session)

    def folder_exists(self, folder):
        """Check if the folder exists"""
        resp = self.session.get(self.site_url + "_api/web/GetFolderByServerRelativeUrl('{}')".format(folder))
        logger.debug(
            "Folder exists check response: folder=%r, code=%s, body=%r",
            folder,
            resp.status_code,
            resp.text,
        )
        if resp.status_code in (401, 403):
            logger.error(
                "Access to folder %r is denied, code=%s", folder, resp.status_code
            )
            raise SharePointUploadError("Access to folder denied")
        if resp.status_code == 404:
            if resp.json()["error"]["code"] != "-2147024894, System.IO.FileNotFoundException":
                raise AssertionError("SharePoint not found error code does not match!")
            return False
        if resp.status_code != 200:
            raise AssertionError(
                f"SharePoint returned a unexpected response code {resp.status_code}"
            )
        return True

    def file_exists(self, path):
        """Check if a file already exists"""
        # You don't need base_path in GetFolderByServerRelativeUrl, but it seems you do
        #  in GetFileByServerRelativeUrl - thanks for consistency Microsoft
        resp = self.session.get(
            self.site_url + "_api/web/GetFileByServerRelativeUrl('{}')".format(
                self.base_path + path
            )
        )
        logger.debug(
            "File exists check response: file=%r, code=%s, body=%r",
            path,
            resp.status_code,
            resp.text,
        )
        if resp.status_code == 200:
            return True
        return False

    def file_checked_out(self, path):
        """Get user file is checked out to, or None"""
        resp = self.session.get(
            self.site_url + "_api/web/GetFileByServerRelativeUrl('{}')/CheckedOutByUser".format(
                self.base_path + path
            )
        )
        logger.debug("File checked out query response: file=%r code=%s, json=%r",
            path,
            resp.status_code,
            resp.json(),
        )
        if "LoginName" in resp.json()["d"]:
            return re.sub(r"^.*\|", "", resp.json()["d"]["LoginName"])
        return None

    def checkout_file(self, path):
        """Checkout a file to logged in user"""
        resp = self.session.post(
            self.site_url + "_api/web/GetFileByServerRelativeUrl('{}')/Checkout()".format(
                self.base_path + path
            )
        )
        logger.debug(
            "File checkout query response: file=%r code=%s, json=%r",
            path,
            resp.status_code,
            resp.json(),
        )
        if resp.status_code != 200:
            raise SharePointUploadError(f"File {path!r} could not be checked out")

    def checkin_file(self, path):
        """Checkin a file"""
        resp = self.session.post(
            self.site_url + "_api/web/GetFileByServerRelativeUrl('{}')/CheckIn(comment='Upload from SharePy',checkintype=0)".format(
                    self.base_path + path
                )
            )
        logger.debug(
            "File checkin query response: file=%r code=%s, json=%r",
            path,
            resp.status_code,
            resp.json(),
        )
        if resp.status_code != 200:
            raise SharePointUploadError(f"File {path!r} could not be checked in")

    def upload_file(self, source, dest_folder, dest_file, exists=False):
        """Upload a file"""
        logger.info(
            "Uploading file %r to '%s/%s', exists=%s",
            source,
            dest_folder,
            dest_file,
            exists,
        )
        with open(source, "rb") as source_fp:
            if exists:
                resp = self.session.post(
                    self.site_url + "_api/web/GetFileByServerRelativeUrl('{}/{}')/$value".format(
                        self.base_path + dest_folder, dest_file
                    ),
                    data=source_fp,
                    headers={"X-HTTP-Method": "PUT"},
                )
            else:
                resp = self.session.post(
                    self.site_url + "_api/web/GetFolderByServerRelativeUrl('{}')/Files/add(url='{}',overwrite=true)".format(
                        dest_folder, dest_file
                    ),
                    data=source_fp,
                )
        logger.debug(
            "File upload response: file='%s/%s' code=%s, body=%r",
            dest_folder,
            dest_file,
            resp.status_code,
            resp.text,
        )
        if resp.status_code not in (200, 204):
            if "error" in resp.json():
                error = "\n" + resp.json()["error"]["message"]["value"]
            else:
                error = ""
            raise SharePointUploadError(
                f"File upload failed! status={resp.status_code} error={error!r}"
            )

    def _create_folder(self, path):
        """Create a folder (parent must exist)"""
        data = {
            "__metadata": {
                "type": "SP.Folder",
            },
            "ServerRelativeUrl": self.base_path + path,
        }
        resp = self.session.post(self.site_url + "_api/web/folders", json=data)
        logger.debug(
            "Create folder response: folder=%r code=%s, json=%r",
            path,
            resp.status_code,
            resp.json(),
        )
        if resp.status_code != 201:
            raise SharePointUploadError(f"Folder {path!r} could not be created")

    def create_folder(self, path):
        """Create folder including parents if required"""
        # Split path into folders to create each one separately
        logger.info("Creating folder(s) for path: %r", path)
        folders = path.split("/")
        for i in range(len(folders) + 1):
            self._create_folder("/".join(folders[:i]))

    def upload(self, source, dest):
        logger.info("Starting checks to upload file: %s", source)
        exists = False
        dest_folder, dest_file = os.path.split(dest)
        # Check folder exists on server
        if not self.folder_exists(dest_folder):
            self.create_folder(dest_folder)

        # Check if file already exists
        if self.file_exists(dest):
            exists = True
            logger.info("File %r already exists", dest)
            # Check file is checked out to current user
            if user := self.file_checked_out(dest):
                logger.info(
                    "File already checked out to user: %r, currently logged in as %r",
                    user,
                    self.session.auth.username,
                )
                if user.lower() != self.session.auth.username.lower():
                    raise SharePointUploadError(f"File already checked out to '{user}'")
            else:
                self.checkout_file(dest)

        try:
            self.upload_file(source, dest_folder, dest_file, exists=exists)
        finally:
            if exists:
                self.checkin_file(dest)
        logger.info("Upload complete!")


def main():
    parser = argparse.ArgumentParser(description="Upload a file to a SharePoint site")
    parser.add_argument("site_url", help="URL to sharepoint site")
    parser.add_argument("-u", dest="username", help="Login username/email")
    parser.add_argument("-p", dest="password", help="Login password")
    parser.add_argument("src", help="Source path to upload")
    parser.add_argument("dest", help="Destination path in SharePoint")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO)

    uploader = SharePointUpload.create(args.site_url, args.username, args.password)
    uploader.upload(args.src, args.dest)


if __name__ == "__main__":
    main()
