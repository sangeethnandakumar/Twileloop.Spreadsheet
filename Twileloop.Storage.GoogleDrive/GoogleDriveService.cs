using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Upload;

namespace Twileloop.Storage.GoogleDrive
{
    public class GoogleDriveService : IGoogleDriveService
    {
        private readonly DriveService _driveService;

        public GoogleDriveService(string credentialsFilePath, string appName)
        {
            _driveService = InitializeDriveService(credentialsFilePath, appName);
        }

        private DriveService InitializeDriveService(string credentialsFilePath, string appName)
        {
            GoogleCredential credential;
            using (var stream = new FileStream(credentialsFilePath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(DriveService.Scope.Drive);
            }

            return new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = appName,
            });
        }

        public async Task<IEnumerable<GoogleDriveItem>> GetAllFilesAndDirectoriesAsync()
        {
            var request = _driveService.Files.List();
            request.Fields = "files(id, name, parents, mimeType)";

            var result = await request.ExecuteAsync();
            return result.Files.Select(file => new GoogleDriveItem
            {
                Id = file.Id,
                Name = file.Name,
                ParentId = file.Parents?.FirstOrDefault(),
                IsFile = file.MimeType != "application/vnd.google-apps.folder"
            });
        }

        public async Task<GoogleDriveItem> GetCurrentDirectoryAsync()
        {
            var rootFolder = await _driveService.Files.Get("root").ExecuteAsync();
            return new GoogleDriveItem
            {
                Id = rootFolder.Id,
                Name = rootFolder.Name,
                IsFile = false
            };
        }

        public async Task ShareFileWithSpecificUsers(string fileId, List<string> emails)
        {
            foreach (var email in emails)
            {
                var permission = new Permission
                {
                    Type = "user",
                    Role = "writer",
                    EmailAddress = email
                };

                await _driveService.Permissions.Create(permission, fileId).ExecuteAsync();
            }
        }

        public async Task<string> GenerateShareableLink(string fileId)
        {
            var permission = new Permission
            {
                Type = "anyone",
                Role = "reader"
            };

            await _driveService.Permissions.Create(permission, fileId).ExecuteAsync();

            return $"https://drive.google.com/file/d/{fileId}/view";
        }




        public async Task UploadFileAsync(string filePath, string mimeType, string parentId = null, Action<long, long> progress = null, int chunkSizeInMB = 10)
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File
            {
                Name = Path.GetFileName(filePath),
                Parents = new List<string> { parentId }
            };

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                var totalBytes = stream.Length;
                var request = _driveService.Files.Create(fileMetadata, stream, mimeType);
                request.Fields = "id";
                request.ChunkSize = chunkSizeInMB * 1024 * 1024; // Convert MB to bytes
                request.ProgressChanged += (IUploadProgress p) =>
                {
                    progress?.Invoke(totalBytes, p.BytesSent);
                };
                await request.UploadAsync();
            }
        }


        public async Task DownloadFileAsync(string fileId, string destinationPath, Action<long, long> progress = null)
        {
            var request = _driveService.Files.Get(fileId);
            var stream = new MemoryStream();
            request.MediaDownloader.ProgressChanged += (IDownloadProgress p) =>
            {
                progress?.Invoke(p.BytesDownloaded, p.BytesDownloaded);
            };
            await request.DownloadAsync(stream);
            using (var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write))
            {
                stream.WriteTo(fileStream);
            }
        }

        public async Task MoveFileAsync(string fileId, string folderId)
        {
            var file = await _driveService.Files.Get(fileId).ExecuteAsync();
            var previousParents = String.Join(",", file.Parents);
            var request = _driveService.Files.Update(new Google.Apis.Drive.v3.Data.File(), fileId);
            request.AddParents = folderId;
            request.RemoveParents = previousParents;
            file = await request.ExecuteAsync();
        }

        public async Task RenameFileAsync(string fileId, string newFileName)
        {
            var file = new Google.Apis.Drive.v3.Data.File
            {
                Name = newFileName
            };

            await _driveService.Files.Update(file, fileId).ExecuteAsync();
        }

        public async Task DeleteFileAsync(string fileId)
        {
            await _driveService.Files.Delete(fileId).ExecuteAsync();
        }

        public async Task<bool> CheckFilePermissionsAsync(string fileId)
        {
            try
            {
                await _driveService.Permissions.List(fileId).ExecuteAsync();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task CreateDirectoryAsync(string directoryName, string parentDirectoryId = null)
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File
            {
                Name = directoryName,
                MimeType = "application/vnd.google-apps.folder",
                Parents = new List<string> { parentDirectoryId }
            };

            await _driveService.Files.Create(fileMetadata).ExecuteAsync();
        }



        public async Task CopyFileAsync(string fileId, string folderId)
        {
            var copiedFile = new Google.Apis.Drive.v3.Data.File
            {
                Parents = new List<string> { folderId }
            };
            await _driveService.Files.Copy(copiedFile, fileId).ExecuteAsync();
        }

        public async Task<long> GetFileSizeAsync(string fileId)
        {
            var file = await _driveService.Files.Get(fileId).ExecuteAsync();
            return file.Size ?? 0;
        }
    }

}
