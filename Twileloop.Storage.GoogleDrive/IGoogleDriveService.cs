namespace Twileloop.Storage.GoogleDrive
{

    public interface IGoogleDriveService
    {
        Task<IEnumerable<GoogleDriveItem>> GetAllFilesAndDirectoriesAsync();
        Task<GoogleDriveItem> GetCurrentDirectoryAsync();
        Task UploadFileAsync(string filePath, string mimeType, string parentId = null, Action<long, long> progress = null, int chunkSizeInMB = 10);
        Task DownloadFileAsync(string fileId, string destinationPath, Action<long, long> progress = null);
        Task RenameFileAsync(string fileId, string newFileName);
        Task DeleteFileAsync(string fileId);
        Task<bool> CheckFilePermissionsAsync(string fileId);
        Task CreateDirectoryAsync(string directoryName, string parentDirectoryId = null);
        Task MoveFileAsync(string fileId, string folderId);
        Task CopyFileAsync(string fileId, string folderId);
        Task<long> GetFileSizeAsync(string fileId);
        Task ShareFileWithSpecificUsers(string fileId, List<string> emails);
        Task<string> GenerateShareableLink(string fileId);
    }

}
