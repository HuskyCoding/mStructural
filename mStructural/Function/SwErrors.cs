using SolidWorks.Interop.sldworks;
using System;

namespace mStructural.Function
{
    public class SwErrors
    {
        #region Public Variables

        #endregion

        #region Private Variables

        #endregion

        #region enum
        [Flags]
        enum FileLoadError : int
        {
            GenericError = 1,
            FileNotFoundError = 2,
            InvalidFileTypeError = 1024,
            FutureVersion = 8192,
            FileWithSameTitleAlreadyOpen = 65536,
            LiquidMachineDoc = 131072,
            LowResourcesError = 262144,
            NoDisplayData = 524288,
            AddinInteruptError = 1048576,
            FileRequiresRepairError = 2097152,
            FileCriticalDataRepairError = 4194304,
            ApplicationBusy = 8388608,
            ConnectedIsOffline = 16777216
        }

        [Flags]
        enum FileSaveError: int
        {
            GenericSaveError = 1,
            ReadOnlySaveError = 2,
            FileNameEmpty = 4,
            FileNameContainsAtSign = 8,
            FileLockError = 16,
            FileSaveFormatNotAvailable = 32,
            FileSaveAsDoNotOverwrite = 128,
            FileSaveAsInvalidFileExtension = 256,
            FileSaveAsNoSelection = 512,
            FileSaveAsBadEDrawingsVersion = 1024,
            FileSaveAsNameExceedsMaxPathLength = 2048,
            FileSaveAsNotSupported = 4096,
            FileSaveRequiresSavingReferences = 8192
        }

        #endregion

        // Constructor
        public SwErrors()
        {
        }

        // Method to get file load error
        public string GetFileLoadError(int errorCode)
        {
            return ((FileLoadError)errorCode).ToString();
        }

        // Method to get file save error
        public string GetFileSaveError(int errorCode)
        {
            return ((FileSaveError)errorCode).ToString();
        }
    }
}
