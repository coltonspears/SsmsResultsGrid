using System;

namespace SsmsResultsGrid
{
    internal static class PackageGuids
    {
        public const string PackageGuidString = "8f1c9c2c-4f0a-4e9e-9d7b-1e2a4f3b0001";
        public const string CommandSetGuidString = "8f1c9c2c-4f0a-4e9e-9d7b-1e2a4f3b0002";
        public const string ToolWindowGuidString = "8f1c9c2c-4f0a-4e9e-9d7b-1e2a4f3b0003";

        public static readonly Guid CommandSet = new Guid(CommandSetGuidString);

        public const int ShowFilterableGridCmdId = 0x0100;
        public const int ToolsMenuGroupId = 0x1020;
    }
}
