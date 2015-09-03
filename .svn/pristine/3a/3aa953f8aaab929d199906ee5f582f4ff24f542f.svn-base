using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pololu.Usc.ScopeFocus
{
    class StatusHelper
    {
        #region Members
        private IntPtr _handle;
        private string[] _captions;
        private int _panelCount;
        private int _pid;
        #endregion

        #region Constructor
        public StatusHelper(IntPtr hWnd)
        {
            this._handle = hWnd;
            this._panelCount = -1;
            this._pid = -1;
        }
        #endregion

        #region Imports
        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(
            ProcessAccessTypes desiredAccess,
            Boolean inheritHandle,
            Int32 processId
            );

        [DllImport("kernel32.dll")]
        private static extern int CloseHandle(
            IntPtr hObject
            );

        [DllImport("kernel32.dll")]
        private static extern IntPtr VirtualAllocEx(
            IntPtr hProcess,
            IntPtr address,
            UInt32 size,
            VirtualAllocExTypes allocationType,
            AccessProtectionFlags flags
            );

        [DllImport("kernel32.dll")]
        private static extern bool VirtualFreeEx(
            IntPtr hProcess,
            IntPtr address,
            UInt32 size,
            VirtualAllocExTypes dwFreeType
            );

        [DllImport("kernel32.dll")]
        private static extern bool ReadProcessMemory(
            IntPtr hProcess,
            IntPtr baseAddress,
            byte[] buffer,
            UInt32 dwSize,
            out UInt32 numberOfBytesRead
            );

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(
            IntPtr hWnd,
            out Int32 lpdwProcessId
            );

        [DllImport("user32.dll")]
        private static extern uint SendMessage(
            IntPtr hWnd,
            UInt32 wMsg,
            IntPtr wParam,
            IntPtr lParam
            );

        [DllImport("user32.dll")]
        private static extern uint SendMessage(
            IntPtr hWnd,
            UInt32 wMsg,
            UInt32 wParam,
            UInt32 lParam
            );
        #endregion

        #region Constants
        private const uint WM_USER = 0x0400;
        private const uint SB_SETPARTS = WM_USER + 4;
        private const uint SB_GETPARTS = WM_USER + 6;
        private const uint SB_GETTEXTLENGTH = WM_USER + 12;
        private const uint SB_GETTEXT = WM_USER + 13;
        #endregion

        #region Enumerations
        private enum ProcessAccessTypes
        {
            PROCESS_TERMINATE = 0x00000001,
            PROCESS_CREATE_THREAD = 0x00000002,
            PROCESS_SET_SESSIONID = 0x00000004,
            PROCESS_VM_OPERATION = 0x00000008,
            PROCESS_VM_READ = 0x00000010,
            PROCESS_VM_WRITE = 0x00000020,
            PROCESS_DUP_HANDLE = 0x00000040,
            PROCESS_CREATE_PROCESS = 0x00000080,
            PROCESS_SET_QUOTA = 0x00000100,
            PROCESS_SET_INFORMATION = 0x00000200,
            PROCESS_QUERY_INFORMATION = 0x00000400,
            STANDARD_RIGHTS_REQUIRED = 0x000F0000,
            SYNCHRONIZE = 0x00100000,
            PROCESS_ALL_ACCESS = PROCESS_TERMINATE | PROCESS_CREATE_THREAD | PROCESS_SET_SESSIONID | PROCESS_VM_OPERATION |
              PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_DUP_HANDLE | PROCESS_CREATE_PROCESS | PROCESS_SET_QUOTA |
              PROCESS_SET_INFORMATION | PROCESS_QUERY_INFORMATION | STANDARD_RIGHTS_REQUIRED | SYNCHRONIZE
        }

        private enum VirtualAllocExTypes
        {
            WRITE_WATCH_FLAG_RESET = 0x00000001, // Win98 only
            MEM_COMMIT = 0x00001000,
            MEM_RESERVE = 0x00002000,
            MEM_COMMIT_OR_RESERVE = 0x00003000,
            MEM_DECOMMIT = 0x00004000,
            MEM_RELEASE = 0x00008000,
            MEM_FREE = 0x00010000,
            MEM_PUBLIC = 0x00020000,
            MEM_MAPPED = 0x00040000,
            MEM_RESET = 0x00080000, // Win2K only
            MEM_TOP_DOWN = 0x00100000,
            MEM_WRITE_WATCH = 0x00200000, // Win98 only
            MEM_PHYSICAL = 0x00400000, // Win2K only
            //MEM_4MB_PAGES    = 0x80000000, // ??
            SEC_IMAGE = 0x01000000,
            MEM_IMAGE = SEC_IMAGE
        }

        private enum AccessProtectionFlags
        {
            PAGE_NOACCESS = 0x001,
            PAGE_READONLY = 0x002,
            PAGE_READWRITE = 0x004,
            PAGE_WRITECOPY = 0x008,
            PAGE_EXECUTE = 0x010,
            PAGE_EXECUTE_READ = 0x020,
            PAGE_EXECUTE_READWRITE = 0x040,
            PAGE_EXECUTE_WRITECOPY = 0x080,
            PAGE_GUARD = 0x100,
            PAGE_NOCACHE = 0x200,
            PAGE_WRITECOMBINE = 0x400
        }
        #endregion

        #region Private Methods
        private int GetPanelCount()
        {
            if (this._handle != IntPtr.Zero)
                return (int)SendMessage(this._handle, SB_GETPARTS, 0, 0);

            return 0;
        }
        private string[] GetCaptions()
        {
            int count = this.PanelCount;
            string[] captions = new string[count];

            for (uint i = 0; i < count; i++)
                captions[i] = this.GetCaption(i);

            return captions;
        }
        private string GetCaption(uint index)
        {
            uint length = (uint)SendMessage(this._handle, SB_GETTEXTLENGTH, index, 0);

            // Low part is the count. High part is the window type. Mask out the high bits.
            // The returned text will also be unicode so double the length to accomodate our buffer
            length = (length & 0x0000ffff) * 2;

            IntPtr hProcess = IntPtr.Zero;
            IntPtr allocated = IntPtr.Zero;

            try
            {
                hProcess = StatusHelper.OpenProcess(ProcessAccessTypes.PROCESS_ALL_ACCESS, false, this.OwningPID);
                if (hProcess != IntPtr.Zero)
                {
                    // Allocate memory in the remote process
                    allocated = StatusHelper.VirtualAllocEx(hProcess, IntPtr.Zero, length, (VirtualAllocExTypes.MEM_COMMIT_OR_RESERVE), AccessProtectionFlags.PAGE_READWRITE);

                    if (allocated != IntPtr.Zero)
                    {
                        uint bytesRead = 0;
                        byte[] buffer = new byte[length];

                        // SB_GETTEXT tells the remote process to write out text to the remote memory we allocated.
                        StatusHelper.SendMessage(this._handle, SB_GETTEXT, (IntPtr)index, allocated);

                        // Now we need to read that memory from the remote process into a local buffer.
                        bool success = StatusHelper.ReadProcessMemory(hProcess, allocated, buffer, length, out bytesRead);

                        if (success)
                        {
                            // Each char takes 2 bytes.
                            char[] characters = new char[length / 2];

                            for (int i = 0; i < buffer.Length; i = i + 2)
                            {
                                // Even though the second byte will probably always be 0 for en-us let's so a bit shift
                                // then "or" the first and second bytes together before casting to char.
                                uint a = (uint)buffer[i];
                                uint b = (uint)buffer[i + 1] << 8;

                                characters[i / 2] = (char)(a | b);
                            }

                            return new string(characters);
                        }
                    }
                }
            }
            finally
            {
                if (hProcess != IntPtr.Zero)
                {
                    if (allocated != IntPtr.Zero)
                    {

                        // Free the memory in the remote process
                        StatusHelper.VirtualFreeEx(hProcess, allocated, 0, VirtualAllocExTypes.MEM_RELEASE);
                    }

                    // Close the process handle
                    StatusHelper.CloseHandle(hProcess);
                }
            }

            return string.Empty;
        }
        private int GetOwningPid()
        {
            int pid = 0;

            if (this._handle != IntPtr.Zero)
                StatusHelper.GetWindowThreadProcessId(this._handle, out pid);

            return pid;
        }
        #endregion

        #region Public Methods
        public void SetCaptions(int index, string caption)
        {
            if (index == -1)
            {
                string[] oldParts = this.Captions;
                string[] newParts = caption.Split(new string[] { " | " }, StringSplitOptions.None);

                if ((oldParts.Length == newParts.Length) && (newParts.Length > 0))
                {
                    for (int i = 0; i < oldParts.Length; i++)
                    {
                        if (oldParts[i] != newParts[i])
                            this.SetCaption(i, newParts[i]);
                    }
                }
            }
            else
            {
                this.SetCaption(index, caption);
            }
        }
        private void SetCaption(int index, string caption)
        {
            throw new NotImplementedException("Sorry... You'll have to figure out SB_SETTEXT.");
        }
        #endregion

        #region Properties
        public string[] Captions
        {
            get
            {
                if (this._captions == null)
                    this._captions = this.GetCaptions();

                return this._captions;
            }
        }
        public string Caption
        {
            get { return string.Join(" | ", this.Captions); }
            set { this.SetCaptions(-1, value); }
        }
        public int PanelCount
        {
            get
            {
                if (this._panelCount == -1)
                    this._panelCount = this.GetPanelCount();

                return this._panelCount;
            }
        }
        public int OwningPID
        {
            get
            {
                if (this._pid == -1)
                    this._pid = this.GetOwningPid();

                return this._pid;
            }
        }
        #endregion
    }
}

 