using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;

namespace PIEHidNetCore
{
    /// <summary>
    ///     PIE Device
    /// </summary>
    public class PIEDevice
    {
        private static readonly ushort[] convertToSplatModeSausages = { 7, 5, 4, 3, 2, 1 };
        private static readonly ushort[] ledSausages = { 7, 3, 1, 6, 4, 2 };
        private bool connected;
        private bool dataThreadActive;
        private Thread dataThreadHandle;

        private int errCodeR;
        private int errCodeReadError;
        private int errCodeW;
        private int errCodeWriteError;
        private bool errorThreadActive;
        private Thread errorThreadHandle;
        private bool holdDataThreadOpen;
        private bool holdErrorThreadOpen;
        private IntPtr readEvent;
        private IntPtr readFileH;
        private SafeFileHandle readFileHandle;
        private RingBuffer readRing;
        private bool readThreadActive;

        private Thread readThreadHandle;
        private PIEDataHandler registeredDataHandler;
        private PIEErrorHandler registeredErrorHandler;
        private FileIOApiDeclarations.SECURITY_ATTRIBUTES securityAttrUnused;
        private IntPtr writeEvent;
        private SafeFileHandle writeFileHandle;
        private RingBuffer writeRing;
        private bool writeThreadActive;
        private Thread writeThreadHandle;

        /// <summary>
        ///     public ctor
        /// </summary>
        /// <param name="path"></param>
        /// <param name="vid"></param>
        /// <param name="pid"></param>
        /// <param name="version"></param>
        /// <param name="hidUsage"></param>
        /// <param name="hidUsagePage"></param>
        /// <param name="readSize"></param>
        /// <param name="writeSize"></param>
        /// <param name="manufacturersString"></param>
        /// <param name="productString"></param>
        public PIEDevice(string path, int vid, int pid, int version, int hidUsage, int hidUsagePage, int readSize,
            int writeSize, string manufacturersString, string productString)
        {
            Path = path;
            Vid = vid;
            Pid = pid;
            Version = version;
            HidUsage = hidUsage;
            HidUsagePage = hidUsagePage;
            ReadLength = readSize;
            WriteLength = writeSize;
            ManufacturersString = manufacturersString;
            ProductString = productString;
            securityAttrUnused.bInheritHandle = 1;
        }

        /// <summary>
        ///     Device Path
        /// </summary>
        public string Path { get; }

        /// <summary>
        ///     Vendor ID
        /// </summary>
        public int Vid { get; }

        /// <summary>
        ///     Product ID
        /// </summary>
        public int Pid { get; }

        /// <summary>
        ///     Version
        /// </summary>
        public int Version { get; }

        /// <summary>
        ///     HID Usage
        /// </summary>
        public int HidUsage { get; }

        /// <summary>
        ///     HID Usage Page
        /// </summary>
        public int HidUsagePage { get; }

        /// <summary>
        ///     Read Buffer Length
        /// </summary>
        public int ReadLength { get; }

        /// <summary>
        ///     Write Buffer Length
        /// </summary>
        public int WriteLength { get; }

        /// <summary>
        ///     Manufacturer Name
        /// </summary>
        public string ManufacturersString { get; }

        /// <summary>
        ///     Product Name
        /// </summary>
        public string ProductString { get; }

        /// <summary>
        ///     Suppresses duplicate/same reports, only reporting changes
        /// </summary>
        public bool suppressDuplicateReports { get; set; }

        /// <summary>
        ///     Completely disables reporting by device
        /// </summary>
        public bool disableReporting { get; set; }

        /// <summary>
        ///     Translating error codes into messages
        /// </summary>
        /// <param name="error"></param>
        /// <returns></returns>
        public static string GetErrorString(int error)
        {
            if (!ErrorMessages.Messages.TryGetValue(error, out var message))
                message = "Unknown Error" + error;
            return message;
        }

        private void ErrorThread()
        {
            while (errorThreadActive)
            {
                if (errCodeReadError != 0)
                {
                    holdDataThreadOpen = true;
                    registeredErrorHandler.HandlePIEHidError(this, errCodeReadError);
                    holdDataThreadOpen = false;
                }

                if (errCodeWriteError != 0)
                {
                    holdErrorThreadOpen = true;
                    registeredErrorHandler.HandlePIEHidError(this, errCodeWriteError);
                    holdErrorThreadOpen = false;
                }

                errCodeReadError = 0;
                errCodeWriteError = 0;
                Thread.Sleep(25);
            }
        }

        /// <summary>
        ///     Write Thread
        /// </summary>
        private void WriteThread()
        {
            //   FileIOApiDeclarations.SECURITY_ATTRIBUTES securityAttrUnused = new FileIOApiDeclarations.SECURITY_ATTRIBUTES();
            var overlapEvent = FileIOApiDeclarations.CreateEvent(ref securityAttrUnused, 1, 0, "");
            var overlapped = new FileIOApiDeclarations.OVERLAPPED
            {
                Offset = 0,
                OffsetHigh = 0,
                hEvent = overlapEvent,
                Internal = IntPtr.Zero,
                InternalHigh = IntPtr.Zero
            };
            if (WriteLength == 0)
                return;

            var buffer = new byte[WriteLength];
            var wgch = GCHandle.Alloc(buffer, GCHandleType.Pinned); //onur March 2009 - pinning is required

            var byteCount = 0;
            ;

            errCodeW = 0;
            errCodeWriteError = 0;
            while (writeThreadActive)
            {
                if (writeRing == null)
                {
                    errCodeW = 407;
                    errCodeWriteError = 407;
                    goto Error;
                }

                while (writeRing.Get(buffer) == 0)
                {
                    if (0 == FileIOApiDeclarations.WriteFile(writeFileHandle, wgch.AddrOfPinnedObject(), WriteLength,
                            ref byteCount, ref overlapped))
                    {
                        var result = Marshal.GetLastWin32Error();
                        if (result != FileIOApiDeclarations.ERROR_IO_PENDING)
                            //if ((result == FileIOApiDeclarations.ERROR_INVALID_HANDLE) ||
                            //    (result == FileIOApiDeclarations.ERROR_DEVICE_NOT_CONNECTED))
                        {
                            if (result == 87)
                            {
                                errCodeW = 412;
                                errCodeWriteError = 412;
                            }
                            else
                            {
                                errCodeW = result;
                                errCodeWriteError = 408;
                            }

                            goto Error;
                        }

                        result = FileIOApiDeclarations.WaitForSingleObject(overlapEvent, 1000);
                        if (result == FileIOApiDeclarations.WAIT_OBJECT_0) goto WriteCompleted;
                        errCodeW = 411;
                        errCodeWriteError = 411;

                        goto Error;
                    }

                    if ((long)byteCount != WriteLength)
                    {
                        errCodeW = 410;
                        errCodeWriteError = 410;
                    }

                    WriteCompleted: ;
                }

                _ = FileIOApiDeclarations.WaitForSingleObject(writeEvent, 100);
                _ = FileIOApiDeclarations.ResetEvent(writeEvent);
            }

            Error:
            wgch.Free(); //onur
        }

        private void ReadThread()
        {
            var overlapEvent = FileIOApiDeclarations.CreateEvent(ref securityAttrUnused, 1, 0, "");
            var overlapped = new FileIOApiDeclarations.OVERLAPPED
            {
                Offset = 0,
                OffsetHigh = 0,
                hEvent = overlapEvent,
                Internal = IntPtr.Zero,
                InternalHigh = IntPtr.Zero
            };
            if (ReadLength == 0)
            {
                errCodeR = 302;
                errCodeReadError = 302;
                return;
            }

            errCodeR = 0;
            errCodeReadError = 0;

            var buffer = new byte[ReadLength];
            var gch = GCHandle.Alloc(buffer, GCHandleType.Pinned); //onur March 2009 - pinning is required

            while (readThreadActive)
            {
                var dataRead = 0; //FileIOApiDeclarations.
                if (readFileHandle.IsInvalid)
                {
                    errCodeReadError = errCodeR = 320;
                    goto EXit;
                }

                if (0 == FileIOApiDeclarations.ReadFile(readFileHandle, gch.AddrOfPinnedObject(), ReadLength,
                        ref dataRead, ref overlapped)) //ref readFileBuffer[0]
                {
                    var result = Marshal.GetLastWin32Error();
                    if (result != FileIOApiDeclarations
                            .ERROR_IO_PENDING) //|| result == FileIOApiDeclarations.ERROR_DEVICE_NOT_CONNECTED)
                    {
                        if (readFileHandle.IsInvalid)
                        {
                            errCodeReadError = errCodeR = 321;
                            goto EXit;
                        }

                        errCodeR = result;
                        errCodeReadError = 308;
                        goto EXit;
                    }

                    // gch.Free(); //onur
                    while (readThreadActive)
                    {
                        result = FileIOApiDeclarations.WaitForSingleObject(overlapEvent, 50);
                        if (FileIOApiDeclarations.WAIT_OBJECT_0 == result)
                        {
                            if (0 == FileIOApiDeclarations.GetOverlappedResult(readFileHandle, ref overlapped,
                                    ref dataRead, 0))
                            {
                                result = Marshal.GetLastWin32Error();
                                if (result == FileIOApiDeclarations.ERROR_INVALID_HANDLE ||
                                    result == FileIOApiDeclarations.ERROR_DEVICE_NOT_CONNECTED)
                                {
                                    errCodeR = 309;
                                    errCodeReadError = 309;
                                    goto EXit;
                                }
                            }

                            // buffer[0] = 89;
                            goto ReadCompleted;
                        }
                    } //while

                    continue;
                }

                //buffer[0] = 90;
                ReadCompleted:
                if (dataRead != ReadLength)
                {
                    errCodeR = 310;
                    errCodeReadError = 310;
                    goto EXit;
                }

                if (suppressDuplicateReports)
                {
                    var r = readRing.TryPutChanged(buffer);
                    if (r == 0)
                        _ = FileIOApiDeclarations.SetEvent(readEvent);
                }
                else
                {
                    readRing.Put(buffer);
                    _ = FileIOApiDeclarations.SetEvent(readEvent);
                }
            } //while

            EXit:
            _ = FileIOApiDeclarations.CancelIo(readFileHandle);
            readFileHandle = null;
            gch.Free();
        }

        private void DataEventThread()
        {
            var currBuff = new byte[ReadLength];

            while (dataThreadActive)
            {
                if (readRing == null)
                    return;
                if (!disableReporting)
                {
                    if (errCodeR != 0)
                    {
                        Array.Clear(currBuff, 0, ReadLength);
                        holdDataThreadOpen = true;
                        registeredDataHandler.HandlePIEHidData(currBuff, this, errCodeR);
                        holdDataThreadOpen = false;
                        dataThreadActive = false;
                    }
                    else if (readRing.Get(currBuff) == 0)
                    {
                        holdDataThreadOpen = true;
                        registeredDataHandler.HandlePIEHidData(currBuff, this, 0);
                        holdDataThreadOpen = false;
                    }

                    if (readRing.IsEmpty())
                        _ = FileIOApiDeclarations.ResetEvent(readEvent);
                }

                // System.Threading.Thread.Sleep(10);
                _ = FileIOApiDeclarations.WaitForSingleObject(readEvent, 100);
            }
        }

        //-----------------------------------------------------------------------------
        /// <summary>
        ///     Sets connection to the enumerated device.
        ///     If inputReportSize greater than zero it generates a read handle.
        ///     If outputReportSize greater than zero it generates a write handle.
        /// </summary>
        /// <returns></returns>
        public int SetupInterface()
        {
            var retin = 0;
            var retout = 0;

            if (connected) return 203;
            if (ReadLength > 0)
            {
                readFileH = FileIOApiDeclarations.CreateFile(Path, FileIOApiDeclarations.GENERIC_READ,
                    FileIOApiDeclarations.FILE_SHARE_READ | FileIOApiDeclarations.FILE_SHARE_WRITE,
                    IntPtr.Zero, FileIOApiDeclarations.OPEN_EXISTING, FileIOApiDeclarations.FILE_FLAG_OVERLAPPED, 0);

                readFileHandle = new SafeFileHandle(readFileH, true);
                if (readFileHandle.IsInvalid)
                {
                    //readEvent = null;
                    //readFileHandle = null;
                    readRing = null;
                    //CloseInterface();
                    retin = 207;
                    goto outputinit;
                }

                readEvent = FileIOApiDeclarations.CreateEvent(ref securityAttrUnused, 1, 0, "");
                readRing = new RingBuffer(128, ReadLength);
                readThreadHandle = new Thread(ReadThread)
                {
                    IsBackground = true,
                    Name = $"PIEHidReadThread for {Pid}"
                };
                readThreadActive = true;
                readThreadHandle.Start();
            }

            outputinit:
            if (WriteLength > 0)
            {
                var writeFileH = FileIOApiDeclarations.CreateFile(Path, FileIOApiDeclarations.GENERIC_WRITE,
                    FileIOApiDeclarations.FILE_SHARE_READ | FileIOApiDeclarations.FILE_SHARE_WRITE,
                    IntPtr.Zero, FileIOApiDeclarations.OPEN_EXISTING,
                    FileIOApiDeclarations.FILE_FLAG_OVERLAPPED,
                    0);
                writeFileHandle = new SafeFileHandle(writeFileH, true);
                if (writeFileHandle.IsInvalid)
                {
                    // writeEvent = null;
                    // writeFileHandle = null;
                    writeRing = null;
                    //CloseInterface();
                    retout = 208;
                    goto ErrorOut;
                }

                writeEvent = FileIOApiDeclarations.CreateEvent(ref securityAttrUnused, 1, 0, "");
                writeRing = new RingBuffer(128, WriteLength);
                writeThreadHandle = new Thread(WriteThread)
                {
                    IsBackground = true,
                    Name = $"PIEHidWriteThread for {Pid}"
                };
                writeThreadActive = true;
                writeThreadHandle.Start();
            }

            connected = true;
            ErrorOut:
            if (retin == 0 && retout == 0)
                return 0;
            if (retin == 207 && retout == 208)
                return 209;
            return retin + retout;
        }

        /// <summary>
        ///     CLoses any open handles and shut down the active interface
        /// </summary>
        public void CloseInterface()
        {
            if (holdErrorThreadOpen || holdDataThreadOpen) return;

            // Shut down event thread
            if (dataThreadActive)
            {
                dataThreadActive = false;
                _ = FileIOApiDeclarations.SetEvent(readEvent);
                var n = 0;
                if (dataThreadHandle != null)
                {
                    while (dataThreadHandle.IsAlive)
                    {
                        Thread.Sleep(10);
                        n++;
                        if (n == 10)
                        {
                            dataThreadHandle.Abort();
                            break;
                        }
                    }

                    dataThreadHandle = null;
                }
            }

            // Shut down read thread
            if (readThreadActive)
            {
                readThreadActive = false;
                // Wait for thread to exit
                if (readThreadHandle != null)
                {
                    var n = 0;
                    while (readThreadHandle.IsAlive)
                    {
                        Thread.Sleep(10);
                        n++;
                        if (n == 10)
                        {
                            readThreadHandle.Abort();
                            break;
                        }
                    }

                    readThreadHandle = null;
                }
            }

            if (writeThreadActive)
            {
                writeThreadActive = false;
                _ = FileIOApiDeclarations.SetEvent(writeEvent);
                if (writeThreadHandle != null)
                {
                    var n = 0;
                    while (writeThreadHandle.IsAlive)
                    {
                        Thread.Sleep(10);
                        n++;
                        if (n == 10)
                        {
                            writeThreadHandle.Abort();
                            break;
                        }
                    }

                    writeThreadHandle = null;
                }
            }

            if (errorThreadActive)
            {
                errorThreadActive = false;
                if (errorThreadHandle != null)
                {
                    var n = 0;
                    while (errorThreadHandle.IsAlive)
                    {
                        Thread.Sleep(10);
                        n++;
                        if (n == 10)
                        {
                            errorThreadHandle.Abort();
                            break;
                        }
                    }

                    errorThreadHandle = null;
                }
            }

            if (writeRing != null) writeRing = null;
            if (readRing != null) readRing = null;

            //  if (readEvent != null) {readEvent = null;}
            //  if (writeEvent != null) { writeEvent = null; }

            if (0x00FF != Pid && 0x00FE != Pid && 0x00FD != Pid && 0x00FC != Pid && 0x00FB != Pid || Version > 272)
            {
                // it's not an old VEC foot pedal (those hang when closing the handle)
                if (readFileHandle !=
                    null) // 9/1/09 - readFileHandle != null ||added by Onur to avoid null reference exception
                    if (!readFileHandle.IsInvalid)
                        readFileHandle.Close();
                if (writeFileHandle != null)
                    if (!writeFileHandle.IsInvalid)
                        writeFileHandle.Close();
            }

            connected = false;
        }

        /// <summary>
        ///     PIEDataHandler Setup
        /// </summary>
        /// <param name="handler"></param>
        /// <returns></returns>
        public int SetDataCallback(PIEDataHandler handler)
        {
            if (!connected)
                return 702;
            if (ReadLength == 0)
                return 703;

            if (registeredDataHandler == null)
            {
                //registeredDataHandler is not defined so define it and create thread. 
                registeredDataHandler = handler;
                dataThreadHandle = new Thread(DataEventThread)
                {
                    IsBackground = true,
                    Name = $"PIEHidEventThread for {Pid}"
                };
                dataThreadActive = true;
                dataThreadHandle.Start();
            }
            else
            {
                return 704; //Only the eventType has been changed.
            }

            return 0;
        }

        /// <summary>
        ///     PIEErrorHandler Setup
        /// </summary>
        /// <param name="handler"></param>
        /// <returns></returns>
        public int SetErrorCallback(PIEErrorHandler handler)
        {
            if (!connected)
                return 802;

            if (registeredErrorHandler == null)
            {
                //registeredErrorHandler is not defined so define it and create thread. 
                registeredErrorHandler = handler;
                errorThreadHandle = new Thread(ErrorThread)
                {
                    IsBackground = true,
                    Name = $"PIEHidErrorThread for {Pid}"
                };
                errorThreadActive = true;
                errorThreadHandle.Start();
            }
            else
            {
                return 804; //Error Handler Already Exists.
            }

            return 0;
        }

        /// <summary>
        ///     Reading last n bytes from buffer
        /// </summary>
        /// <param name="dest"></param>
        /// <returns></returns>
        public int ReadLast(ref byte[] dest)
        {
            if (ReadLength == 0)
                return 502;
            if (!connected)
                return 507;
            if (dest == null)
                dest = new byte[ReadLength];
            if (dest.Length < ReadLength)
                return 503;
            if (readRing.GetLast(dest) != 0)
                return 504;
            return 0;
        }

        /// <summary>
        ///     Reading n bytes from buffer
        /// </summary>
        /// <param name="dest"></param>
        /// <returns></returns>
        public int ReadData(ref byte[] dest)
        {
            if (!connected)
                return 303;
            if (dest == null)
                dest = new byte[ReadLength];
            if (dest.Length < ReadLength)
                return 311;
            if (readRing.Get(dest) != 0)
                return 304;
            return 0;
        }

        /// <summary>
        ///     Blocking Read, waiting for data
        /// </summary>
        /// <param name="dest"></param>
        /// <param name="maxMillis"></param>
        /// <returns></returns>
        public int BlockingReadData(ref byte[] dest, int maxMillis)
        {
            var startTicks = DateTime.UtcNow.Ticks;
            var ret = 304;
            var mills = maxMillis;
            while (mills > 0 && ret == 304)
            {
                if ((ret = ReadData(ref dest)) == 0) break;
                var nowTicks = DateTime.UtcNow.Ticks;
                mills = maxMillis - (int)(nowTicks - startTicks) / 10000;
                Thread.Sleep(10);
            }

            return ret;
        }

        /// <summary>
        ///     Writing to the device
        /// </summary>
        /// <param name="wData"></param>
        /// <returns></returns>
        public int WriteData(byte[] wData)
        {
            if (WriteLength == 0)
                return 402;
            if (!connected)
                return 406;
            if (wData.Length < WriteLength)
                return 403;
            if (writeRing == null)
                return 405;
            if (errCodeW != 0)
                return errCodeW;
            if (writeRing.TryPut(wData) == 3)
            {
                Thread.Sleep(1);
                return 404;
            }

            _ = FileIOApiDeclarations.SetEvent(writeEvent);
            return 0;
        }


        /// <summary>
        ///     Enumerates all valid PIE USB devics.
        /// </summary>
        /// <returns>list of all devices found, ordered by USB port connection</returns>
        public static PIEDevice[] EnumeratePIE()
        {
            return EnumeratePIE(0x05F3);
        }

        /// <summary>
        ///     Enumerates all valid USB devics of the specified Vid.
        /// </summary>
        /// <returns>list of all devices found, ordered by USB port connection</returns>
        public static PIEDevice[] EnumeratePIE(int vid)
        {
            // FileIOApiDeclarations.SECURITY_ATTRIBUTES securityAttrUnusedE = new FileIOApiDeclarations.SECURITY_ATTRIBUTES();  
            var devices = new LinkedList<PIEDevice>();

            // Get all device paths
            var guid = Guid.Empty;
            HidApiDeclarations.HidD_GetHidGuid(ref guid);
            var deviceInfoSet = DeviceManagementApiDeclarations.SetupDiGetClassDevs(ref guid, null, IntPtr.Zero,
                DeviceManagementApiDeclarations.DIGCF_PRESENT
                | DeviceManagementApiDeclarations.DIGCF_DEVICEINTERFACE);

            var deviceInterfaceData = new DeviceManagementApiDeclarations.SP_DEVICE_INTERFACE_DATA();
            deviceInterfaceData.Size = Marshal.SizeOf(deviceInterfaceData); //28;

            var paths = new LinkedList<string>();

            for (var i = 0;
                 0 != DeviceManagementApiDeclarations.SetupDiEnumDeviceInterfaces(deviceInfoSet, 0, ref guid, i,
                     ref deviceInterfaceData);
                 i++)
            {
                var buffSize = 0;
                DeviceManagementApiDeclarations.SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref deviceInterfaceData,
                    IntPtr.Zero, 0, ref buffSize, IntPtr.Zero);
                // Use IntPtr to simulate detail data structure
                var detailBuffer = Marshal.AllocHGlobal(buffSize);

                // sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA) depends on the process bitness,
                // it's 6 with an X86 process (byte packing + 1 char, auto -> unicode -> 4 + 2*1)
                // and 8 with an X64 process (8 bytes packing anyway).
                Marshal.WriteInt32(detailBuffer, Environment.Is64BitProcess ? 8 : 6);

                if (DeviceManagementApiDeclarations.SetupDiGetDeviceInterfaceDetail(deviceInfoSet,
                        ref deviceInterfaceData, detailBuffer, buffSize, ref buffSize, IntPtr.Zero))
                    // convert buffer (starting past the cbsize variable) to string path
                    paths.AddLast(Marshal.PtrToStringAuto(detailBuffer + 4));
            }

            _ = DeviceManagementApiDeclarations.SetupDiDestroyDeviceInfoList(deviceInfoSet);
            //Security attributes not used anymore - not necessary Onur March 2009
            // Open each device file and test for vid
            var securityAttributes = new FileIOApiDeclarations.SECURITY_ATTRIBUTES
            {
                lpSecurityDescriptor = IntPtr.Zero,
                bInheritHandle = Convert.ToInt32(true) //patti keep Int32 here
            };
            securityAttributes.nLength = Marshal.SizeOf(securityAttributes);

            for (var en = paths.GetEnumerator(); en.MoveNext();)
            {
                var path = en.Current;

                var fileH = FileIOApiDeclarations.CreateFile(path, FileIOApiDeclarations.GENERIC_WRITE,
                    FileIOApiDeclarations.FILE_SHARE_READ | FileIOApiDeclarations.FILE_SHARE_WRITE,
                    IntPtr.Zero, FileIOApiDeclarations.OPEN_EXISTING, 0, 0);
                var fileHandle = new SafeFileHandle(fileH, true);
                if (fileHandle.IsInvalid)
                    // Bad handle, try next path
                    continue;
                try
                {
                    var hidAttributes = new HidApiDeclarations.HIDD_ATTRIBUTES();
                    hidAttributes.Size = Marshal.SizeOf(hidAttributes);
                    if (0 != HidApiDeclarations.HidD_GetAttributes(fileHandle, ref hidAttributes)
                        && hidAttributes.VendorID == vid)
                    {
                        // Good attributes and right vid, try to get caps
                        var pPerparsedData = new IntPtr();
                        if (HidApiDeclarations.HidD_GetPreparsedData(fileHandle, ref pPerparsedData))
                        {
                            var hidCaps = new HidApiDeclarations.HIDP_CAPS();
                            if (0 != HidApiDeclarations.HidP_GetCaps(pPerparsedData, ref hidCaps))
                            {
                                // Got Capabilities, add device to list
                                var Mstring = new byte[128];
                                var ssss = "";
                                ;
                                if (0 != HidApiDeclarations.HidD_GetManufacturerString(fileHandle, ref Mstring[0], 128))
                                    for (var i = 0; i < 64; i++)
                                    {
                                        var t = new byte[2];
                                        t[0] = Mstring[2 * i];
                                        t[1] = Mstring[2 * i + 1];
                                        if (t[0] == 0) break;
                                        ssss += Encoding.Unicode.GetString(t);
                                    }

                                var Pstring = new byte[128];
                                var psss = "";
                                if (0 != HidApiDeclarations.HidD_GetProductString(fileHandle, ref Pstring[0], 128))
                                    for (var i = 0; i < 64; i++)
                                    {
                                        var t = new byte[2];
                                        t[0] = Pstring[2 * i];
                                        t[1] = Pstring[2 * i + 1];
                                        if (t[0] == 0) break;
                                        psss += Encoding.Unicode.GetString(t);
                                    }

                                devices.AddLast(new PIEDevice(path, hidAttributes.VendorID, hidAttributes.ProductID,
                                    hidAttributes.VersionNumber, hidCaps.Usage,
                                    hidCaps.UsagePage, hidCaps.InputReportByteLength, hidCaps.OutputReportByteLength,
                                    ssss, psss));
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }
                finally
                {
                    fileHandle.Close();
                }
            }

            var ret = new PIEDevice[devices.Count];
            devices.CopyTo(ret, 0);
            return ret;
        }

        private int SendSausageCommands(ushort[] commandSequence)
        {
            if (WriteLength != 2 || HidUsagePage != 1 || HidUsage != 6) // hid page 1, usage 6 is keyboard
                return 1302;

            var securityAttributes = new FileIOApiDeclarations.SECURITY_ATTRIBUTES
            {
                lpSecurityDescriptor = IntPtr.Zero,
                bInheritHandle = Convert.ToInt32(true)
            };
            securityAttributes.nLength = Marshal.SizeOf(securityAttributes);

            var hF = FileIOApiDeclarations.CreateFile(Path,
                FileIOApiDeclarations.GENERIC_WRITE,
                FileIOApiDeclarations.FILE_SHARE_READ | FileIOApiDeclarations.FILE_SHARE_WRITE,
                IntPtr.Zero,
                FileIOApiDeclarations.OPEN_EXISTING,
                0,
                0);
            var hFile = new SafeFileHandle(hF, true);
            if (hFile.IsInvalid)
            {
                return 1301;
                ;
            }

            var overlapped = new FileIOApiDeclarations.OVERLAPPED
            {
                hEvent = IntPtr.Zero,
                Offset = 0, // IntPtr.Zero;
                OffsetHigh = 0 // IntPtr.Zero;
            };
            foreach (var command in commandSequence)
            {
                var cmd = (uint)command << 16;
                uint bytesReturned = 0;

                if (!DeviceManagementApiDeclarations.DeviceIoControl(hFile, 0x000b0008, ref cmd, 4, IntPtr.Zero, 0,
                        ref bytesReturned, ref overlapped)) return Marshal.GetLastWin32Error();
            }

            hFile.Close();
            return 0;
        }

        /// <summary>
        ///     ???
        /// </summary>
        /// <returns></returns>
        public int ConvertToSplatMode()
        {
            return SendSausageCommands(convertToSplatModeSausages);
        }

        /// <summary>
        ///     ???
        /// </summary>
        /// <returns></returns>
        public int SendLEDSausage()
        {
            return SendSausageCommands(ledSausages);
        }
    }
}