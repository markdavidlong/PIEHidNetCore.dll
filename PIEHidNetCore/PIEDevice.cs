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
        private bool _connected;
        private bool _dataThreadActive;
        private Thread _dataThreadHandle;

        private int _errCodeR;
        private int _errCodeReadError;
        private int _errCodeW;
        private int _errCodeWriteError;
        private bool _errorThreadActive;
        private Thread _errorThreadHandle;
        private bool _holdDataThreadOpen;
        private bool _holdErrorThreadOpen;
        private IntPtr _readEvent;
        private IntPtr _readFileH;
        private SafeFileHandle _readFileHandle;
        private RingBuffer _readRing;
        private bool _readThreadActive;

        private Thread _readThreadHandle;
        private PIEDataHandler _registeredDataHandler;
        private PIEErrorHandler _registeredErrorHandler;
        private FileIOApiDeclarations.SECURITY_ATTRIBUTES _securityAttrUnused;
        private IntPtr _writeEvent;
        private SafeFileHandle _writeFileHandle;
        private RingBuffer _writeRing;
        private bool _writeThreadActive;
        private Thread _writeThreadHandle;

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
            _securityAttrUnused.bInheritHandle = 1;
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
        public bool SuppressDuplicateReports { get; set; }

        /// <summary>
        ///     Completely disables reporting by device
        /// </summary>
        public bool DisableReporting { get; set; }

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
            while (_errorThreadActive)
            {
                if (_errCodeReadError != 0)
                {
                    _holdDataThreadOpen = true;
                    _registeredErrorHandler.HandlePIEHidError(this, _errCodeReadError);
                    _holdDataThreadOpen = false;
                }

                if (_errCodeWriteError != 0)
                {
                    _holdErrorThreadOpen = true;
                    _registeredErrorHandler.HandlePIEHidError(this, _errCodeWriteError);
                    _holdErrorThreadOpen = false;
                }

                _errCodeReadError = 0;
                _errCodeWriteError = 0;
                Thread.Sleep(25);
            }
        }

        /// <summary>
        ///     Write Thread
        /// </summary>
        private void WriteThread()
        {
            var overlapEvent = FileIOApiDeclarations.CreateEvent(ref _securityAttrUnused, 1, 0, "");
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

            try
            {
                _errCodeW = 0;
                _errCodeWriteError = 0;
                while (_writeThreadActive)
                {
                    if (_writeRing == null)
                    {
                        _errCodeW = 407;
                        _errCodeWriteError = 407;
                        throw new Exception("Write Result = 407");
                    }

                    while (_writeRing.Get(buffer) == 0)
                    {
                        if (0 == FileIOApiDeclarations.WriteFile(_writeFileHandle, wgch.AddrOfPinnedObject(),
                                WriteLength,
                                ref byteCount, ref overlapped))
                        {
                            var result = Marshal.GetLastWin32Error();
                            if (result != FileIOApiDeclarations.ERROR_IO_PENDING)
                                //if ((result == FileIOApiDeclarations.ERROR_INVALID_HANDLE) ||
                                //    (result == FileIOApiDeclarations.ERROR_DEVICE_NOT_CONNECTED))
                            {
                                if (result == 87)
                                {
                                    _errCodeW = 412;
                                    _errCodeWriteError = 412;
                                }
                                else
                                {
                                    _errCodeW = result;
                                    _errCodeWriteError = 408;
                                }

                                throw new Exception("Write Result = 87");
                              
                            }

                            result = FileIOApiDeclarations.WaitForSingleObject(overlapEvent, 1000);
                            if (result == FileIOApiDeclarations.WAIT_OBJECT_0) goto WriteCompleted;
                            _errCodeW = 411;
                            _errCodeWriteError = 411;

                            throw new Exception("Write Result = 411");
                            
                        }
                        
                        if ((long)byteCount != WriteLength)
                        {
                            _errCodeW = 410;
                            _errCodeWriteError = 410;
                        }

                        WriteCompleted: ;
                    }

                    _ = FileIOApiDeclarations.WaitForSingleObject(_writeEvent, 100);
                    _ = FileIOApiDeclarations.ResetEvent(_writeEvent);
                }
            }
                
            catch (Exception e)
            {
                Console.WriteLine("Exception thrown while writing: " + e.ToString());
            }
            finally
            {
                wgch.Free(); //onur
            }
        }

        private void ReadThread()
        {
            var overlapEvent = FileIOApiDeclarations.CreateEvent(ref _securityAttrUnused, 1, 0, "");
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
                _errCodeR = 302;
                _errCodeReadError = 302;
                return;
            }

            _errCodeR = 0;
            _errCodeReadError = 0;

            var buffer = new byte[ReadLength];
            var gch = GCHandle.Alloc(buffer, GCHandleType.Pinned); //onur March 2009 - pinning is required

            try
            {
                while (_readThreadActive)
                {
                    var dataRead = 0; //FileIOApiDeclarations.
                    if (_readFileHandle.IsInvalid)
                    {
                        _errCodeReadError = _errCodeR = 320;
                        throw new Exception("Read error = 320");
                        
                    }

                    if (0 == FileIOApiDeclarations.ReadFile(_readFileHandle, gch.AddrOfPinnedObject(), ReadLength,
                            ref dataRead, ref overlapped)) //ref readFileBuffer[0]
                    {
                        var result = Marshal.GetLastWin32Error();
                        if (result != FileIOApiDeclarations
                                .ERROR_IO_PENDING) //|| result == FileIOApiDeclarations.ERROR_DEVICE_NOT_CONNECTED)
                        {
                            if (_readFileHandle.IsInvalid)
                            {
                                _errCodeReadError = _errCodeR = 321;
                                throw new Exception("Read error = 321");
                                
                            }

                            _errCodeR = result;
                            _errCodeReadError = 308;
                            throw new Exception("Read error = 308");
                            
                        }

                        // gch.Free(); //onur
                        while (_readThreadActive)
                        {
                            result = FileIOApiDeclarations.WaitForSingleObject(overlapEvent, 50);
                            if (FileIOApiDeclarations.WAIT_OBJECT_0 == result)
                            {
                                if (0 == FileIOApiDeclarations.GetOverlappedResult(_readFileHandle, ref overlapped,
                                        ref dataRead, 0))
                                {
                                    result = Marshal.GetLastWin32Error();
                                    if (result == FileIOApiDeclarations.ERROR_INVALID_HANDLE ||
                                        result == FileIOApiDeclarations.ERROR_DEVICE_NOT_CONNECTED)
                                    {
                                        _errCodeR = 309;
                                        _errCodeReadError = 309;
                                        throw new Exception("Read error = 309");
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
                        _errCodeR = 310;
                        _errCodeReadError = 310;
                        throw new Exception("Read Error = 310");
                    }

                    if (SuppressDuplicateReports)
                    {
                        var r = _readRing.TryPutChanged(buffer);
                        if (r == 0)
                            _ = FileIOApiDeclarations.SetEvent(_readEvent);
                    }
                    else
                    {
                        _readRing.Put(buffer);
                        _ = FileIOApiDeclarations.SetEvent(_readEvent);
                    }
                } //while
            }
            catch (Exception e)
            {
                Console.Write("Exception thrown:" + e.ToString());
            }
            finally
            {
                _ = FileIOApiDeclarations.CancelIo(_readFileHandle);
                _readFileHandle = null;
                gch.Free();
            }
        }

        private void DataEventThread()
        {
            var currBuff = new byte[ReadLength];

            while (_dataThreadActive)
            {
                if (_readRing == null)
                    return;
                if (!DisableReporting)
                {
                    if (_errCodeR != 0)
                    {
                        Array.Clear(currBuff, 0, ReadLength);
                        _holdDataThreadOpen = true;
                        _registeredDataHandler.HandlePIEHidData(currBuff, this, _errCodeR);
                        _holdDataThreadOpen = false;
                        _dataThreadActive = false;
                    }
                    else if (_readRing.Get(currBuff) == 0)
                    {
                        _holdDataThreadOpen = true;
                        _registeredDataHandler.HandlePIEHidData(currBuff, this, 0);
                        _holdDataThreadOpen = false;
                    }

                    if (_readRing.IsEmpty())
                        _ = FileIOApiDeclarations.ResetEvent(_readEvent);
                }

                // System.Threading.Thread.Sleep(10);
                _ = FileIOApiDeclarations.WaitForSingleObject(_readEvent, 100);
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

            if (_connected) return 203;
            if (ReadLength > 0)
            {
                _readFileH = FileIOApiDeclarations.CreateFile(Path, FileIOApiDeclarations.GENERIC_READ,
                    FileIOApiDeclarations.FILE_SHARE_READ | FileIOApiDeclarations.FILE_SHARE_WRITE,
                    IntPtr.Zero, FileIOApiDeclarations.OPEN_EXISTING, FileIOApiDeclarations.FILE_FLAG_OVERLAPPED, 0);

                _readFileHandle = new SafeFileHandle(_readFileH, true);
                if (_readFileHandle.IsInvalid)
                {
                    _readRing = null;
                    retin = 207;
                }
                else
                {
                    _readEvent = FileIOApiDeclarations.CreateEvent(ref _securityAttrUnused, 1, 0, "");
                    _readRing = new RingBuffer(128, ReadLength);
                    _readThreadHandle = new Thread(ReadThread)
                    {
                        IsBackground = true,
                        Name = $"PIEHidReadThread for {Pid}"
                    };
                    _readThreadActive = true;
                    _readThreadHandle.Start();
                }
            }


            if (WriteLength > 0)
            {
                var writeFileH = FileIOApiDeclarations.CreateFile(Path, FileIOApiDeclarations.GENERIC_WRITE,
                    FileIOApiDeclarations.FILE_SHARE_READ | FileIOApiDeclarations.FILE_SHARE_WRITE,
                    IntPtr.Zero, FileIOApiDeclarations.OPEN_EXISTING,
                    FileIOApiDeclarations.FILE_FLAG_OVERLAPPED,
                    0);
                _writeFileHandle = new SafeFileHandle(writeFileH, true);
                if (_writeFileHandle.IsInvalid)
                {
                    // writeEvent = null;
                    // writeFileHandle = null;
                    _writeRing = null;
                    //CloseInterface();
                    retout = 208;
                }
                else
                {
                    _writeEvent = FileIOApiDeclarations.CreateEvent(ref _securityAttrUnused, 1, 0, "");
                    _writeRing = new RingBuffer(128, WriteLength);
                    _writeThreadHandle = new Thread(WriteThread)
                    {
                        IsBackground = true,
                        Name = $"PIEHidWriteThread for {Pid}"
                    };
                    _writeThreadActive = true;
                    _writeThreadHandle.Start();
                }
            }


            if (retin == 0 && retout == 0)
            {
                _connected = true;
                return 0;
            }

            if (retin == 207 && retout == 208)
                return 209;
            return retin + retout;
        }

        /// <summary>
        ///     CLoses any open handles and shut down the active interface
        /// </summary>
        public void CloseInterface()
        {
            if (_holdErrorThreadOpen || _holdDataThreadOpen) return;

            // Shut down event thread
            if (_dataThreadActive)
            {
                _dataThreadActive = false;
                _ = FileIOApiDeclarations.SetEvent(_readEvent);
                var n = 0;
                if (_dataThreadHandle != null)
                {
                    while (_dataThreadHandle.IsAlive)
                    {
                        Thread.Sleep(10);
                        n++;
                        if (n == 10)
                        {
                            _dataThreadHandle.Abort();
                            break;
                        }
                    }

                    _dataThreadHandle = null;
                }
            }

            // Shut down read thread
            if (_readThreadActive)
            {
                _readThreadActive = false;
                // Wait for thread to exit
                if (_readThreadHandle != null)
                {
                    var n = 0;
                    while (_readThreadHandle.IsAlive)
                    {
                        Thread.Sleep(10);
                        n++;
                        if (n == 10)
                        {
                            _readThreadHandle.Abort();
                            break;
                        }
                    }

                    _readThreadHandle = null;
                }
            }

            if (_writeThreadActive)
            {
                _writeThreadActive = false;
                _ = FileIOApiDeclarations.SetEvent(_writeEvent);
                if (_writeThreadHandle != null)
                {
                    var n = 0;
                    while (_writeThreadHandle.IsAlive)
                    {
                        Thread.Sleep(10);
                        n++;
                        if (n == 10)
                        {
                            _writeThreadHandle.Abort();
                            break;
                        }
                    }

                    _writeThreadHandle = null;
                }
            }

            if (_errorThreadActive)
            {
                _errorThreadActive = false;
                if (_errorThreadHandle != null)
                {
                    var n = 0;
                    while (_errorThreadHandle.IsAlive)
                    {
                        Thread.Sleep(10);
                        n++;
                        if (n == 10)
                        {
                            _errorThreadHandle.Abort();
                            break;
                        }
                    }

                    _errorThreadHandle = null;
                }
            }

            _writeRing = null;
            _readRing = null;

            //  if (readEvent != null) {readEvent = null;}
            //  if (writeEvent != null) { writeEvent = null; }

            if (0x00FF != Pid && 0x00FE != Pid && 0x00FD != Pid && 0x00FC != Pid && 0x00FB != Pid || Version > 272)
            {
                // it's not an old VEC foot pedal (those hang when closing the handle)
                if (_readFileHandle !=
                    null) // 9/1/09 - readFileHandle != null ||added by Onur to avoid null reference exception
                    if (!_readFileHandle.IsInvalid)
                        _readFileHandle.Close();
                if (_writeFileHandle != null)
                    if (!_writeFileHandle.IsInvalid)
                        _writeFileHandle.Close();
            }

            _connected = false;
        }

        /// <summary>
        ///     PIEDataHandler Setup
        /// </summary>
        /// <param name="handler"></param>
        /// <returns></returns>
        public int SetDataCallback(PIEDataHandler handler)
        {
            if (!_connected)
                return 702;
            if (ReadLength == 0)
                return 703;

            if (_registeredDataHandler == null)
            {
                //registeredDataHandler is not defined so define it and create thread. 
                _registeredDataHandler = handler;
                _dataThreadHandle = new Thread(DataEventThread)
                {
                    IsBackground = true,
                    Name = $"PIEHidEventThread for {Pid}"
                };
                _dataThreadActive = true;
                _dataThreadHandle.Start();
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
            if (!_connected)
                return 802;

            if (_registeredErrorHandler == null)
            {
                //registeredErrorHandler is not defined so define it and create thread. 
                _registeredErrorHandler = handler;
                _errorThreadHandle = new Thread(ErrorThread)
                {
                    IsBackground = true,
                    Name = $"PIEHidErrorThread for {Pid}"
                };
                _errorThreadActive = true;
                _errorThreadHandle.Start();
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
            if (!_connected)
                return 507;
            if (dest == null)
                dest = new byte[ReadLength];
            if (dest.Length < ReadLength)
                return 503;
            if (_readRing.GetLast(dest) != 0)
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
            if (!_connected)
                return 303;
            if (dest == null)
                dest = new byte[ReadLength];
            if (dest.Length < ReadLength)
                return 311;
            if (_readRing.Get(dest) != 0)
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
            if (!_connected)
                return 406;
            if (wData.Length < WriteLength)
                return 403;
            if (_writeRing == null)
                return 405;
            if (_errCodeW != 0)
                return _errCodeW;
            if (_writeRing.TryPut(wData) == 3)
            {
                Thread.Sleep(1);
                return 404;
            }

            _ = FileIOApiDeclarations.SetEvent(_writeEvent);
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
                                var mstring = new byte[128];
                                var ssss = "";
                                ;
                                if (0 != HidApiDeclarations.HidD_GetManufacturerString(fileHandle, ref mstring[0], 128))
                                    for (var i = 0; i < 64; i++)
                                    {
                                        var t = new byte[2];
                                        t[0] = mstring[2 * i];
                                        t[1] = mstring[2 * i + 1];
                                        if (t[0] == 0) break;
                                        ssss += Encoding.Unicode.GetString(t);
                                    }

                                var pstring = new byte[128];
                                var psss = "";
                                if (0 != HidApiDeclarations.HidD_GetProductString(fileHandle, ref pstring[0], 128))
                                    for (var i = 0; i < 64; i++)
                                    {
                                        var t = new byte[2];
                                        t[0] = pstring[2 * i];
                                        t[1] = pstring[2 * i + 1];
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
                catch (Exception e)
                {
                    Console.WriteLine("Exception thrown: " + e.ToString() );
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
    }
}