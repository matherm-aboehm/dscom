using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Runtime.InteropServices;
using System.Text;

namespace dSPACE.Runtime.InteropServices;

internal sealed unsafe class MetadataOnlyPEImageStream : Stream
{
    [StructLayout(LayoutKind.Sequential)]
    private struct IMAGE_DOS_HEADER
    {
        public ushort e_magic;
        public ushort e_cblp;
        public ushort e_cp;
        public ushort e_crlc;
        public ushort e_cparhdr;
        public ushort e_minalloc;
        public ushort e_maxalloc;
        public ushort e_ss;
        public ushort e_sp;
        public ushort e_csum;
        public ushort e_ip;
        public ushort e_cs;
        public ushort e_lfarlc;
        public ushort e_ovno;
        public fixed ushort e_res[4];
        public ushort e_oemid;
        public ushort e_oeminfo;
        public fixed ushort e_res2[10];
        public int e_lfanew;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IMAGE_FILE_HEADER
    {
        public ushort Machine;
        public ushort NumberOfSections;
        public uint TimeDateStamp;
        public uint PointerToSymbolTable;
        public uint NumberOfSymbols;
        public ushort SizeOfOptionalHeader;
        public ushort Characteristics;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IMAGE_DATA_DIRECTORY
    {
        public uint VirtualAddress;
        public uint Size;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IMAGE_OPTIONAL_HEADER
    {
        public PEMagic Magic;
        public byte MajorLinkerVersion;
        public byte MinorLinkerVersion;
        public uint SizeOfCode;
        public uint SizeOfInitializedData;
        public uint SizeOfUninitializedData;
        public uint AddressOfEntryPoint;
        public uint BaseOfCode;
        public uint BaseOfData;
        private uint _imageBase;
        public ulong ImageBase
        {
            get => Magic == PEMagic.PE32Plus ?
                BaseOfCode + ((ulong)_imageBase << 32) : _imageBase;
            set
            {
                if (Magic == PEMagic.PE32Plus)
                {
                    BaseOfCode = (uint)(value & 0xffffffff);
                    _imageBase = (uint)(value >> 32);
                }
                else
                {
                    Debug.Assert(Magic == PEMagic.PE32);
                    _imageBase = (uint)value;
                }
            }
        }
        public uint SectionAlignment;
        public uint FileAlignment;
        public ushort MajorOperatingSystemVersion;
        public ushort MinorOperatingSystemVersion;
        public ushort MajorImageVersion;
        public ushort MinorImageVersion;
        public ushort MajorSubsystemVersion;
        public ushort MinorSubsystemVersion;
        public uint Win32VersionValue;
        public uint SizeOfImage;
        public uint SizeOfHeaders;
        public uint CheckSum;
        public ushort Subsystem;
        public ushort DllCharacteristics;
        public UIntPtr SizeOfStackReserve;
        public UIntPtr SizeOfStackCommit;
        public UIntPtr SizeOfHeapReserve;
        public UIntPtr SizeOfHeapCommit;
        public uint LoaderFlags;
        public uint NumberOfRvaAndSizes;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 16)]
        public IMAGE_DATA_DIRECTORY[] DataDirectory;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IMAGE_NT_HEADERS
    {
        public uint Signature;
        public IMAGE_FILE_HEADER FileHeader;
        public IMAGE_OPTIONAL_HEADER OptionalHeader;
    }


    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    private struct IMAGE_SECTION_HEADER
    {
        /*[StructLayout(LayoutKind.Explicit)]
        public struct MiscUnion
        {
            [FieldOffset(0)]
            public uint PhysicalAddress;
            [FieldOffset(0)]
            public uint VirtualSize;
        }*/
        private const int NameSize = 8;
        //can't use ByValTStr as it will be zero terminated
        //[MarshalAs(UnmanagedType.ByValTStr, SizeConst = NameSize)]
        private fixed byte _name[NameSize];
        public string Name
        {
            get
            {
                fixed (byte* name = _name)
                {
                    var i = NameSize - 1;
                    for (; i >= 0 && _name[i] == 0; i--)
                    {
                    }
                    return Encoding.UTF8.GetString(name, i + 1);
                }
            }
            set
            {
                var nameBytes = Encoding.UTF8.GetBytes(value);
                if (nameBytes.Length > NameSize)
                {
                    throw new ArgumentException("Max size of Name is 8.", nameof(value));
                }
                fixed (byte* name = _name)
                {
                    Marshal.Copy(nameBytes, 0, new IntPtr(name), nameBytes.Length);
                    for (var i = nameBytes.Length; i < NameSize; i++)
                    {
                        name[i] = 0;
                    }
                }
            }
        }
        //public MiscUnion Misc;
        public uint PhysicalAddressOrVirtualSize;
        public uint VirtualAddress;
        public uint SizeOfRawData;
        public uint PointerToRawData;
        public uint PointerToRelocations;
        public uint PointerToLinenumbers;
        public ushort NumberOfRelocations;
        public ushort NumberOfLinenumbers;
        public uint Characteristics;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IMAGE_COR20_HEADER
    {
        // Header versioning
        public uint cb;
        public ushort MajorRuntimeVersion;
        public ushort MinorRuntimeVersion;
        public IMAGE_DATA_DIRECTORY MetaData;
        public CorFlags Flags;
        /*[StructLayout(LayoutKind.Explicit)]
        public struct EntryPointUnion
        {
            [FieldOffset(0)]
            public uint EntryPointToken;
            [FieldOffset(0)]
            public uint EntryPointRVA;
        }
        public EntryPointUnion EntryPoint;*/
        public uint EntryPointTokenOrRVA;
        public IMAGE_DATA_DIRECTORY Resources;
        public IMAGE_DATA_DIRECTORY StrongNameSignature;
        public IMAGE_DATA_DIRECTORY CodeManagerTable;
        public IMAGE_DATA_DIRECTORY VTableFixups;
        public IMAGE_DATA_DIRECTORY ExportAddressTableJumps;
        public IMAGE_DATA_DIRECTORY ManagedNativeHeader;
    }

    private readonly struct ImagePart<T>
    {
        private readonly byte[]? _data;
        private readonly byte* _dataPtr;
        public ImagePart([DisallowNull] T data, long startOffset)
        {
            StartOffset = startOffset;
            Size = Marshal.SizeOf<T>();
            _data = new byte[Size];
            _dataPtr = null;
            fixed (byte* blob = _data)
            {
                Marshal.StructureToPtr(data, new IntPtr(blob), true);
            };
        }
        public ImagePart(byte* data, int size, long startOffset)
        {
            StartOffset = startOffset;
            Size = size;
            _data = null;
            _dataPtr = data;

        }
        public long StartOffset { get; }
        public int Size { get; }

        public bool CanReadFromPart(long pos, int count)
        {
            return (pos + count) > StartOffset && pos < (StartOffset + Size);
        }

        public int Read(ref long pos, byte[] buffer, ref int offset, ref int count)
        {
            var offsetInPart = (int)(pos - StartOffset);
            var bytesToFill = 0;
            if (offsetInPart < 0)
            {
                bytesToFill = -offsetInPart;
                offsetInPart = 0;
                Array.Clear(buffer, offset, bytesToFill);
                pos += bytesToFill;
                offset += bytesToFill;
                count -= bytesToFill;
            }
            var bytesToCopy = Math.Min(count, Size - offsetInPart);
            if (_data is not null)
            {
                Buffer.BlockCopy(_data, offsetInPart, buffer, offset, bytesToCopy);
            }
            else
            {
                var ptr = new IntPtr(_dataPtr + offsetInPart);
                Marshal.Copy(ptr, buffer, offset, bytesToCopy);
            }
            pos += bytesToCopy;
            offset += bytesToCopy;
            count -= bytesToCopy;
            return bytesToFill + bytesToCopy;
        }
    }

    private const ushort DosSignature = 0x5A4D;     // 'M' 'Z'
    private const int PESignatureOffsetLocation = 0x3C;
    private const uint PESignature = 0x00004550;    // PE00
    private const int PESignatureSize = sizeof(uint);
    private const int IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR = 14;
    private const int DefaultSectionAlignment = 0x10000;
    private const int DefaultFileAlignment = 0x200;

    private byte* _metadataBytes;
    private int _metadataSize;
    private long _pos;
    private readonly long _imageFileSize;
    private readonly ImagePart<IMAGE_DOS_HEADER> _dosHeaderPart;
    private readonly ImagePart<IMAGE_NT_HEADERS> _ntHeaderPart;
    private readonly ImagePart<IMAGE_SECTION_HEADER> _corSecHeaderPart;
    private readonly ImagePart<IMAGE_COR20_HEADER> _corHeaderPart;
    private readonly ImagePart<byte> _metadataPart;

    public MetadataOnlyPEImageStream(Assembly assembly)
    {
        assembly.ManifestModule.GetPEKind(out var peKind, out var machine);
        assembly.TryGetRawMetadata(out _metadataBytes, out _metadataSize);
        var is64bitRuntime = IntPtr.Size == 8;
        Debug.Assert((peKind & PortableExecutableKinds.ILOnly) != 0 ||
            is64bitRuntime == ((peKind & PortableExecutableKinds.PE32Plus) != 0));

        var headerSize = AlignValue((uint)(sizeof(IMAGE_DOS_HEADER) + Marshal.SizeOf<IMAGE_NT_HEADERS>()
            + Marshal.SizeOf<IMAGE_SECTION_HEADER>()), DefaultFileAlignment);
        var corHeaderSize = (uint)Marshal.SizeOf<IMAGE_COR20_HEADER>();
        _imageFileSize = headerSize + corHeaderSize + _metadataSize;
        var sectionStart = AlignValue(headerSize, DefaultSectionAlignment);
        var sectionSize = AlignValue(corHeaderSize + (uint)_metadataSize, DefaultSectionAlignment);
        var imageSize = sectionStart + sectionSize;
        var dosHeader = new IMAGE_DOS_HEADER
        {
            e_magic = DosSignature,
            e_lfanew = sizeof(IMAGE_DOS_HEADER)
        };
        _dosHeaderPart = new(dosHeader, 0);
        var ntHeader = new IMAGE_NT_HEADERS
        {
            Signature = PESignature,
            FileHeader = new()
            {
                Machine = (ushort)machine,
                NumberOfSections = 1,
                SizeOfOptionalHeader = (ushort)Marshal.SizeOf<IMAGE_OPTIONAL_HEADER>(),
                Characteristics = (ushort)Characteristics.Dll
            },
            OptionalHeader = new()
            {
                //ignore peKind here, use actual platform type
                Magic = is64bitRuntime ? PEMagic.PE32Plus : PEMagic.PE32,
                SectionAlignment = DefaultSectionAlignment,
                FileAlignment = DefaultFileAlignment,
                SizeOfImage = imageSize,
                SizeOfHeaders = headerSize,
                NumberOfRvaAndSizes = IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR + 1,
                DataDirectory = new IMAGE_DATA_DIRECTORY[16]
            }
        };
        ref var CorHeaderTableDirectory = ref ntHeader.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR];
        CorHeaderTableDirectory.VirtualAddress = sectionStart;
        CorHeaderTableDirectory.Size = corHeaderSize;
        _ntHeaderPart = new(ntHeader, dosHeader.e_lfanew);
        var corSecHeader = new IMAGE_SECTION_HEADER
        {
            Name = ".cormeta",
            PhysicalAddressOrVirtualSize = sectionSize,
            VirtualAddress = sectionStart,
            SizeOfRawData = (uint)(corHeaderSize + _metadataSize),
            PointerToRawData = headerSize,
        };
        _corSecHeaderPart = new(corSecHeader, _ntHeaderPart.Size + dosHeader.e_lfanew);
        //var runtimeVersion = Version.Parse(assembly.ImageRuntimeVersion.TrimStart('v'));
        var corHeader = new IMAGE_COR20_HEADER
        {
            cb = corHeaderSize,
            MajorRuntimeVersion = 2,
            MinorRuntimeVersion = 5,
            Flags = ConvertPEKindToCorFlags(peKind, assembly.GetName()),
            MetaData = new()
            {
                VirtualAddress = sectionStart + corHeaderSize,
                Size = (uint)_metadataSize
            }
        };
        _corHeaderPart = new(corHeader, headerSize);
        _metadataPart = new(_metadataBytes, _metadataSize, headerSize + corHeaderSize);
    }

    private static uint AlignValue(uint value, uint alignment)
    {
        return (value + (alignment - 1)) & ~(alignment - 1);
    }

    private static CorFlags ConvertPEKindToCorFlags(PortableExecutableKinds peKind, AssemblyName assemblyName)
    {
        CorFlags flags = 0;
        if ((peKind & PortableExecutableKinds.ILOnly) != 0)
        {
            flags |= CorFlags.ILOnly;
        }
        if ((peKind & PortableExecutableKinds.Required32Bit) != 0)
        {
            flags |= CorFlags.Requires32Bit;
        }
        if ((peKind & PortableExecutableKinds.Preferred32Bit) != 0)
        {
            flags |= CorFlags.Prefers32Bit;
        }
        if ((assemblyName.Flags & AssemblyNameFlags.PublicKey) != 0)
        {
            flags |= CorFlags.StrongNameSigned;
        }
        return flags;
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);

        _metadataBytes = null;
        _metadataSize = 0;
    }

    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => false;
    public override long Length => _imageFileSize;
    public override long Position { get => _pos; set => Seek(value, SeekOrigin.Begin); }

    public override int Read(byte[] buffer, int offset, int count)
    {
        if (_metadataBytes is null)
        {
            throw new ObjectDisposedException(nameof(MetadataOnlyPEImageStream));
        }
        if (buffer is null)
        {
            throw new ArgumentNullException(nameof(buffer));
        }
        if (offset < 0 || count < 0)
        {
            throw new ArgumentOutOfRangeException(offset < 0 ? nameof(offset) : nameof(count));
        }
        if (buffer.Length < offset + count)
        {
            throw new ArgumentException("Can not read behind the buffer length", nameof(buffer));
        }

        var dataread = 0;
        if (_dosHeaderPart.CanReadFromPart(_pos, count))
        {
            dataread += _dosHeaderPart.Read(ref _pos, buffer, ref offset, ref count);
        }
        if (_ntHeaderPart.CanReadFromPart(_pos, count))
        {
            dataread += _ntHeaderPart.Read(ref _pos, buffer, ref offset, ref count);
        }
        if (_corSecHeaderPart.CanReadFromPart(_pos, count))
        {
            dataread += _corSecHeaderPart.Read(ref _pos, buffer, ref offset, ref count);
        }
        if (_corHeaderPart.CanReadFromPart(_pos, count))
        {
            dataread += _corHeaderPart.Read(ref _pos, buffer, ref offset, ref count);
        }
        if (_metadataPart.CanReadFromPart(_pos, count))
        {
            dataread += _metadataPart.Read(ref _pos, buffer, ref offset, ref count);
        }
        if (count > 0 && dataread == 0)
        {
            throw new IOException();
        }
        return dataread;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        if (_metadataBytes is null)
        {
            throw new ObjectDisposedException(nameof(MetadataOnlyPEImageStream));
        }
        var newpos = origin switch
        {
            SeekOrigin.Begin => offset,
            SeekOrigin.Current => _pos + offset,
            SeekOrigin.End => _imageFileSize - offset,
            _ => throw new ArgumentOutOfRangeException(nameof(origin))
        };
        if (newpos < 0)
        {
            throw new ArgumentException("Seeking before the beginning of the stream", nameof(offset));
        }
        if (newpos > _imageFileSize)
        {
            throw new EndOfStreamException("Seeking past the end of the stream");
        }
        _pos = newpos;
        return newpos;
    }

    public override void Flush() => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}
