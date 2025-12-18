package xls

import (
	"encoding/binary"
	"io"
)

// CFB (Compound File Binary) / OLE2 container implementation for XLS (BIFF8) files

const (
	cfbHeaderSize     = 512
	cfbSectorSize     = 512
	cfbMiniSectorSize = 64
	cfbDIFATSize      = 109
	cfbMaxRegSector   = 0xFFFFFFFA
	cfbFATSector      = 0xFFFFFFFD
	cfbEndOfChain     = 0xFFFFFFFE
	cfbFreeSector     = 0xFFFFFFFF
)

// CFBHeader represents the CFB file header
type CFBHeader struct {
	Signature          [8]byte
	CLSID              [16]byte
	MinorVersion       uint16
	MajorVersion       uint16
	ByteOrder          uint16
	SectorShift        uint16
	MiniSectorShift    uint16
	Reserved           [6]byte
	TotalSectors       uint32
	FATSectors         uint32
	FirstDirSector     uint32
	TransactionSig     uint32
	MiniStreamCutoff   uint32
	FirstMiniFATSector uint32
	MiniFATSectors     uint32
	FirstDIFATSector   uint32
	DIFATSectors       uint32
	DIFAT              [109]uint32
}

// NewCFBHeader creates a new CFB header
func NewCFBHeader() *CFBHeader {
	h := &CFBHeader{
		Signature:          [8]byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1},
		MinorVersion:       0x003E,
		MajorVersion:       0x0003,
		ByteOrder:          0xFFFE,
		SectorShift:        0x0009,
		MiniSectorShift:    0x0006,
		MiniStreamCutoff:   0x00001000,
		FirstMiniFATSector: cfbEndOfChain,
		FirstDIFATSector:   cfbEndOfChain,
	}

	for i := range h.DIFAT {
		h.DIFAT[i] = cfbFreeSector
	}

	return h
}

// WriteTo writes the header to the writer
func (h *CFBHeader) WriteTo(w io.Writer) error {
	buf := make([]byte, cfbHeaderSize)

	copy(buf[0:8], h.Signature[:])
	copy(buf[8:24], h.CLSID[:])
	binary.LittleEndian.PutUint16(buf[24:26], h.MinorVersion)
	binary.LittleEndian.PutUint16(buf[26:28], h.MajorVersion)
	binary.LittleEndian.PutUint16(buf[28:30], h.ByteOrder)
	binary.LittleEndian.PutUint16(buf[30:32], h.SectorShift)
	binary.LittleEndian.PutUint16(buf[32:34], h.MiniSectorShift)
	copy(buf[34:40], h.Reserved[:])
	binary.LittleEndian.PutUint32(buf[40:44], h.TotalSectors)
	binary.LittleEndian.PutUint32(buf[44:48], h.FATSectors)
	binary.LittleEndian.PutUint32(buf[48:52], h.FirstDirSector)
	binary.LittleEndian.PutUint32(buf[52:56], h.TransactionSig)
	binary.LittleEndian.PutUint32(buf[56:60], h.MiniStreamCutoff)
	binary.LittleEndian.PutUint32(buf[60:64], h.FirstMiniFATSector)
	binary.LittleEndian.PutUint32(buf[64:68], h.MiniFATSectors)
	binary.LittleEndian.PutUint32(buf[68:72], h.FirstDIFATSector)
	binary.LittleEndian.PutUint32(buf[72:76], h.DIFATSectors)

	for i, v := range h.DIFAT {
		binary.LittleEndian.PutUint32(buf[76+i*4:80+i*4], v)
	}

	_, err := w.Write(buf)
	return err
}

// CFBDirectoryEntry represents a directory entry
type CFBDirectoryEntry struct {
	Name            [64]byte
	NameLength      uint16
	ObjectType      byte
	ColorFlag       byte
	LeftSiblingDID  uint32
	RightSiblingDID uint32
	ChildDID        uint32
	CLSID           [16]byte
	StateBits       uint32
	CreationTime    uint64
	ModifiedTime    uint64
	StartSector     uint32
	StreamSize      uint64
}

// WriteTo writes the directory entry to the writer
func (e *CFBDirectoryEntry) WriteTo(w io.Writer) error {
	buf := make([]byte, 128)

	copy(buf[0:64], e.Name[:])
	binary.LittleEndian.PutUint16(buf[64:66], e.NameLength)
	buf[66] = e.ObjectType
	buf[67] = e.ColorFlag
	binary.LittleEndian.PutUint32(buf[68:72], e.LeftSiblingDID)
	binary.LittleEndian.PutUint32(buf[72:76], e.RightSiblingDID)
	binary.LittleEndian.PutUint32(buf[76:80], e.ChildDID)
	copy(buf[80:96], e.CLSID[:])
	binary.LittleEndian.PutUint32(buf[96:100], e.StateBits)
	binary.LittleEndian.PutUint64(buf[100:108], e.CreationTime)
	binary.LittleEndian.PutUint64(buf[108:116], e.ModifiedTime)
	binary.LittleEndian.PutUint32(buf[116:120], e.StartSector)
	binary.LittleEndian.PutUint64(buf[120:128], e.StreamSize)

	_, err := w.Write(buf)
	return err
}

// stringToUTF16LE converts a string to UTF-16LE
func stringToUTF16LE(s string) []byte {
	runes := []rune(s)
	buf := make([]byte, len(runes)*2)
	for i, r := range runes {
		binary.LittleEndian.PutUint16(buf[i*2:], uint16(r))
	}
	return buf
}

// WriteCFB wraps BIFF8 data in a CFB container and writes it to the writer
func WriteCFB(w io.Writer, workbookData []byte) error {
	// Set minimum size to 4096 bytes to avoid Mini Stream requirement
	dataSize := len(workbookData)
	if dataSize < 4096 {
		dataSize = 4096
	}
	dataSectors := (dataSize + cfbSectorSize - 1) / cfbSectorSize

	// Sector layout:
	// Sector 0-(dataSectors-1): Data
	// Sector dataSectors: FAT
	// Sector dataSectors+1: Directory
	fatSector := dataSectors
	dirSector := dataSectors + 1

	header := NewCFBHeader()
	header.FATSectors = 1
	header.FirstDirSector = uint32(dirSector)
	header.DIFAT[0] = uint32(fatSector)

	if err := header.WriteTo(w); err != nil {
		return err
	}

	paddedData := make([]byte, dataSectors*cfbSectorSize)
	copy(paddedData, workbookData)
	if _, err := w.Write(paddedData); err != nil {
		return err
	}

	// Write FAT (File Allocation Table)
	fat := make([]uint32, cfbSectorSize/4)
	for i := range fat {
		fat[i] = cfbFreeSector
	}

	for i := 0; i < dataSectors; i++ {
		if i == dataSectors-1 {
			fat[i] = cfbEndOfChain
		} else {
			fat[i] = uint32(i + 1)
		}
	}

	fat[fatSector] = cfbFATSector
	fat[dirSector] = cfbEndOfChain

	fatBuf := make([]byte, cfbSectorSize)
	for i, v := range fat {
		binary.LittleEndian.PutUint32(fatBuf[i*4:], v)
	}
	if _, err := w.Write(fatBuf); err != nil {
		return err
	}

	// Write Directory
	dirBuf := make([]byte, cfbSectorSize)

	rootName := stringToUTF16LE("Root Entry")
	root := &CFBDirectoryEntry{
		NameLength:      uint16(len(rootName) + 2),
		ObjectType:      5,
		ColorFlag:       1,
		LeftSiblingDID:  cfbFreeSector,
		RightSiblingDID: cfbFreeSector,
		ChildDID:        1,
		StartSector:     cfbEndOfChain,
		StreamSize:      0,
	}
	copy(root.Name[:], rootName)
	root.Name[len(rootName)] = 0
	root.Name[len(rootName)+1] = 0

	wbName := stringToUTF16LE("Workbook")
	workbook := &CFBDirectoryEntry{
		NameLength:      uint16(len(wbName) + 2),
		ObjectType:      2,
		ColorFlag:       1,
		LeftSiblingDID:  cfbFreeSector,
		RightSiblingDID: cfbFreeSector,
		ChildDID:        cfbFreeSector,
		StartSector:     0,
		StreamSize:      uint64(dataSize),
	}
	copy(workbook.Name[:], wbName)
	workbook.Name[len(wbName)] = 0
	workbook.Name[len(wbName)+1] = 0

	empty := &CFBDirectoryEntry{
		ObjectType:      0,
		LeftSiblingDID:  cfbFreeSector,
		RightSiblingDID: cfbFreeSector,
		ChildDID:        cfbFreeSector,
		StartSector:     cfbEndOfChain,
	}

	tmpBuf := make([]byte, 128)

	root.WriteTo(&bufferWriter{buf: dirBuf[0:128]})
	workbook.WriteTo(&bufferWriter{buf: dirBuf[128:256]})
	empty.WriteTo(&bufferWriter{buf: tmpBuf})
	copy(dirBuf[256:384], tmpBuf)
	empty.WriteTo(&bufferWriter{buf: tmpBuf})
	copy(dirBuf[384:512], tmpBuf)

	if _, err := w.Write(dirBuf); err != nil {
		return err
	}

	return nil
}

// bufferWriter writes to a fixed-size buffer
type bufferWriter struct {
	buf []byte
	pos int
}

func (bw *bufferWriter) Write(p []byte) (n int, err error) {
	n = copy(bw.buf[bw.pos:], p)
	bw.pos += n
	return n, nil
}
