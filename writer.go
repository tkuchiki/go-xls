package xls

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"math"
	"os"

	"golang.org/x/text/encoding/unicode"
)

// BIFF8 record types
const (
	recTypeBOF        = 0x0809
	recTypeEOF        = 0x000A
	recTypeDIMENSIONS = 0x0200
	recTypeROW        = 0x0208
	recTypeLABEL      = 0x0204
	recTypeNUMBER     = 0x0203
	recTypeBOOLERR    = 0x0205
	recTypeSST        = 0x00FC
	recTypeEXTSST     = 0x00FF
	recTypeLABELSST   = 0x00FD
	recTypeCODEPAGE   = 0x0042
	recTypeFONT       = 0x0031
	recTypeFORMAT     = 0x041E
	recTypeXF         = 0x00E0
	recTypeSTYLE      = 0x0293
	recTypeBOUNDSHEET       = 0x0085
	recTypeWINDOW1          = 0x003D
	recTypeWINDOW2          = 0x023E
	recTypeDEFAULTROWHEIGHT = 0x0225
	recTypeDEFCOLWIDTH      = 0x0055
	recTypeWSBOOL           = 0x0081
	recTypeBOOKBOOL         = 0x00DA

	recTypeINTERFACEHDR   = 0x00E1
	recTypeMMS            = 0x00C1
	recTypeINTERFACEEND   = 0x00E2
	recTypeWRITEACCESS    = 0x005C
	recTypeDATEMODE       = 0x0022
	recTypePRECISION      = 0x000E
	recTypeREFRESHALL     = 0x01B7
	recTypeCALCMODE       = 0x000D
	recTypeCALCCOUNT      = 0x000C
	recTypeREFMODE        = 0x000F
	recTypeITERATION      = 0x0011
	recTypeDELTA          = 0x0010
	recTypeSAVERECALC     = 0x005F
	recTypePRINTHEADERS   = 0x002A
	recTypePRINTGRIDLINES = 0x002B
	recTypePROTECT        = 0x0012
	recTypePASSWORD       = 0x0013
	recTypeBACKUP         = 0x0040
	recTypeHIDEOBJ        = 0x008D
	recTypeWINDOWPROTECT  = 0x0019
	recTypeDSF            = 0x0161
	recTypePROT4REV       = 0x01AF
	recTypePASSWORDREV4   = 0x01BC
	recTypeFNGROUPCOUNT   = 0x013D
	recTypeUSESELFS       = 0x0160
	recTypeUNKNOWN9C      = 0x009C

	recTypeLEFTMARGIN   = 0x0026
	recTypeRIGHTMARGIN  = 0x0027
	recTypeTOPMARGIN    = 0x0028
	recTypeBOTTOMMARGIN = 0x0029
	recTypeHCENTER      = 0x0083
	recTypeVCENTER      = 0x0084
	recTypeSETUP        = 0x00A1
	recTypeGRIDSET      = 0x0082
	recTypeGUTS         = 0x0080
	recTypeOBJPROTECT   = 0x0063
	recTypeSCENPROTECT  = 0x00DD
	recTypeHBREAK       = 0x001B
	recTypeVBREAK       = 0x001A
	recTypeHEADER       = 0x0014
	recTypeFOOTER       = 0x0015
)

const (
	biffVersion  = 0x0600 // BIFF8
	bofWorkbook  = 0x0005 // Workbook globals
	bofWorksheet = 0x0010 // Worksheet
)

// Writer writes Excel XLS files in BIFF8 format.
type Writer struct {
	data      [][]interface{}
	sheetName string
}

// New creates a new Writer.
func New() *Writer {
	return &Writer{
		sheetName: "Sheet1",
	}
}

// SetSheetName sets the sheet name.
func (w *Writer) SetSheetName(name string) {
	w.sheetName = name
}

// Write sets the data to be written.
func (w *Writer) Write(data [][]interface{}) error {
	w.data = data
	return nil
}

// SaveAs writes the XLS file to the specified path.
func (w *Writer) SaveAs(filename string) error {
	buf := new(bytes.Buffer)
	if err := w.writeBIFF8(buf); err != nil {
		return fmt.Errorf("failed to write BIFF8 data: %w", err)
	}

	file, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	if err := WriteCFB(file, buf.Bytes()); err != nil {
		return fmt.Errorf("failed to write CFB container: %w", err)
	}

	return nil
}

func (w *Writer) writeBIFF8(buf *bytes.Buffer) error {
	// Build Shared String Table (SST)
	sst := newSST()
	for _, row := range w.data {
		for _, cell := range row {
			if str, ok := cell.(string); ok {
				sst.addString(str)
			}
		}
	}

	// BOF (Workbook Globals)
	if err := w.writeBOF(buf, bofWorkbook); err != nil {
		return err
	}

	if err := w.writeInterfaceHdr(buf); err != nil {
		return err
	}

	if err := w.writeMMS(buf); err != nil {
		return err
	}

	if err := w.writeInterfaceEnd(buf); err != nil {
		return err
	}

	if err := w.writeWriteAccess(buf); err != nil {
		return err
	}

	if err := w.writeCodePage(buf); err != nil {
		return err
	}

	if err := w.writeDSF(buf); err != nil {
		return err
	}

	if err := w.writeFnGroupCount(buf); err != nil {
		return err
	}

	if err := w.writeUnknown9C(buf); err != nil {
		return err
	}

	if err := w.writeWindowProtect(buf); err != nil {
		return err
	}

	if err := w.writeProtect(buf); err != nil {
		return err
	}

	if err := w.writeObjProtect(buf); err != nil {
		return err
	}

	if err := w.writePassword(buf); err != nil {
		return err
	}

	if err := w.writeProt4Rev(buf); err != nil {
		return err
	}

	if err := w.writePasswordRev4(buf); err != nil {
		return err
	}

	if err := w.writeBackup(buf); err != nil {
		return err
	}

	if err := w.writeHideObj(buf); err != nil {
		return err
	}

	if err := w.writeWindow1(buf); err != nil {
		return err
	}

	if err := w.writeDateMode(buf); err != nil {
		return err
	}

	if err := w.writePrecision(buf); err != nil {
		return err
	}

	if err := w.writeRefreshAll(buf); err != nil {
		return err
	}

	if err := w.writeBookBool(buf); err != nil {
		return err
	}

	// BIFF8 requires 7 default font records
	for i := 0; i < 7; i++ {
		if err := w.writeDefaultFont(buf); err != nil {
			return err
		}
	}

	if err := w.writeFormat(buf); err != nil {
		return err
	}

	// First 16 XF records are style XF
	for i := 0; i < 16; i++ {
		if err := w.writeXF(buf, true, 6); err != nil {
			return err
		}
	}
	// Cell XF records
	if err := w.writeXF(buf, false, 6); err != nil {
		return err
	}
	if err := w.writeXF(buf, false, 7); err != nil {
		return err
	}

	if err := w.writeDefaultStyle(buf); err != nil {
		return err
	}

	if err := w.writeUseSelfs(buf); err != nil {
		return err
	}

	// Calculate worksheet offset for BOUNDSHEET record
	sstBuf := new(bytes.Buffer)
	if err := w.writeSST(sstBuf, sst); err != nil {
		return err
	}

	sheetNameBytes := stringToUTF16LE(w.sheetName)
	boundsheetSize := 4 + 6 + 1 + len(sheetNameBytes) + 1

	worksheetOffset := buf.Len() + sstBuf.Len() + boundsheetSize + 4 // +4 for EOF

	if _, err := buf.Write(sstBuf.Bytes()); err != nil {
		return err
	}

	if err := w.writeBoundSheet(buf, uint32(worksheetOffset), w.sheetName); err != nil {
		return err
	}

	if err := w.writeEOF(buf); err != nil {
		return err
	}

	// BOF (Worksheet)
	if err := w.writeBOF(buf, bofWorksheet); err != nil {
		return err
	}

	if err := w.writeCalcMode(buf); err != nil {
		return err
	}
	if err := w.writeCalcCount(buf); err != nil {
		return err
	}
	if err := w.writeRefMode(buf); err != nil {
		return err
	}
	if err := w.writeIteration(buf); err != nil {
		return err
	}
	if err := w.writeDelta(buf); err != nil {
		return err
	}
	if err := w.writeSaveRecalc(buf); err != nil {
		return err
	}

	if err := w.writeGuts(buf); err != nil {
		return err
	}

	if err := w.writeDefaultRowHeight(buf); err != nil {
		return err
	}

	if err := w.writeWSBool(buf); err != nil {
		return err
	}

	// DIMENSIONS must come before ROW records
	if err := w.writeDimensions(buf); err != nil {
		return err
	}

	if err := w.writePrintHeaders(buf); err != nil {
		return err
	}
	if err := w.writePrintGridlines(buf); err != nil {
		return err
	}
	if err := w.writeGridSet(buf); err != nil {
		return err
	}
	if err := w.writeHBreak(buf); err != nil {
		return err
	}
	if err := w.writeVBreak(buf); err != nil {
		return err
	}
	if err := w.writeHeader(buf); err != nil {
		return err
	}
	if err := w.writeFooter(buf); err != nil {
		return err
	}
	if err := w.writeHCenter(buf); err != nil {
		return err
	}
	if err := w.writeVCenter(buf); err != nil {
		return err
	}
	if err := w.writeLeftMargin(buf); err != nil {
		return err
	}
	if err := w.writeRightMargin(buf); err != nil {
		return err
	}
	if err := w.writeTopMargin(buf); err != nil {
		return err
	}
	if err := w.writeBottomMargin(buf); err != nil {
		return err
	}
	if err := w.writeSetup(buf); err != nil {
		return err
	}

	if err := w.writeProtect(buf); err != nil {
		return err
	}
	if err := w.writeScenProtect(buf); err != nil {
		return err
	}
	if err := w.writeWindowProtect(buf); err != nil {
		return err
	}
	if err := w.writeObjProtect(buf); err != nil {
		return err
	}
	if err := w.writePassword(buf); err != nil {
		return err
	}

	if err := w.writeRowsAndCells(buf, sst); err != nil {
		return err
	}

	// WINDOW2 must come after cell data
	if err := w.writeWindow2(buf); err != nil {
		return err
	}

	if err := w.writeEOF(buf); err != nil {
		return err
	}

	return nil
}

// Close releases resources.
func (w *Writer) Close() error {
	return nil
}

func (w *Writer) writeBOF(writer io.Writer, subType uint16) error {
	data := make([]byte, 16)
	binary.LittleEndian.PutUint16(data[0:2], biffVersion)
	binary.LittleEndian.PutUint16(data[2:4], subType)
	binary.LittleEndian.PutUint16(data[4:6], 0x0DBB) // Build identifier (Excel 2000)
	binary.LittleEndian.PutUint16(data[6:8], 0x07CC) // Build year (1996)
	binary.LittleEndian.PutUint32(data[8:12], 0x00000000)
	binary.LittleEndian.PutUint32(data[12:16], 0x00000006) // Lowest BIFF version
	return w.writeRecord(writer, recTypeBOF, data)
}

func (w *Writer) writeEOF(writer io.Writer) error {
	return w.writeRecord(writer, recTypeEOF, []byte{})
}

func (w *Writer) writeCodePage(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0x04B0) // UTF-16LE (1200)
	return w.writeRecord(writer, recTypeCODEPAGE, data)
}

func (w *Writer) writeDefaultFont(writer io.Writer) error {
	fontName := "Arial"

	// FONT record uses compressed string (8-bit)
	data := make([]byte, 14+1+1+len(fontName))
	binary.LittleEndian.PutUint16(data[0:2], 200) // Height (200 = 10pt)
	binary.LittleEndian.PutUint16(data[2:4], 0)
	binary.LittleEndian.PutUint16(data[4:6], 0x7FFF) // Color index
	binary.LittleEndian.PutUint16(data[6:8], 400) // Weight
	binary.LittleEndian.PutUint16(data[8:10], 0)
	data[10] = 0
	data[11] = 0
	data[12] = 1 // Character set (1 = default)
	data[13] = 0
	data[14] = byte(len(fontName))
	data[15] = 0x00 // Compressed string (8-bit)
	copy(data[16:], []byte(fontName))

	return w.writeRecord(writer, recTypeFONT, data)
}

func (w *Writer) writeFormat(writer io.Writer) error {
	formatString := "General"

	data := make([]byte, 2+2+1+len(formatString))
	binary.LittleEndian.PutUint16(data[0:2], 0x00A4) // Format index (164 = user-defined)
	binary.LittleEndian.PutUint16(data[2:4], uint16(len(formatString)))
	data[4] = 0x00 // Compressed string (8-bit)
	copy(data[5:], []byte(formatString))

	return w.writeRecord(writer, recTypeFORMAT, data)
}

func (w *Writer) writeXF(writer io.Writer, isStyleXF bool, fontIndex uint16) error {
	data := make([]byte, 20)

	if isStyleXF {
		binary.LittleEndian.PutUint16(data[0:2], fontIndex)
		binary.LittleEndian.PutUint16(data[2:4], 0x00A4) // Format index (164 = General)
		binary.LittleEndian.PutUint16(data[4:6], 0xFFF5) // Style XF flag
		binary.LittleEndian.PutUint16(data[6:8], 0x0020)
		binary.LittleEndian.PutUint32(data[8:12], 0x0000F400)
		binary.LittleEndian.PutUint32(data[12:16], 0x00000000)
		binary.LittleEndian.PutUint32(data[16:20], 0x20C00000)
	} else {
		binary.LittleEndian.PutUint16(data[0:2], fontIndex)
		binary.LittleEndian.PutUint16(data[2:4], 0x00A4)
		binary.LittleEndian.PutUint16(data[4:6], 0x0001) // Parent style XF (XF #0)
		binary.LittleEndian.PutUint16(data[6:8], 0x0020)
		binary.LittleEndian.PutUint32(data[8:12], 0x0000F800)
		binary.LittleEndian.PutUint32(data[12:16], 0x00000000)
		binary.LittleEndian.PutUint32(data[16:20], 0x20C00000)
	}

	return w.writeRecord(writer, recTypeXF, data)
}

func (w *Writer) writeDefaultStyle(writer io.Writer) error {
	data := make([]byte, 4)
	binary.LittleEndian.PutUint16(data[0:2], 0x8000) // Built-in style
	data[2] = 0
	data[3] = 0xFF
	return w.writeRecord(writer, recTypeSTYLE, data)
}

func (w *Writer) writeWindow1(writer io.Writer) error {
	data := make([]byte, 18)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	binary.LittleEndian.PutUint16(data[2:4], 0)
	binary.LittleEndian.PutUint16(data[4:6], 0x4000)
	binary.LittleEndian.PutUint16(data[6:8], 0x3000)
	binary.LittleEndian.PutUint16(data[8:10], 0x0038)
	binary.LittleEndian.PutUint16(data[10:12], 0)
	binary.LittleEndian.PutUint16(data[12:14], 0)
	binary.LittleEndian.PutUint16(data[14:16], 1)
	binary.LittleEndian.PutUint16(data[16:18], 600)
	return w.writeRecord(writer, recTypeWINDOW1, data)
}

func (w *Writer) writeWindow2(writer io.Writer) error {
	data := make([]byte, 18)
	binary.LittleEndian.PutUint16(data[0:2], 0x06B6)
	binary.LittleEndian.PutUint16(data[2:4], 0)
	binary.LittleEndian.PutUint16(data[4:6], 0)
	binary.LittleEndian.PutUint16(data[6:8], 0x0040)
	binary.LittleEndian.PutUint16(data[8:10], 0)
	binary.LittleEndian.PutUint16(data[10:12], 0)
	binary.LittleEndian.PutUint32(data[12:16], 0)
	binary.LittleEndian.PutUint16(data[16:18], 0)
	return w.writeRecord(writer, recTypeWINDOW2, data)
}

func (w *Writer) writeDefColWidth(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 8)
	return w.writeRecord(writer, recTypeDEFCOLWIDTH, data)
}

func (w *Writer) writeDefaultRowHeight(writer io.Writer) error {
	data := make([]byte, 4)
	binary.LittleEndian.PutUint16(data[0:2], 0x0000)
	binary.LittleEndian.PutUint16(data[2:4], 0x00FF) // 1/20 point units (255 = 12.75pt)
	return w.writeRecord(writer, recTypeDEFAULTROWHEIGHT, data)
}

func (w *Writer) writeWSBool(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0x04C1)
	return w.writeRecord(writer, recTypeWSBOOL, data)
}

func (w *Writer) writeBookBool(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeBOOKBOOL, data)
}

func (w *Writer) writeInterfaceHdr(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0x04B0)
	return w.writeRecord(writer, recTypeINTERFACEHDR, data)
}

func (w *Writer) writeMMS(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeMMS, data)
}

func (w *Writer) writeInterfaceEnd(writer io.Writer) error {
	return w.writeRecord(writer, recTypeINTERFACEEND, []byte{})
}

func (w *Writer) writeWriteAccess(writer io.Writer) error {
	// Fixed length: 112 bytes, space-padded
	data := make([]byte, 112)
	username := "Go XLS Writer"
	copy(data, []byte(username))
	for i := len(username); i < 112; i++ {
		data[i] = 0x20
	}
	return w.writeRecord(writer, recTypeWRITEACCESS, data)
}

func (w *Writer) writeDateMode(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0) // 0 = 1900 date system
	return w.writeRecord(writer, recTypeDATEMODE, data)
}

func (w *Writer) writePrecision(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 1) // 1 = calculate with displayed precision
	return w.writeRecord(writer, recTypePRECISION, data)
}

func (w *Writer) writeRefreshAll(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeREFRESHALL, data)
}

func (w *Writer) writeCalcMode(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 1) // 1 = automatic calculation
	return w.writeRecord(writer, recTypeCALCMODE, data)
}

func (w *Writer) writeCalcCount(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 100) // Default iteration count
	return w.writeRecord(writer, recTypeCALCCOUNT, data)
}

func (w *Writer) writeRefMode(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 1) // 1 = A1 reference style
	return w.writeRecord(writer, recTypeREFMODE, data)
}

func (w *Writer) writeIteration(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0) // 0 = iteration off
	return w.writeRecord(writer, recTypeITERATION, data)
}

func (w *Writer) writeDelta(writer io.Writer) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data[0:8], math.Float64bits(0.001))
	return w.writeRecord(writer, recTypeDELTA, data)
}

func (w *Writer) writeSaveRecalc(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 1) // 1 = recalculate on save
	return w.writeRecord(writer, recTypeSAVERECALC, data)
}

func (w *Writer) writePrintHeaders(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypePRINTHEADERS, data)
}

func (w *Writer) writePrintGridlines(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypePRINTGRIDLINES, data)
}

func (w *Writer) writeProtect(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypePROTECT, data)
}

func (w *Writer) writePassword(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypePASSWORD, data)
}

func (w *Writer) writeBackup(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeBACKUP, data)
}

func (w *Writer) writeHideObj(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeHIDEOBJ, data)
}

func (w *Writer) writeWindowProtect(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeWINDOWPROTECT, data)
}

func (w *Writer) writeDSF(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0) // 0 = single stream file
	return w.writeRecord(writer, recTypeDSF, data)
}

func (w *Writer) writeFnGroupCount(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0x0001)
	return w.writeRecord(writer, recTypeFNGROUPCOUNT, data)
}

func (w *Writer) writeUnknown9C(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0x000E)
	return w.writeRecord(writer, recTypeUNKNOWN9C, data)
}

func (w *Writer) writeUseSelfs(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 1) // 1 = use natural language formulas
	return w.writeRecord(writer, recTypeUSESELFS, data)
}

func (w *Writer) writeProt4Rev(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypePROT4REV, data)
}

func (w *Writer) writePasswordRev4(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypePASSWORDREV4, data)
}

func (w *Writer) writeLeftMargin(writer io.Writer) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data[0:8], math.Float64bits(0.75)) // 0.75 inches
	return w.writeRecord(writer, recTypeLEFTMARGIN, data)
}

func (w *Writer) writeRightMargin(writer io.Writer) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data[0:8], math.Float64bits(0.75)) // 0.75 inches
	return w.writeRecord(writer, recTypeRIGHTMARGIN, data)
}

func (w *Writer) writeTopMargin(writer io.Writer) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data[0:8], math.Float64bits(1.0)) // 1.0 inches
	return w.writeRecord(writer, recTypeTOPMARGIN, data)
}

func (w *Writer) writeBottomMargin(writer io.Writer) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint64(data[0:8], math.Float64bits(1.0)) // 1.0 inches
	return w.writeRecord(writer, recTypeBOTTOMMARGIN, data)
}

func (w *Writer) writeHCenter(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeHCENTER, data)
}

func (w *Writer) writeVCenter(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeVCENTER, data)
}

func (w *Writer) writeSetup(writer io.Writer) error {
	data := make([]byte, 34)
	binary.LittleEndian.PutUint16(data[0:2], 1)
	binary.LittleEndian.PutUint16(data[2:4], 100)
	binary.LittleEndian.PutUint16(data[4:6], 1)
	binary.LittleEndian.PutUint16(data[6:8], 1)
	binary.LittleEndian.PutUint16(data[8:10], 1)
	binary.LittleEndian.PutUint16(data[10:12], 0x0000)
	binary.LittleEndian.PutUint16(data[12:14], 600)
	binary.LittleEndian.PutUint16(data[14:16], 600)
	binary.LittleEndian.PutUint16(data[16:18], 1)
	return w.writeRecord(writer, recTypeSETUP, data)
}

func (w *Writer) writeGridSet(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 1)
	return w.writeRecord(writer, recTypeGRIDSET, data)
}

func (w *Writer) writeGuts(writer io.Writer) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	binary.LittleEndian.PutUint16(data[2:4], 0)
	binary.LittleEndian.PutUint16(data[4:6], 0)
	binary.LittleEndian.PutUint16(data[6:8], 0)
	return w.writeRecord(writer, recTypeGUTS, data)
}

func (w *Writer) writeObjProtect(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeOBJPROTECT, data)
}

func (w *Writer) writeScenProtect(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeSCENPROTECT, data)
}

func (w *Writer) writeHBreak(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeHBREAK, data)
}

func (w *Writer) writeVBreak(writer io.Writer) error {
	data := make([]byte, 2)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	return w.writeRecord(writer, recTypeVBREAK, data)
}

func (w *Writer) writeHeader(writer io.Writer) error {
	data := make([]byte, 5)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	data[2] = 0x00
	data[3] = 0x00
	data[4] = 0x00
	return w.writeRecord(writer, recTypeHEADER, data)
}

func (w *Writer) writeFooter(writer io.Writer) error {
	data := make([]byte, 5)
	binary.LittleEndian.PutUint16(data[0:2], 0)
	data[2] = 0x00
	data[3] = 0x00
	data[4] = 0x00
	return w.writeRecord(writer, recTypeFOOTER, data)
}

func (w *Writer) writeBoundSheet(writer io.Writer, offset uint32, sheetName string) error {
	nameBytes := stringToUTF16LE(sheetName)
	nameLen := len([]rune(sheetName))

	data := make([]byte, 6+1+1+len(nameBytes))
	binary.LittleEndian.PutUint32(data[0:4], offset)
	data[4] = 0
	data[5] = 0
	data[6] = byte(nameLen) // Character count
	data[7] = 0x01 // Unicode flag (UTF-16LE)
	copy(data[8:], nameBytes)

	return w.writeRecord(writer, recTypeBOUNDSHEET, data)
}

func (w *Writer) writeDimensions(writer io.Writer) error {
	rowCount := uint32(len(w.data))
	colCount := uint16(0)
	for _, row := range w.data {
		if uint16(len(row)) > colCount {
			colCount = uint16(len(row))
		}
	}

	data := make([]byte, 14)
	binary.LittleEndian.PutUint32(data[0:4], 0)
	binary.LittleEndian.PutUint32(data[4:8], rowCount) // Last row + 1
	binary.LittleEndian.PutUint16(data[8:10], 0)
	binary.LittleEndian.PutUint16(data[10:12], colCount) // Last column + 1
	binary.LittleEndian.PutUint16(data[12:14], 0)

	return w.writeRecord(writer, recTypeDIMENSIONS, data)
}

func (w *Writer) writeRowsAndCells(writer io.Writer, sst *sharedStringTable) error {
	for rowIndex, row := range w.data {
		if err := w.writeRow(writer, uint16(rowIndex), uint16(len(row))); err != nil {
			return err
		}

		for colIndex, cell := range row {
			if err := w.writeCell(writer, uint16(rowIndex), uint16(colIndex), cell, sst); err != nil {
				return err
			}
		}
	}
	return nil
}

func (w *Writer) writeRow(writer io.Writer, rowIndex, colCount uint16) error {
	data := make([]byte, 16)
	binary.LittleEndian.PutUint16(data[0:2], rowIndex)
	binary.LittleEndian.PutUint16(data[2:4], 0)
	binary.LittleEndian.PutUint16(data[4:6], colCount) // Last defined column + 1
	binary.LittleEndian.PutUint16(data[6:8], 0x00FF)
	binary.LittleEndian.PutUint16(data[8:10], 0)
	binary.LittleEndian.PutUint16(data[10:12], 0)
	binary.LittleEndian.PutUint32(data[12:16], 0x000F0000)

	return w.writeRecord(writer, recTypeROW, data)
}

func (w *Writer) writeCell(writer io.Writer, row, col uint16, value interface{}, sst *sharedStringTable) error {
	switch v := value.(type) {
	case string:
		return w.writeLabelSST(writer, row, col, v, sst)
	case int:
		return w.writeNumber(writer, row, col, float64(v))
	case int8:
		return w.writeNumber(writer, row, col, float64(v))
	case int16:
		return w.writeNumber(writer, row, col, float64(v))
	case int32:
		return w.writeNumber(writer, row, col, float64(v))
	case int64:
		return w.writeNumber(writer, row, col, float64(v))
	case uint:
		return w.writeNumber(writer, row, col, float64(v))
	case uint8:
		return w.writeNumber(writer, row, col, float64(v))
	case uint16:
		return w.writeNumber(writer, row, col, float64(v))
	case uint32:
		return w.writeNumber(writer, row, col, float64(v))
	case uint64:
		return w.writeNumber(writer, row, col, float64(v))
	case float32:
		return w.writeNumber(writer, row, col, float64(v))
	case float64:
		return w.writeNumber(writer, row, col, v)
	case bool:
		return w.writeBool(writer, row, col, v)
	default:
		return w.writeLabelSST(writer, row, col, fmt.Sprintf("%v", v), sst)
	}
}

func (w *Writer) writeLabelSST(writer io.Writer, row, col uint16, value string, sst *sharedStringTable) error {
	sstIndex := sst.getIndex(value)

	data := make([]byte, 10)
	binary.LittleEndian.PutUint16(data[0:2], row)
	binary.LittleEndian.PutUint16(data[2:4], col)
	binary.LittleEndian.PutUint16(data[4:6], 0)
	binary.LittleEndian.PutUint32(data[6:10], uint32(sstIndex))

	return w.writeRecord(writer, recTypeLABELSST, data)
}

func (w *Writer) writeNumber(writer io.Writer, row, col uint16, value float64) error {
	data := make([]byte, 14)
	binary.LittleEndian.PutUint16(data[0:2], row)
	binary.LittleEndian.PutUint16(data[2:4], col)
	binary.LittleEndian.PutUint16(data[4:6], 0)
	binary.LittleEndian.PutUint64(data[6:14], math.Float64bits(value))

	return w.writeRecord(writer, recTypeNUMBER, data)
}

func (w *Writer) writeBool(writer io.Writer, row, col uint16, value bool) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint16(data[0:2], row)
	binary.LittleEndian.PutUint16(data[2:4], col)
	binary.LittleEndian.PutUint16(data[4:6], 0)
	if value {
		data[6] = 1
	} else {
		data[6] = 0
	}
	data[7] = 0 // Not an error

	return w.writeRecord(writer, recTypeBOOLERR, data)
}

func (w *Writer) writeSST(writer io.Writer, sst *sharedStringTable) error {
	data := make([]byte, 8)
	binary.LittleEndian.PutUint32(data[0:4], uint32(sst.totalCount))
	binary.LittleEndian.PutUint32(data[4:8], uint32(sst.uniqueCount))

	for _, str := range sst.strings {
		strData, err := encodeStringForSST(str)
		if err != nil {
			return err
		}
		data = append(data, strData...)
	}

	return w.writeRecord(writer, recTypeSST, data)
}

func (w *Writer) writeRecord(writer io.Writer, recType uint16, data []byte) error {
	header := make([]byte, 4)
	binary.LittleEndian.PutUint16(header[0:2], recType)
	binary.LittleEndian.PutUint16(header[2:4], uint16(len(data)))

	if _, err := writer.Write(header); err != nil {
		return err
	}
	if len(data) > 0 {
		if _, err := writer.Write(data); err != nil {
			return err
		}
	}
	return nil
}

// sharedStringTable manages the Shared String Table.
type sharedStringTable struct {
	strings     []string
	stringMap   map[string]int
	uniqueCount int
	totalCount  int
}

func newSST() *sharedStringTable {
	return &sharedStringTable{
		strings:   make([]string, 0),
		stringMap: make(map[string]int),
	}
}

func (sst *sharedStringTable) addString(s string) {
	sst.totalCount++
	if _, exists := sst.stringMap[s]; !exists {
		sst.stringMap[s] = sst.uniqueCount
		sst.strings = append(sst.strings, s)
		sst.uniqueCount++
	}
}

func (sst *sharedStringTable) getIndex(s string) int {
	return sst.stringMap[s]
}

// encodeString encodes a string in BIFF8 format (length + flag + UTF-16LE).
func encodeString(s string) ([]byte, error) {
	encoder := unicode.UTF16(unicode.LittleEndian, unicode.IgnoreBOM).NewEncoder()
	utf16, err := encoder.String(s)
	if err != nil {
		return nil, err
	}

	result := make([]byte, 3+len(utf16))
	result[0] = byte(len([]rune(s))) // Character count (not byte count)
	result[1] = 0x01                 // Unicode flag (UTF-16LE)
	result[2] = 0x00
	copy(result[3:], utf16)

	return result, nil
}

// encodeStringForSST encodes a string for the SST record.
func encodeStringForSST(s string) ([]byte, error) {
	encoder := unicode.UTF16(unicode.LittleEndian, unicode.IgnoreBOM).NewEncoder()
	utf16, err := encoder.String(s)
	if err != nil {
		return nil, err
	}

	result := make([]byte, 3+len(utf16))
	binary.LittleEndian.PutUint16(result[0:2], uint16(len([]rune(s)))) // Character count
	result[2] = 0x01 // Unicode flag
	copy(result[3:], utf16)

	return result, nil
}

// Option is a functional option for configuring the Writer.
type Option func(*Writer)

// WithSheetName sets the sheet name.
func WithSheetName(name string) Option {
	return func(w *Writer) {
		w.sheetName = name
	}
}

// WriteToFile writes the data directly to a file with optional configurations.
func WriteToFile(filename string, data [][]interface{}, opts ...Option) error {
	w := New()
	defer w.Close()

	for _, opt := range opts {
		opt(w)
	}

	if err := w.Write(data); err != nil {
		return err
	}

	return w.SaveAs(filename)
}
