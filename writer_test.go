package xls

import (
	"os"
	"testing"
)

func TestNew(t *testing.T) {
	w := New()
	if w == nil {
		t.Fatal("New() returned nil")
	}
	if w.sheetName != "Sheet1" {
		t.Errorf("Expected default sheet name 'Sheet1', got '%s'", w.sheetName)
	}
	w.Close()
}

func TestSetSheetName(t *testing.T) {
	w := New()
	defer w.Close()

	newName := "TestSheet"
	w.SetSheetName(newName)

	if w.sheetName != newName {
		t.Errorf("Expected sheet name '%s', got '%s'", newName, w.sheetName)
	}
}

func TestWrite(t *testing.T) {
	w := New()
	defer w.Close()

	data := [][]interface{}{
		{"Name", "Age", "City"},
		{"Alice", 30, "Tokyo"},
		{"Bob", 25, "Osaka"},
	}

	err := w.Write(data)
	if err != nil {
		t.Fatalf("Write() failed: %v", err)
	}

	if len(w.data) != len(data) {
		t.Errorf("Expected data length %d, got %d", len(data), len(w.data))
	}
}

func TestSaveAs(t *testing.T) {
	w := New()
	defer w.Close()

	data := [][]interface{}{
		{"Name", "Age"},
		{"Alice", 30},
	}

	err := w.Write(data)
	if err != nil {
		t.Fatalf("Write() failed: %v", err)
	}

	tmpFile := "test_save.xls"
	defer os.Remove(tmpFile)

	err = w.SaveAs(tmpFile)
	if err != nil {
		t.Fatalf("SaveAs() failed: %v", err)
	}

	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("File was not created")
	}

	info, err := os.Stat(tmpFile)
	if err != nil {
		t.Fatalf("Failed to stat file: %v", err)
	}
	if info.Size() == 0 {
		t.Error("File size is 0")
	}
}

func TestWriteToFile(t *testing.T) {
	tmpFile := "test_write_to_file.xls"
	defer os.Remove(tmpFile)

	data := [][]interface{}{
		{"Header1", "Header2", "Header3"},
		{1, 2, 3},
		{"A", "B", "C"},
	}

	err := WriteToFile(tmpFile, data)
	if err != nil {
		t.Fatalf("WriteToFile() failed: %v", err)
	}

	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("File was not created")
	}

	info, err := os.Stat(tmpFile)
	if err != nil {
		t.Fatalf("Failed to stat file: %v", err)
	}
	if info.Size() == 0 {
		t.Error("File size is 0")
	}
}

func TestWriteToFileWithSheetName(t *testing.T) {
	tmpFile := "test_write_to_file_with_sheet.xls"
	defer os.Remove(tmpFile)

	sheetName := "TestData"
	data := [][]interface{}{
		{"Item", "Quantity"},
		{"Apple", 10},
		{"Banana", 20},
	}

	err := WriteToFile(tmpFile, data, WithSheetName(sheetName))
	if err != nil {
		t.Fatalf("WriteToFile() with WithSheetName() failed: %v", err)
	}

	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("File was not created")
	}
}

func TestWriteEmptyData(t *testing.T) {
	w := New()
	defer w.Close()

	data := [][]interface{}{}

	err := w.Write(data)
	if err != nil {
		t.Fatalf("Write() with empty data failed: %v", err)
	}

	tmpFile := "test_empty.xls"
	defer os.Remove(tmpFile)

	err = w.SaveAs(tmpFile)
	if err != nil {
		t.Fatalf("SaveAs() failed: %v", err)
	}
}

func TestWriteWithDifferentTypes(t *testing.T) {
	w := New()
	defer w.Close()

	data := [][]interface{}{
		{"String", "Int", "Float", "Bool"},
		{"text", 42, 3.14, true},
		{"another", -10, -2.5, false},
	}

	err := w.Write(data)
	if err != nil {
		t.Fatalf("Write() with different types failed: %v", err)
	}

	tmpFile := "test_types.xls"
	defer os.Remove(tmpFile)

	err = w.SaveAs(tmpFile)
	if err != nil {
		t.Fatalf("SaveAs() failed: %v", err)
	}

	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Fatal("File was not created")
	}
}

func TestWriteLargeData(t *testing.T) {
	w := New()
	defer w.Close()

	// Create 100 rows x 10 columns of data
	data := make([][]interface{}, 100)
	for i := 0; i < 100; i++ {
		data[i] = make([]interface{}, 10)
		for j := 0; j < 10; j++ {
			if j == 0 {
				data[i][j] = "Row " + string(rune('A'+i%26))
			} else {
				data[i][j] = i*10 + j
			}
		}
	}

	err := w.Write(data)
	if err != nil {
		t.Fatalf("Write() with large data failed: %v", err)
	}

	tmpFile := "test_large.xls"
	defer os.Remove(tmpFile)

	err = w.SaveAs(tmpFile)
	if err != nil {
		t.Fatalf("SaveAs() failed: %v", err)
	}
}

func TestSharedStringTable(t *testing.T) {
	sst := newSST()

	sst.addString("Hello")
	sst.addString("World")
	sst.addString("Hello") // duplicate

	if sst.uniqueCount != 2 {
		t.Errorf("Expected uniqueCount 2, got %d", sst.uniqueCount)
	}

	if sst.totalCount != 3 {
		t.Errorf("Expected totalCount 3, got %d", sst.totalCount)
	}

	if idx := sst.getIndex("Hello"); idx != 0 {
		t.Errorf("Expected index 0 for 'Hello', got %d", idx)
	}

	if idx := sst.getIndex("World"); idx != 1 {
		t.Errorf("Expected index 1 for 'World', got %d", idx)
	}
}

func TestEncodeString(t *testing.T) {
	str := "Test"
	encoded, err := encodeString(str)
	if err != nil {
		t.Fatalf("encodeString() failed: %v", err)
	}

	minLen := 3 + len(str)*2
	if len(encoded) < minLen {
		t.Errorf("Expected encoded length at least %d, got %d", minLen, len(encoded))
	}

	if encoded[0] != byte(len(str)) {
		t.Errorf("Expected length byte %d, got %d", len(str), encoded[0])
	}

	if encoded[1] != 0x01 {
		t.Errorf("Expected Unicode flag 0x01, got 0x%02x", encoded[1])
	}
}

func TestEncodeStringForSST(t *testing.T) {
	str := "SST"
	encoded, err := encodeStringForSST(str)
	if err != nil {
		t.Fatalf("encodeStringForSST() failed: %v", err)
	}

	minLen := 3 + len(str)*2
	if len(encoded) < minLen {
		t.Errorf("Expected encoded length at least %d, got %d", minLen, len(encoded))
	}

	if encoded[2] != 0x01 {
		t.Errorf("Expected Unicode flag 0x01, got 0x%02x", encoded[2])
	}
}
