# go-xls

A simple Go library for writing Excel XLS files (BIFF8/Excel 97-2003 format) from 2D slices with minimal dependencies

## ⚠️ Status

**This library is not production ready.** It is currently in early development and may have bugs or incomplete features. Use at your own risk.

## Features

- Simple API
- Generate XLS files from 2D slices (`[][]interface{}`)
- Native BIFF8 format implementation (Excel 97-2003)
- No external dependencies (only `golang.org/x/text`)
- Support for various data types: strings, numbers, booleans
- UTF-16LE character encoding support

## Important Note

This library generates legacy Excel format (.xls / BIFF8). If you need the modern Excel format (.xlsx), please use other libraries such as [github.com/xuri/excelize](https://github.com/xuri/excelize).

## Installation

```bash
go get github.com/tkuchiki/go-xls
```

## Usage

### Basic Usage

```go
package main

import (
    "log"
    "github.com/tkuchiki/go-xls"
)

func main() {
    data := [][]interface{}{
        {"Name", "Age", "City"},
        {"Alice", 30, "Tokyo"},
        {"Bob", 25, "Osaka"},
        {"Charlie", 35, "Kyoto"},
    }

    if err := xls.WriteToFile("output.xls", data); err != nil {
        log.Fatal(err)
    }
}
```

### Custom Sheet Name

```go
data := [][]interface{}{
    {"Product", "Price", "Stock"},
    {"Apple", 100, 50},
    {"Banana", 80, 100},
}

err := xls.WriteToFile("products.xls", data, xls.WithSheetName("Product List"))
if err != nil {
    log.Fatal(err)
}
```

Note: The current implementation sets the sheet name, but does not support multiple sheets.

### Using Writer for More Control

```go
writer := xls.New()
defer writer.Close()

writer.SetSheetName("Sales Report")

data := [][]interface{}{
    {"Month", "Sales", "Profit"},
    {"January", 10000, 2000},
    {"February", 12000, 2400},
}

if err := writer.Write(data); err != nil {
    log.Fatal(err)
}

if err := writer.SaveAs("sales.xls"); err != nil {
    log.Fatal(err)
}
```

## Supported Data Types

- `string` - Strings (UTF-16LE encoding)
- `int`, `int8`, `int16`, `int32`, `int64` - Integers
- `uint`, `uint8`, `uint16`, `uint32`, `uint64` - Unsigned integers
- `float32`, `float64` - Floating point numbers
- `bool` - Boolean values
- Other types - Converted to string via `fmt.Sprintf("%v", value)`

## API

### Functions

#### `WriteToFile(filename string, data [][]interface{}, opts ...Option) error`

Writes 2D slice data directly to a file with optional configurations.

**Parameters:**
- `filename`: Path to the output XLS file
- `data`: Data to write (2D slice)
- `opts`: Optional configuration options (e.g., `WithSheetName()`)

**Returns:**
- `error` if an error occurred, `nil` on success

#### `WithSheetName(name string) Option`

Returns an option to set a custom sheet name.

**Parameters:**
- `name`: Sheet name to set

**Returns:**
- `Option` function to configure the Writer

### Writer Type

#### `New() *Writer`

Creates a new Writer.

**Returns:**
- A new `*Writer` instance

#### `(*Writer) SetSheetName(name string)`

Sets the sheet name.

**Parameters:**
- `name`: Sheet name to set

#### `(*Writer) Write(data [][]interface{}) error`

Stores 2D slice data in memory.

**Parameters:**
- `data`: Data to write (2D slice)

**Returns:**
- `error` if an error occurred, `nil` on success

#### `(*Writer) SaveAs(filename string) error`

Saves the stored data as an XLS file to the specified path.

**Parameters:**
- `filename`: Path to the output XLS file

**Returns:**
- `error` if an error occurred, `nil` on success

#### `(*Writer) Close() error`

Releases resources. Currently does nothing but provided for future extensions.

**Returns:**
- Always `nil`

## Examples

See example/main.go for usage examples.

```bash
cd example
go run main.go
```

This will generate:
- `simple.xls` - Basic usage example
- `products.xls` - Custom sheet name example
- `sales.xls` - Writer usage example

## Tests

```bash
go test -v
```

## Technical Details

This library implements the BIFF8 (Binary Interchange File Format version 8) format used by Excel 97-2003.

### Implemented BIFF Records

- **BOF** (Beginning of File)
- **EOF** (End of File)
- **DIMENSIONS** - Worksheet dimension information
- **ROW** - Row definition
- **LABELSST** - String cell (via Shared String Table)
- **NUMBER** - Number cell
- **BOOLERR** - Boolean/Error cell
- **SST** (Shared String Table)
- **CODEPAGE** - Character encoding
- **FONT** - Font definition
- **XF** (Extended Format) - Format definition
- **STYLE** - Style definition
- And many more...

### Limitations

- Multiple sheets are not supported
- Cell formatting (colors, fonts, borders, etc.) is not supported
- Formula writing is not supported
- Image and chart embedding is not supported

If you need these features, consider using libraries that support the XLSX format.

## License

MIT
