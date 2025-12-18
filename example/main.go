package main

import (
	"fmt"
	"log"

	"github.com/tkuchiki/go-xls"
)

func main() {
	// Example 1: Simple usage
	simpleExample()

	// Example 2: Custom sheet name
	customSheetExample()

	// Example 3: Using Writer for more control
	writerExample()
}

func simpleExample() {
	fmt.Println("Example 1: Simple usage")

	data := [][]interface{}{
		{"Name", "Age", "City"},
		{"Alice", 30, "Tokyo"},
		{"Bob", 25, "Osaka"},
		{"Charlie", 35, "Kyoto"},
	}

	if err := xls.WriteToFile("simple.xls", data); err != nil {
		log.Fatalf("Failed to write file: %v", err)
	}

	fmt.Println("  Created: simple.xls")
}

func customSheetExample() {
	fmt.Println("Example 2: Custom sheet name")

	data := [][]interface{}{
		{"Product", "Price", "Stock"},
		{"Apple", 100, 50},
		{"Banana", 80, 100},
		{"Orange", 120, 30},
	}

	if err := xls.WriteToFile("products.xls", data, xls.WithSheetName("Product List")); err != nil {
		log.Fatalf("Failed to write file: %v", err)
	}

	fmt.Println("  Created: products.xls")
}

func writerExample() {
	fmt.Println("Example 3: Using Writer for more control")

	// Numeric data example
	data := [][]interface{}{
		{"Month", "Sales", "Profit"},
		{"January", 10000, 2000},
		{"February", 12000, 2400},
		{"March", 15000, 3000},
		{"April", 13000, 2600},
	}

	writer := xls.New()
	defer writer.Close()

	writer.SetSheetName("Sales Report")

	if err := writer.Write(data); err != nil {
		log.Fatalf("Failed to write data: %v", err)
	}

	if err := writer.SaveAs("sales.xls"); err != nil {
		log.Fatalf("Failed to save file: %v", err)
	}

	fmt.Println("  Created: sales.xls")
}
