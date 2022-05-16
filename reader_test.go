package xlsx

import (
	"archive/zip"
	"log"
	"os"
	"strings"
	"testing"
)

func TestParseShared(t *testing.T) {
	sts := []string{
		"A", "B", "C", "D",
	}
	const sharedStringsText = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>A</t></si><si><t>B</t></si><si><t>C</t></si><si><t>D</t></si></sst>`
	r := strings.NewReader(sharedStringsText)
	ss, err := parseShared(r)
	if err != nil {
		t.Fatalf("Unexpected err: %q", err)
	}

	for i := range ss {
		if ss[i] != sts[i] {
			t.Fatalf("Unexpected: %s<>%s", ss[i], sts[i])
		}
	}
}

func TestParseSharedWithEmpty(t *testing.T) {
	sts := []string{
		"A", "B", "", "C", "", "D",
	}
	const sharedStringsText = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>A</t></si><si><t>B</t></si><si><t/></si><si><t>C</t></si><si><t/></si><si><t>D</t></si></sst>`
	r := strings.NewReader(sharedStringsText)
	ss, err := parseShared(r)
	if err != nil {
		t.Fatalf("Unexpected err: %q", err)
	}

	for i := range ss {
		if ss[i] != sts[i] {
			t.Fatalf("Unexpected: %s<>%s", ss[i], sts[i])
		}
	}
}

const xlsxFile = "test/spreadsheet.xlsx"

func TestParseContentType(t *testing.T) {
	file, err := os.Open(xlsxFile)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	st, err := file.Stat()
	if err != nil {
		t.Fatal(err)
	}

	zr, err := zip.NewReader(file, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	for _, zFile := range zr.File {
		// read where the worksheets are
		if zFile.Name == "[Content_Types].xml" {
			index, err := parseContentType(zFile)
			if err != nil {
				t.Fatal(err)
			}

			if len(index.files) != 2 {
				t.Fatalf("Unexpected len: %d. Expected 2", len(index.files))
			}
			if index.files[0] != "/xl/worksheets/sheet1.xml" {
				t.Fatalf("Unexpected spreadsheet file: %s", index.files[0])
			}
			if index.files[1] != "/xl/worksheets/sheet2.xml" {
				t.Fatalf("Unexpected spreadsheet file: %s", index.files[1])
			}

			if index.sharedStr != "/xl/sharedStrings.xml" {
				t.Fatalf("Unexpected sharedStrings file: %s", index.sharedStr)
			}

			break
		}
	}
}

func TestReadShared(t *testing.T) {
	file, err := os.Open(xlsxFile)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	st, err := file.Stat()
	if err != nil {
		t.Fatal(err)
	}

	zr, err := zip.NewReader(file, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	for _, zFile := range zr.File {
		// read where the worksheets are
		if zFile.Name == "[Content_Types].xml" {
			index, err := parseContentType(zFile)
			if err != nil {
				t.Fatal(err)
			}

			if len(index.files) != 2 {
				t.Fatalf("Unexpected len: %d. Expected 1", len(index.files))
			}
			if index.files[0] != "/xl/worksheets/sheet1.xml" {
				t.Fatalf("Unexpected spreadsheet file: %s", index.files[0])
			}
			if index.files[1] != "/xl/worksheets/sheet2.xml" {
				t.Fatalf("Unexpected spreadsheet file: %s", index.files[1])
			}

			if index.sharedStr != "/xl/sharedStrings.xml" {
				t.Fatalf("Unexpected sharedStrings file: %s", index.sharedStr)
			}

			shared, err := readShared(zr, index.sharedStr)
			if err != nil {
				t.Fatal(err)
			}

			expectedShared := []string{
				"Date", "A", "B", "C", "D",
			}
			for i := range shared {
				if shared[i] != expectedShared[i] {
					t.Fatalf("%s <> %s", shared[i], expectedShared[i])
				}
			}

			break
		}
	}
}

func TestReadFile(t *testing.T) {
	file, err := Open(xlsxFile)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	expectedRows := map[string][][]string{
		"sheet1": {
			{"Date", "A", "B", "C", "D"},
			{"43922", "1", "2", "3", "4"},
			{"43923", "5", "6", "7", "8"},
		},
		"sheet2": {
			{"A", "B", "C", "D"},
			{"1", "2", "3", "4"},
			{"5", "6", "7", "8"},
		},
	}

	for _, sheet := range file.Sheets() {
		expected := expectedRows[sheet.Name]

		r, err := sheet.Open()
		if err != nil {
			log.Fatalln(err)
		}

		i := 0
		for r.Next() {
			for j, s := range r.Row() {
				if s != expected[i][j] {
					t.Fatalf("%s <> %s", s, expected[i][j])
				}
			}
			i++
		}
		if r.Error() != nil {
			t.Fatal(r.Error())
		}

		r.Close()
	}
}
