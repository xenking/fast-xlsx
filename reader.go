package xlsx

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	xml "github.com/dgrr/quickxml"
)

// XLSX ...
type XLSX struct {
	sharedStrings []string
	zr            *zip.Reader
	sheets        []*Sheet
	closer        io.Closer
}

// sheetData
//   row: r="1"
//     c: r="A1" t="inlineStr"|"n" s="1"
//       is:
//         t:
//       v:

// xlsxIndex represents an index of the sharedStrings file and
// the spreadsheets files.
type xlsxIndex struct {
	sharedStr string
	files     []string
}

// Close closes all the buffers and readers.
func (xlsx *XLSX) Close() error {
	if xlsx.zr == nil {
		return nil
	}

	if xlsx.closer != nil {
		return xlsx.closer.Close()
	}

	return nil
}

// Sheets returns the sheets.
func (xlsx *XLSX) Sheets() []*Sheet {
	return xlsx.sheets
}

// SharedStrings returns the shared strings.
func (xlsx *XLSX) SharedStrings() []string {
	return xlsx.sharedStrings
}

// Open just opens the file for reading.
func Open(filename string) (*XLSX, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}

	st, err := file.Stat()
	if err != nil {
		return nil, err
	}

	return OpenReader(file, st.Size())
}

// OpenReader opens the reader as XLSX file.
func OpenReader(r io.ReaderAt, size int64) (*XLSX, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, err
	}

	var (
		xlsx  *XLSX
		index xlsxIndex
	)

	sheetsName := make(map[string]string)

	for _, zFile := range zr.File {
		switch zFile.Name {
		case "xl/workbook.xml":
			sheetsName, err = parseWorkbook(zFile)
			if err != nil {
				return nil, fmt.Errorf("parseWorkbook: %s", err)
			}

		case "[Content_Types].xml":
			// read where the worksheets are
			index, err = parseContentType(zFile)
			if err != nil {
				return nil, fmt.Errorf("parseContentType: %s", err)
			}
		}
	}

	// read the worksheets
	xlsx, err = extractWorksheets(zr, &index, sheetsName)
	if err == nil {
		closer, ok := r.(io.Closer)
		if ok {
			xlsx.closer = closer
		}
	}

	return xlsx, err
}

func parseWorkbook(zFile *zip.File) (sheets map[string]string, err error) {
	sheets = make(map[string]string)
	var zfr io.ReadCloser

	zfr, err = zFile.Open()
	if err != nil {
		return
	}
	defer zfr.Close()

	r := xml.NewReader(zfr)
	for err == nil && r.Next() {
		switch e := r.Element().(type) {
		case *xml.StartElement:
			if !bytes.Equal(e.NameBytes(), sheetString) {
				continue
			}

			var sheetID string
			kv := e.Attrs().GetBytes(sheetIDString)
			if kv == nil {
				err = errors.New("sheetId parameter not found")
			} else {
				sheetID = kv.Value()
			}

			var name string
			kv = e.Attrs().GetBytes(sheetNameString)
			if kv == nil {
				err = errors.New("name parameter not found")
			} else {
				name = kv.Value()
			}

			sheets[sheetID] = name
		}
	}
	if err == nil {
		if r.Error() != nil && r.Error() != io.EOF {
			err = r.Error()
		}
	}

	return
}

func getPartName(e *xml.StartElement) (partName string, err error) {
	kv := e.Attrs().GetBytes(partNameString)
	if kv != nil {
		partName = kv.Value()
	}
	if partName == "" {
		err = errors.New("PartName parameter not found")
	}
	return
}

func parseContentType(zFile *zip.File) (index xlsxIndex, err error) {
	var zfr io.ReadCloser

	zfr, err = zFile.Open()
	if err != nil {
		return
	}
	defer zfr.Close()

	r := xml.NewReader(zfr)
	for err == nil && r.Next() {
		switch e := r.Element().(type) {
		case *xml.StartElement:
			if !bytes.Equal(e.NameBytes(), overrideString) {
				continue
			}

			var partName string
			kv := e.Attrs().GetBytes(contentTypeString)
			if kv != nil {
				switch {
				case bytes.Equal(kv.ValueBytes(), workSheetURIString):
					partName, err = getPartName(e)
					if err == nil {
						index.files = append(index.files, partName)
					}
				case bytes.Equal(kv.ValueBytes(), sharedStringsURIString):
					partName, err = getPartName(e)
					if err == nil {
						index.sharedStr = partName
					}
				}
			}
		}
	}
	if err == nil {
		if r.Error() != nil && r.Error() != io.EOF {
			err = r.Error()
		} else if len(index.files) == 0 {
			err = errors.New("no data files found")
		}
	}

	return
}

func extractWorksheets(zr *zip.Reader, index *xlsxIndex, sheetsName map[string]string) (*XLSX, error) {
	var (
		err    error
		shared []string
	)
	sharedFile := index.sharedStr

	if len(sharedFile) > 0 {
		shared, err = readShared(zr, sharedFile)
		if err != nil {
			return nil, fmt.Errorf("error reading shared strings: %s", err)
		}
	}

	xs := new(XLSX)
	xs.sharedStrings = shared

	for _, filename := range index.files {
		zFile, err := getZipFile(zr, filename)
		if err != nil {
			xs = nil
			return nil, err
		}

		sheetIDStart := strings.LastIndex(filename, "sheet")
		sheetIDEnd := strings.LastIndexByte(filename, '.')
		sheetID := filename[sheetIDStart+5 : sheetIDEnd]

		xs.sheets = append(xs.sheets, &Sheet{
			Name:   sheetsName[sheetID],
			parent: xs,
			zFile:  zFile,
		})
	}

	return xs, err
}

func findNameIn(name, where string) bool {
	if name[0] == '/' {
		return name[1:] == where
	}
	return strings.Contains(where, name)
}

func getZipFile(zr *zip.Reader, filename string) (zFile *zip.File, err error) {
	var found = false
	for _, zFile = range zr.File {
		found = findNameIn(filename, zFile.Name)
		if found {
			break
		}
	}
	if !found {
		err = fmt.Errorf("%s not found", filename)
	}

	return zFile, err
}

func readShared(zr *zip.Reader, filename string) ([]string, error) {
	var (
		rc    io.ReadCloser
		found bool
		err   error
	)
	for _, zFile := range zr.File {
		found = findNameIn(filename, zFile.Name)
		if found {
			rc, err = zFile.Open()
			break
		}
	}
	if !found {
		err = fmt.Errorf("%s not found", filename)
	}
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	return parseShared(rc)
}

func parseShared(rc io.Reader) ([]string, error) {
	var err error
	ss := make([]string, 0)
	r := xml.NewReader(rc)
	T := false
loop:
	for r.Next() {
		switch e := r.Element().(type) {
		case *xml.StartElement:
			T = bytes.Equal(e.NameBytes(), tString)
			if T && e.HasEnd() {
				// shared strings sometimes contains empty strings. Don't know why
				ss = append(ss, "")
			}
		case *xml.TextElement:
			if T {
				ss = append(ss, string(*e))
			}
		case *xml.EndElement:
			if bytes.Equal(e.NameBytes(), sstString) {
				break loop
			}
		}
	}

	return ss, err
}
