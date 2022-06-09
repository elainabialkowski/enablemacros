package macros

import (
	"fmt"
	"os"
	"path"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestEnableExcelMacros(t *testing.T) {
	dirpath := setup(t)
	err := EnableExcelMacros(dirpath)
	if err != nil {
		t.Fatalf("%s\n", err.Error())
	}
}

func setup(t *testing.T) string {
	dirPath, err := os.MkdirTemp("./", "testing_*")
	if err != nil {
		t.Fatalf("Could not setup test directory: %s\n", err.Error())
	}

	for i := 0; i < 10000; i++ {
		f := excelize.NewFile()

		err = f.SaveAs(path.Join(dirPath, fmt.Sprintf("test_file_%d.xlsx", i)))
		if err != nil {
			t.Fatalf("Could not setup test directory: %s\n", err.Error())
		}

		err = f.Close()
		if err != nil {
			t.Fatalf("Could not setup test directory: %s\n", err.Error())
		}
	}

	return dirPath

}
