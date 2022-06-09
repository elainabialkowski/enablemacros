package macros

import (
	"fmt"
	"io/fs"
	"log"
	"os"
	"path"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

func EnableExcelMacros(root string) error {
	dir := os.DirFS(root)
	files, err := fs.Glob(dir, "*.xlsx")
	if err != nil {
		return err
	}

	wg := &sync.WaitGroup{}
	os.Mkdir(path.Join(root, "output"), os.ModeDir)
	for _, v := range files {

		wg.Add(1)
		go func(v string) error {
			defer wg.Done()

			inputPath := path.Join(root, v)
			f, err := excelize.OpenFile(inputPath)
			if err != nil {
				log.Printf("Could not open xlsx file: %s\n", err.Error())
			}
			defer f.Close()

			outputPath := path.Join(root, "output", fmt.Sprintf("%s.xlsm", strings.Split(v, ".")[0]))
			err = f.SaveAs(outputPath)
			if err != nil {
				log.Printf("Could not save xlsm file: %s\n", err.Error())
			}

			return nil

		}(v)
	}

	wg.Wait()

	return nil

}
