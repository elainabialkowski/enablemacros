package main

import (
	"os"

	macros "github.com/elainabialkowski/enablemacros"
)

func main() {
	macros.EnableExcelMacros(os.Args[1])
}
