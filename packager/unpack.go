// change print out file system

package packager

import (

	"archive/zip"
	"bytes"
	"fmt"
	"path"
	"strings"

    "github.com/prr123/godocx/common/constants"
    "github.com/prr123/godocx/docx"
    "github.com/prr123/godocx/internal"
    "github.com/prr123/godocx/wml/ctypes"
)

// ReadFromZip reads files from a zip archive.
func ReadFromZip(content *[]byte) (map[string][]byte, error) {
	zipReader, err := zip.NewReader(bytes.NewReader(*content), int64(len(*content)))
	if err != nil {
		return nil, err
	}

	var (
		fileList = make(map[string][]byte, len(zipReader.File))
	)

	for _, f := range zipReader.File {

		fileName := strings.ReplaceAll(f.Name, "\\", "/")

		if fileList[fileName], err = internal.ReadFileFromZip(f); err != nil {
			return nil, err
		}
	}

	return fileList, nil
}

func Unpack(content *[]byte) (*docx.RootDoc, error) {

	rd := docx.NewRootDoc()

	fileIndex, err := ReadFromZip(content)
	if err != nil {
		return nil, err
	}

// new
    count:=0
    for filnam, _ := range fileIndex {
        count++
        fmt.Printf("file %d: %s\n", count, filnam)
    }

	// Load content type details
	ctBytes := fileIndex[constants.ConentTypeFileIdx]
	ct, err := LoadContentTypes(ctBytes)
	if err != nil {
		return nil, err
	}
	delete(fileIndex, constants.ConentTypeFileIdx)
	rd.ContentType = *ct

	rd.ImageCount = 0

	rootRelURI, err := GetRelsURI("")
	if err != nil {
		return nil, err
	}

	rootRelBytes := fileIndex[*rootRelURI]
	rootRelations, err := LoadRelationShips(*rootRelURI, rootRelBytes)
	if err != nil {
		return nil, err
	}
	delete(fileIndex, *rootRelURI)
	rd.RootRels = *rootRelations

	var docPath string

	for _, relation := range rootRelations.Relationships {
		switch relation.Type {
		case constants.OFFICE_DOC_TYPE:
			docPath = relation.Target
		}
	}

	if docPath == "" {
		return nil, fmt.Errorf("root officeDocument type not found")
	}

	docRelURI, err := GetRelsURI(docPath)
	if err != nil {
		return nil, err
	}

	// Load document
//	fmt.Printf("docPath: %s\n", docPath)

	docFile := fileIndex[docPath]
	docObj, err := docx.LoadDocXml(rd, docPath, docFile)
	if err != nil {
		return nil, err
	}
	delete(fileIndex, docPath)
	rd.Document = docObj

	// Load Relationship details
	docRelFile := fileIndex[*docRelURI]
	docRelations, err := LoadRelationShips(*docRelURI, docRelFile)
	if err != nil {
		return nil, err
	}
	delete(fileIndex, *rootRelURI)
	rd.Document.DocRels = *docRelations

	wordDir := path.Dir(docPath)

	rd.DocStyles = &ctypes.Styles{}
	rID := 0
	for _, relation := range docRelations.Relationships {
		rID += 1
		switch relation.Type {
		case constants.StylesType:
			sFileName := relation.Target
			if sFileName == "" {
				continue
			}
			stylesPath := path.Join(wordDir, sFileName)

			//Load Styles
			stylesFile := fileIndex[stylesPath]
			stylesObj, err := docx.LoadStyles(stylesPath, stylesFile)
			if err != nil {
				return nil, err
			}
			delete(fileIndex, stylesPath)
			rd.DocStyles = stylesObj
		default:
			fmt.Printf("relTyp: %v\n", relation.Type)
		}

	}

	rd.Document.RID = rID

	// retrieve numbering
	for _, relation := range docRelations.Relationships {
		switch relation.Type {
		case constants.NumberingType:
			nFileName := relation.Target
			if nFileName == "" {
				continue
			}
			numPath := path.Join(wordDir, nFileName)

			//Load Styles
			listFile, ok := fileIndex[numPath]
			fmt.Printf("numPath: %s ok: %t\n", numPath, ok)
			listnumObj, err := docx.LoadListnum(Path, listFile)
			if err != nil {
				return nil, err
			}
			delete(fileIndex, numPath)
			rd.DocStyles = stylesObj
			rd.DocLists = listnumObj
		}
//	numCont
	}

	for fileName, fileContent := range fileIndex {
		if strings.HasPrefix(fileName, constants.MediaPath) {
			rd.ImageCount += 1
		}
		rd.FileMap.Store(fileName, fileContent)
	}

	return rd, nil
}
