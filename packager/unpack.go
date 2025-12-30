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

/*
    count:=0
    for filnam, _ := range fileIndex {
        count++
        fmt.Printf("file %d: %s\n", count, filnam)
    }
*/

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
			if err != nil {return nil, err}

			delete(fileIndex, stylesPath)
			rd.DocStyles = stylesObj

		case constants.NumberingType:
			nFileName := relation.Target
			if nFileName == "" {
				continue
			}
			numPath := path.Join(wordDir, nFileName)

			//Load Styles
			numbFile, ok := fileIndex[numPath]
//			fmt.Printf("numPath: %s ok: %t\n", numPath, ok)
			if !ok {continue}

			numObj, err := docx.LoadNumbering(numPath, numbFile)
			if err != nil {return nil, fmt.Errorf("LoadNumbering: %v",err)}
			delete(fileIndex, numPath)

			nmgr := rd.Numbering
    		if numObj != nil {
				nmgr.Fill(numObj, rd)
			}
			rd.Numbering = nmgr

			rd.DLists = make([]docx.DocxList, len(numObj.List))

			for ni:=1; ni<  len(numObj.List)+1; ni++ {
				an:= numObj.NMap[ni-1]
				an1 := numObj.List[an].AbstNumId
//	fmt.Printf("%d: abs num: %d %d num: %d\n", ni, an, an1, ni)
				dl:=rd.DLists[ni-1]
				dl.AbId = an1
				dl.Ord = true
//	fmt.Printf("ord: %t, abnum: %d\n", dl.Ord, dl.AbId)
				if numObj.List[an].Lvl[0].NumFmt.Val == "bullet" {dl.Ord = false}
//		fmt.Printf(" numFmt: %s\n", num.List[ni-1].Lvl[0].NumFmt.Val)
				for il:=0; il<9; il++ {
					dl.Mark[il] = numObj.List[an].Lvl[il].NumFmt.Val
					dl.Start[il] = numObj.List[an].Lvl[il].Start.Val
				}
				rd.DLists[ni-1] = dl
//	fmt.Printf("ord: %t, abnum: %d\n", dl.Ord, dl.AbId)
			}

		default:
			fmt.Printf("relTyp: %v\n", relation.Type)
		}

	}

	rd.Document.RID = rID


	for fileName, fileContent := range fileIndex {
		if strings.HasPrefix(fileName, constants.MediaPath) {
			rd.ImageCount += 1
		}
		rd.FileMap.Store(fileName, fileContent)
	}

	return rd, nil
}
