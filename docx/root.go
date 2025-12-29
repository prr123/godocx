package docx

import (
	"fmt"
	"encoding/xml"
	"sync"

    "github.com/prr123/godocx/wml/ctypes"
)

// RootDoc represents the root document of an Office Open XML (OOXML) document.
// It contains information about the document path, file map, the document structure,
// and relationships with other parts of the document.
type RootDoc struct {
	Path        string        // Path represents the path of the document.
	FileMap     sync.Map      // FileMap is a synchronized map for managing files related to the document.
	RootRels    Relationships // RootRels represents relationships at the root level.
	ContentType ContentTypes
	Document    *Document      // Document is the main document structure.
	DocStyles   *ctypes.Styles // Document styles
	Numbering   *NumberingManager // Numbering manager for list instances
	DLists		[]DocxList  //List types
	rID        int // rId is used to generate unique relationship IDs.
	ImageCount uint
}

// NewRootDoc creates a new instance of the RootDoc structure.
func NewRootDoc() *RootDoc {
	root := &RootDoc{}
	root.Numbering = NewNumberingManager(root)
	return root
}

// LoadDocXml decodes the provided XML data and returns a Document instance.
// It is used to load the main document structure from the document file.
//
// Parameters:
//   - fileName: The name of the document file.
//   - fileBytes: The XML data representing the main document structure.
//
// Returns:
//   - doc: The Document instance containing the decoded main document structure.
//   - err: An error, if any occurred during the decoding process.
func LoadDocXml(rd *RootDoc, fileName string, fileBytes []byte) (*Document, error) {
	doc := Document{
		Root: rd,
	}
	err := xml.Unmarshal(fileBytes, &doc)
	if err != nil {
		return nil, err
	}

	doc.relativePath = fileName
	return &doc, nil
}

// Load styles.xml into Styles struct
func LoadStyles(fileName string, fileBytes []byte) (*ctypes.Styles, error) {
	styles := ctypes.Styles{}
	err := xml.Unmarshal(fileBytes, &styles)
	if err != nil {
		return nil, err
	}

	styles.RelativePath = fileName
	return &styles, nil
}

// Load numbering.xml onto the numberingManager
func LoadNumbering(fileName string, contNumbering []byte) (numObj *Numbering, err error) {

//	filMap:= rdoc.FileMap

//    contNumbering, ok := filMap.Load("word/numbering.xml")
//    if !ok {return nil, fmt.Errorf("cannot load file 'word/numbering.xml'!/n")}

    numObj=&Numbering{}

    err = xml.Unmarshal(contNumbering, numObj)
    if err != nil {return nil, fmt.Errorf("unmarshal: %v\n", err)}

	nlen :=len(numObj.Instances)
	nMap := make(map[int]int)
    for inum:=0; inum<nlen; inum++ {
        nb := numObj.Instances[inum]
//		nMap[nb.AbstNumId.Val] = nb.NumId -1
		nMap[nb.NumId -1] = nb.AbstractNumId.Val
//        fmt.Printf("  Numb: %d Id: %d Abst Id: %d\n",inum, nb.NumId, nb.AbstNumId.Val)
    }

	numObj.NMap = nMap

	return numObj, nil
}


// NewListInstance creates a new numbering instance for the given abstract numbering ID.
// Returns the numId that can be used with paragraph.Numbering().
//
// Parameters:
//   - abstractNumId: An integer representing the abstract numbering definition ID (1-8 or custom).
//
// Returns:
//   - numId: An integer that can be used with paragraph.Numbering(numId, level).
//
// Example:
//
//	numId := doc.NewListInstance(1)
//	p.Numbering(numId, 0)
func (root *RootDoc) NewListInstance(abstractNumId int) int {
	return root.Numbering.NewListInstance(abstractNumId)
}
