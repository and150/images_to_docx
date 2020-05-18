package main

import (
    "fmt"
    "io/ioutil"
    "path/filepath"
    "os"
    "regexp"
    "strconv"
    "sort"

    "github.com/unidoc/unioffice/color"
    "github.com/unidoc/unioffice/common"
    "github.com/unidoc/unioffice/document"
    "github.com/unidoc/unioffice/measurement"
    
    "github.com/unidoc/unioffice/schema/soo/wml"
    "github.com/unidoc/unioffice/schema/soo/ofc/sharedTypes"
)


func main() {

    sourceFolder := os.Args[1]
    imageNames := getFileList(sourceFolder)
    sort.Sort(ByNum(imageNames))
    makeDocument("out.docx", sourceFolder, imageNames)
}



func getFileList(sourceFolder string) []string {

    var fileNames []string

    files, err := ioutil.ReadDir(sourceFolder)
    if err != nil {
        fmt.Println("There was an error: ", err)
    }

    for _, file := range files{
        fileNames = append(fileNames, file.Name())
    }

    return fileNames
}


func getSortKey(name string) int {

    r, _ := regexp.Compile("(\\d+)")

    key, err := strconv.Atoi(r.FindString(name))
    if err != nil {
        key = 0
        for _,i := range name {key += int(i)}
    }

    return key
}

// sort interface for filenames 
type ByNum []string

func (a ByNum) Len() int            {return len(a)}
func (a ByNum) Swap(i, j int)       {a[i], a[j] = a[j], a[i]}
func (a ByNum) Less(i, j int) bool  {return getSortKey(a[i]) < getSortKey(a[j])}



func makeDocument(docName, imageFolder string, imageNames []string) {

    var pic_width float32 = 410

    doc := document.New()

    table := doc.AddTable()
    table.Properties().SetWidthPercent(100)

    borders := table.Properties().Borders()
    borders.SetAll(wml.ST_BorderSingle, color.Auto, 1*measurement.Point)

    for i, file := range imageNames {
        img1, err := common.ImageFromFile(filepath.Join(imageFolder, file))
        if err != nil {
            fmt.Println("Error: ", err)
        }
        img1ref, err := doc.AddImage(img1)
        if err != nil {
            fmt.Println("Error: ", err)
        }

        ratio := float32(img1.Size.Y)/float32(img1.Size.X)

        row := table.AddRow()
        row.AddCell().AddParagraph().AddRun().AddText(fmt.Sprint(i+1))
        inl,_ := row.AddCell().AddParagraph().AddRun().AddDrawingInline(img1ref)

        inl.SetSize(measurement.Distance(pic_width), measurement.Distance(pic_width*ratio))
        row.AddCell().AddParagraph().AddRun().AddText(file)
    }

    section := doc.BodySection()
    section.X().PgSz = wml.NewCT_PageSz()
    section.X().PgSz.OrientAttr = wml.ST_PageOrientationLandscape

    var w sharedTypes.ST_TwipsMeasure
    var h sharedTypes.ST_TwipsMeasure

    wInt := uint64(15840)
    hInt := uint64(12240)

    w.ST_UnsignedDecimalNumber = &wInt
    h.ST_UnsignedDecimalNumber = &hInt

    section.X().PgSz.WAttr = &w
    section.X().PgSz.HAttr = &h

    doc.SaveToFile("out.docx")
}
