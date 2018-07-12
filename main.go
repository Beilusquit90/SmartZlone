package main

import (
	"fmt"
	"os"
	"path/filepath"

	"os/exec"
	"strconv"
	"sync"

	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/PuerkitoBio/goquery"
	"github.com/araddon/dateparse"
	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
	"golang.org/x/sys/windows/registry"
)

type MyMainWindow struct {
	*walk.MainWindow
	edit *walk.TextEdit
	path string
}

var edit *walk.TextEdit
var mykey string = ""
var template = ""
var templater = ""

func _check(err error) {
	if err != nil {
		panic(err)
	}
}

func parseUrl() int {
	url := "https://test.livejournal.com/835.html"
	doc, err := goquery.NewDocument(url)
	_check(err)

	tkey := ""
	flag := 0
	doc.Find("article").Each(func(i int, s *goquery.Selection) {
		if flag == 1 {
			tkey = strings.TrimSpace(s.Text())
		}
		flag++
	})
	ttkey := strings.Split(tkey, "-")
	for _, value := range ttkey {
		if value == mykey {
			return 1
		}
	}
	return 0
}

func setDay(day string, line int, coll int, xlsx *excelize.File) {
	kmap := []string{"W", "AA", "AE", "AI", "AM", "AQ", "AU", "AY", "BC", "BG", "BK", "BO", "BS", "BW", "CA", "CE", "CI", "CM", "CQ", "CU", "CY", "DC", "DG", "DK", "DO", "DS", "DW", "EA", "EE", "EI", "EM", "EQ", "EU", "EY", "FC"}

	if len(day) > 0 {
		drt := day[:2] + "." + day[3:5] + "." + day[6:]
		tt, err := dateparse.ParseLocal(drt)
		t := string(tt.String())

		z := coll
		for _, value := range t[8:10] + "  " + t[5:7] + " " + t[:4] {
			if value == ' ' {
				z++
				continue
			}
			xlsx.SetCellValue("Sheet1", kmap[z]+strconv.Itoa(line), string(value))
			z = z + 1
		}
		if err != nil {
			fmt.Println(err)
		}
	}
}
func setDayO(day string, line int, coll int, xlsx *excelize.File) {
	kmap := []string{"B", "D", "F", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z", "AB", "AD", "AF", "AH", "AJ", "AL", "AN", "AP", "AR", "AT", "AV", "AX", "AZ", "BB", "BD", "BF", "BH", "BJ", "BL", "BN", "BP", "BR", "BT", "BV", "BX"}

	if len(day) > 0 {
		drt := day[:2] + "." + day[3:5] + "." + day[6:]
		tt, err := dateparse.ParseLocal(drt)
		t := string(tt.String())

		z := coll
		for _, value := range t[8:10] + "  " + t[5:7] + " " + t[:4] {
			if value == ' ' {
				z++
				continue
			}
			xlsx.SetCellValue("Sheet1", kmap[z]+strconv.Itoa(line), string(value))
			z = z + 1
		}
		if err != nil {
			fmt.Println(err)
		}
	}
}
func setData(tdata string, line int, coll int, xlsx *excelize.File) {
	if len(tdata) > 0 {
		z := coll
		kmap := []string{"K", "O", "S", "W", "AA", "AE", "AI", "AM", "AQ", "AU", "AY", "BC", "BG", "BK", "BO", "BS", "BW", "CA", "CE", "CI", "CM", "CQ", "CU", "CY", "DC", "DG", "DK", "DO", "DS", "DW", "EA", "EE", "EI", "EM", "EQ", "EU", "EY", "FC"}
		for _, value := range tdata {
			if value == ' ' {
				z++
				continue
			}
			xlsx.SetCellValue("Sheet1", kmap[z]+strconv.Itoa(line), string(value))
			z = z + 1
		}
	}
}
func setDataO(tdata string, line int, coll int, xlsx *excelize.File) {
	if len(tdata) > 0 {
		z := coll
		kmap := []string{"B", "D", "F", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z", "AB", "AD", "AF", "AH", "AJ", "AL", "AN", "AP", "AR", "AT", "AV", "AX", "AZ", "BB", "BD", "BF", "BH", "BJ", "BL", "BN", "BP", "BR", "BT", "BV", "BX"}
		for _, value := range tdata {
			if value == ' ' {
				z++
				continue
			}
			xlsx.SetCellValue("Sheet1", kmap[z]+strconv.Itoa(line), string(value))
			z = z + 1
		}
	}
}
func fzData(tdata string, line int, coll int, xlsx *excelize.File) {
	if len(tdata) > 0 {
		z := coll
		kmap := []string{"M", "P", "S", "V", "Y", "AB", "AE", "AH", "AK", "AN", "AQ", "AT", "AW", "AZ", "BC", "BF", "BI", "BL", "BO", "BR", "BU", "BX", "CA", "CD", "CG", "CJ", "CM", "CP", "CS", "CV", "CY", "DB", "DE", "DH", "DK", "DN", "DQ", "DT", "DW"}
		for _, value := range tdata {
			if value == ' ' {
				z++
				continue
			}
			if z > (len(kmap) - 1) {
				continue
			}
			xlsx.SetCellValue("Sheet1", kmap[z]+strconv.Itoa(line), string(value))
			z = z + 1
		}
	}
}
func fzData2(tdata string, line int, coll int, xlsx *excelize.File) {
	if len(tdata) > 0 {
		z := coll
		kmap := []string{"A", "D", "G", "J", "M", "P", "S", "V", "Y", "AB", "AE", "AH", "AK", "AN", "AQ", "AT", "AW", "AZ", "BC", "BF", "BI", "BL", "BO", "BR", "BU", "BX", "CA", "CD", "CG", "CJ", "CM", "CP", "CS", "CV", "CY", "DB", "DE", "DH", "DK", "DN", "DQ", "DT", "DW"}
		fmt.Println(tdata)
		for _, value := range tdata {
			if value == ' ' {
				z++
				continue
			}

			if z > (len(kmap) - 1) {
				continue
			}
			xlsx.SetCellValue("Sheet1", kmap[z]+strconv.Itoa(line), string(value))

			z = z + 1
		}
	}
}
func setTd(tdata string, line int, coll string, xlsx *excelize.File) {
	if len(tdata) > 0 {
		xlsx.SetCellValue("Sheet1", coll+strconv.Itoa(line), tdata)
	}
}
func h2Create(human []string, count int, firm int) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}
	xlsx, err := excelize.OpenFile(dir + "/template/h2.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	if human[3] == "М" {
		xlsx.SetCellValue("Sheet1", "G20", "Х")
	} else {
		xlsx.SetCellValue("Sheet1", "K20", "Х")
	}
	tname := string(human[1]) + " " + string(human[2])
	setTd(tname, 18, "I", xlsx)
	setTd(string(human[0]), 18, "B", xlsx)
	drt2 := human[4][:2] + "." + human[4][3:5] + "." + human[4][6:]
	tt2, err := dateparse.ParseLocal(drt2)
	t2 := string(tt2.String())
	ttt, err := strconv.Atoi(t2[:4])
	if ttt > 2018 {
		t2 = "19" + t2[2:]
	}
	t2 = t2[8:10] + "." + t2[5:7] + "." + t2[:4]
	drt3 := human[8][:2] + "." + human[8][3:5] + "." + human[8][6:]
	tt3, err := dateparse.ParseLocal(drt3)
	t3 := string(tt3.String())
	t3 = t3[8:10] + "." + t3[5:7] + "." + t3[:4]
	drt4 := human[9][:2] + "." + human[9][3:5] + "." + human[9][6:]
	tt4, err := dateparse.ParseLocal(drt4)
	t4 := string(tt4.String())
	t4 = t4[8:10] + "." + t4[5:7] + "." + t4[:4]
	drt := human[12][:2] + "." + human[12][3:5] + "." + human[12][6:]
	tt, err := dateparse.ParseLocal(drt)
	t := string(tt.String())
	t = t[8:10] + "." + t[5:7] + "." + t[:4]
	drt5 := human[13][:2] + "." + human[13][3:5] + "." + human[13][6:]
	tt5, err := dateparse.ParseLocal(drt5)
	t5 := string(tt5.String())
	t5 = t5[8:10] + "." + t5[5:7] + "." + t5[:4]
	drt6 := human[14][:2] + "." + human[14][3:5] + "." + human[14][6:]
	tt6, err := dateparse.ParseLocal(drt6)
	t6 := string(tt6.String())
	t6 = t6[8:10] + "." + t6[5:7] + "." + t6[:4]
	pass := string(human[6]) + " / " + string(human[7])
	setTd(string(human[5]), 17, "N", xlsx)
	setTd(string(human[5]), 22, "H", xlsx)
	setTd(string(human[5]), 23, "H", xlsx)
	setTd(t2, 20, "M", xlsx)
	setTd(t3, 22, "W", xlsx)
	setTd(t4, 23, "W", xlsx)
	setTd(pass, 24, "H", xlsx)
	setTd(t6, 36, "U", xlsx)
	setTd(t, 26, "H", xlsx)
	setTd(t5, 17, "E", xlsx)
	setTd(string(human[10]), 25, "V", xlsx)
	setTd(string(human[11]), 25, "Z", xlsx)

	err = xlsx.SaveAs(dir + "/docs/h2-" + strconv.Itoa(count) + "-" + human[0] + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
func hCreate(human []string, count int, firm int) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}
	xlsx, err := excelize.OpenFile(dir + "/template/h.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	tname := ""

	for _, fname := range string(human[1]) {
		tname = tname + string(fname) + "."
		break
	}
	if len(human[2]) < 1 {
		tname = tname + ". " + string(human[0])
	} else {
		for _, fname := range string(human[2]) {
			tname = tname + string(fname) + ". " + string(human[0])
			break
		}
	}

	drt := human[12][:2] + "." + human[12][3:5] + "." + human[12][6:]
	tt, err := dateparse.ParseLocal(drt)
	t := string(tt.String())
	drt2 := human[4][:2] + "." + human[4][3:5] + "." + human[4][6:]
	tt2, err := dateparse.ParseLocal(drt2)
	t2 := string(tt2.String())
	ttt, err := strconv.Atoi(t2[:4])
	if ttt > 2018 {
		t = "19" + t2[2:]
	}

	drt3 := human[8][:2] + "." + human[8][3:5] + "." + human[8][6:]
	tt3, err := dateparse.ParseLocal(drt3)
	t3 := string(tt3.String())
	pass := string(human[6]) + " " + string(human[7])
	cx := 0
	country := ""
	for _, fname := range string(human[5]) {
		if cx == 0 {
			country = country + string(fname)
		} else {
			country = country + strings.ToLower(string(fname))
		}
		cx++
	}
	country = country + "а "
	if country == "Украинаа " {
		country = "Украины"
	}
	country = country + string(human[0]) + " " + string(human[1]) + " " + string(human[2]) + ", " + t2[8:10] + "." + t2[5:7] + "." + t2[:4] + " гр., паспорт №" + pass
	pd := "выдан " + t3[8:10] + "." + t3[5:7] + "." + t3[:4] + " г."
	dd := t[8:10] + "." + t[5:7] + "." + t[:4]

	setTd(country, 16, "A", xlsx)
	setTd(pd, 17, "A", xlsx)
	setTd(dd, 18, "I", xlsx)
	err = xlsx.SaveAs(dir + "/docs/h-" + strconv.Itoa(count) + "-" + human[0] + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
func dCreate(human []string, count int, firm int, c2 int) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}
	xlsx, err := excelize.OpenFile(dir + "/template/td.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	tname := ""

	for _, fname := range string(human[1]) {
		tname = tname + string(fname) + "."
		break
	}
	if len(human[2]) < 1 {
		tname = tname + ". " + string(human[0])
	} else {
		for _, fname := range string(human[2]) {
			tname = tname + string(fname) + ". " + string(human[0])
			break
		}
	}

	nname := data[firm][2]
	for _, fname := range string(data[firm][3]) {
		nname = nname + "." + string(fname)
		break
	}
	setTd(human[0]+" "+human[1]+" "+human[2], 6, "V", xlsx)
	setTd("ООО «"+string(data[firm][1])+"» , именуемое в дальнейшем «Работодатель» в лице Ген. директора "+nname, 5, "A", xlsx)
	setTd(human[0]+" "+human[1]+" "+human[2], 6, "V", xlsx)
	setTd(human[0]+" "+human[1]+" "+human[2], 35, "Y", xlsx)
	setTd(tname, 41, "AA", xlsx)
	setTd(nname, 41, "H", xlsx)
	xlsx.SetCellValue("Sheet1", "Q1", c2+count)
	setTd(human[5], 37, "AB", xlsx)
	if len(human[12]) > 0 {
		drt := human[12][:2] + "." + human[12][3:5] + "." + human[12][6:]
		tt, err := dateparse.ParseLocal(drt)
		t := string(tt.String())
		t = t[8:10] + "." + t[5:7] + "." + t[:4]
		setTd(t, 25, "N", xlsx)
		if err != nil {
			fmt.Println(err)
		}
	}
	if len(human[13]) > 0 {
		drt2 := human[13][:2] + "." + human[13][3:5] + "." + human[13][6:]
		tt2, err := dateparse.ParseLocal(drt2)
		t2 := string(tt2.String())
		setTd(t2[8:10]+"."+t2[5:7]+"."+t2[:4], 25, "S", xlsx)
		if err != nil {
			fmt.Println(err)
		}
	}
	if len(human[14]) > 0 {
		drt6 := human[14][:2] + "." + human[14][3:5] + "." + human[14][6:]
		tt6, err := dateparse.ParseLocal(drt6)
		t6 := string(tt6.String())
		t6 = t6[8:10] + "." + t6[5:7] + "." + t6[:4]
		setTd(t6+"г.", 3, "AC", xlsx)
		if err != nil {
			fmt.Println(err)
		}
	}

	pass := string(human[6]) + " " + string(human[7])
	setTd(pass, 36, "AE", xlsx)
	ad1 := string(data[firm][23]) + " Г." + string(data[firm][25])
	ad2 := "УЛ." + string(data[firm][26])
	if len(string(data[firm][28])) > 0 {
		ad2 = ad2 + " ДОМ." + string(data[firm][28])
	}
	if len(string(data[firm][29])) > 0 {
		ad2 = ad2 + " СТР." + string(data[firm][29])
	}
	if len(string(data[firm][30])) > 0 {
		ad2 = ad2 + " КВ." + string(data[firm][30])
	}

	setTd("6. Работнику устанавливается должностной оклад в размере пять тысяч рублей в месяц. Заработная плата выплачивается Работнику по адресу: "+ad1+" "+ad2, 21, "A", xlsx)
	setTd(ad1+" "+ad2, 36, "A", xlsx)
	setTd(ad1+" "+ad2, 39, "V", xlsx)
	setTd(string(data[firm][22]), 38, "C", xlsx)
	setTd("ООО «"+string(data[firm][1])+"»", 35, "A", xlsx)
	err = xlsx.SaveAs(dir + "/docs/td-" + strconv.Itoa(count) + "-" + human[0] + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
func fzC(human []string, count int, firm int) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}
	xlsx, err := excelize.OpenFile(dir + "/template/fz.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	if human[3] == "М" {
		xlsx.SetCellValue("Sheet1", "BL65", "Х")
	} else {
		xlsx.SetCellValue("Sheet1", "BU65", "Х")
	}
	drt := human[4][:2] + "." + human[4][3:5] + "." + human[4][6:]
	tt, err := dateparse.ParseLocal(drt)
	t := string(tt.String())
	t = t[8:10] + "  " + t[5:7] + " " + t[:4]
	drt2 := human[8][:2] + "." + human[8][3:5] + "." + human[8][6:]
	tt2, err := dateparse.ParseLocal(drt2)
	t2 := string(tt2.String())
	t2 = t2[8:10] + "  " + t2[5:7] + "  " + t2[:4]
	ttt, err := strconv.Atoi(t[:4])
	drt3 := human[12][:2] + "." + human[12][3:5] + "." + human[12][6:]
	tt3, err := dateparse.ParseLocal(drt3)
	t3 := string(tt3.String())
	t3 = t3[8:10] + "  " + t3[5:7] + "  " + t3[:4]
	if len(human[14]) > 0 {
		drt4 := human[14][:2] + "." + human[14][3:5] + "." + human[14][6:]
		tt4, err := dateparse.ParseLocal(drt4)
		t4 := string(tt4.String())
		t4 = t4[8:10] + "  " + t4[5:7] + "  " + t4[:4]
		fzData(t4, 84, 12, xlsx)
		fzData(t4, 116, 27, xlsx)
		if err != nil {
			fmt.Println(err)
		}
	}
	if ttt > 2018 {
		t = "19" + t[2:]
	}
	fzData(t, 65, 3, xlsx)
	fzData(human[1]+" "+human[2], 55, 4, xlsx)
	fzData(human[0], 53, 4, xlsx)
	fzData(human[5], 59, 4, xlsx)
	fzData(human[5], 73, 1, xlsx)
	fzData(human[5], 61, 4, xlsx)
	fzData(human[6], 71, 0, xlsx)
	fzData(human[7], 71, 11, xlsx)
	fzData(human[10], 75, 9, xlsx)
	fzData(human[11], 75, 14, xlsx)
	fzData(t2, 71, 27, xlsx)
	fzData(t3, 75, 27, xlsx)

	fzData2(human[23], 99, 0, xlsx)
	fzData2(human[17], 110, 0, xlsx)

	fzData2("'"+string(data[firm][1])+"'  ОРГН "+string(data[firm][31]), 34, 1, xlsx)
	fzData2(data[firm][31], 38, 0, xlsx)
	fzData2(("ИНН " + string(data[firm][22]) + " КПП " + string(data[firm][32])), 40, 0, xlsx)
	if len(string(data[firm][33])) > 0 {
		fzData2(data[firm][33], 50, 11, xlsx)
	}
	//fzData2(data[firm][31], 40, 0, xlsx)
	ad1 := string(data[firm][23]) + " Г." + string(data[firm][25])
	ad2 := "УЛ " + string(data[firm][26])
	if len(string(data[firm][28])) > 0 {
		ad2 = ad2 + " ДОМ " + string(data[firm][28])
	}
	if len(string(data[firm][29])) > 0 {
		ad2 = ad2 + " СТР " + string(data[firm][29])
	}
	if len(string(data[firm][30])) > 0 {
		ad2 = ad2 + " КВ " + string(data[firm][30])
	}
	fzData2(ad1, 42, 0, xlsx)
	fzData2(ad1, 78, 16, xlsx)
	fzData2(ad1, 119, 18, xlsx)
	fzData2(ad2, 44, 0, xlsx)
	fzData2(ad2, 80, 0, xlsx)
	fzData2(ad2, 121, 0, xlsx)
	if len(human[18]) > 0 {
		fzData("ПАТЕНТ", 88, 5, xlsx)
		drt5 := human[20][:2] + "." + human[20][3:5] + "." + human[20][6:]
		tt5, err := dateparse.ParseLocal(drt5)
		t5 := string(tt5.String())
		t5 = t5[8:10] + "  " + t5[5:7] + "  " + t5[:4]
		drt6 := human[21][:2] + "." + human[21][3:5] + "." + human[21][6:]
		tt6, err := dateparse.ParseLocal(drt6)
		t6 := string(tt6.String())
		t6 = t6[8:10] + "  " + t6[5:7] + "  " + t6[:4]
		fmt.Println(t5, t6)
		fzData2(t5, 96, 6, xlsx)
		fzData2(t5, 92, 5, xlsx)
		fzData2(t6, 96, 21, xlsx)
		fzData2(human[22], 94, 5, xlsx)
		fzData2(human[18], 90, 5, xlsx)
		fzData2(human[19], 90, 19, xlsx)

		if err != nil {
			fmt.Println(err)
		}

	}

	err = xlsx.SaveAs(dir + "/docs/fz-" + strconv.Itoa(count) + "-" + human[0] + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
func addH(human []string, count int, firm int) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}

	xlsx, err := excelize.OpenFile(dir + "/template/" + template + ".xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	if human[3] == "М" {
		xlsx.SetCellValue("Sheet1", "CY21", "Х")
		xlsx.SetCellValue("Sheet1", "DC77", "Х")
	} else {
		xlsx.SetCellValue("Sheet1", "DS21", "Х")
		xlsx.SetCellValue("Sheet1", "DW77", "Х")
	}
	if len(data[firm][9]) > 0 {
		drt2 := data[firm][9] + "." + data[firm][9] + "." + data[firm][9]
		tt2, err := dateparse.ParseLocal(drt2)
		t2 := string(tt2.String())
		ttt2, err := strconv.Atoi(t2[:4])

		if ttt2 > 2018 {
			t2 = "19" + t2[2:]
		}
		if err != nil {
			fmt.Println(err)
			return
		}

		setDay(t2[8:10]+"  "+t2[5:7]+" "+t2[:4], 136, 17, xlsx)
	}
	drt := human[4][:2] + "." + human[4][3:5] + "." + human[4][6:]
	tt, err := dateparse.ParseLocal(drt)
	t := string(tt.String())
	ttt, err := strconv.Atoi(t[:4])

	if ttt > 2018 {
		t = "19" + t[2:]
	}

	if len(data[firm][7]) > 0 {
		setDay(data[firm][7], 128, 24, xlsx)
	}
	setDay(human[8], 32, 1, xlsx)
	setDay(human[9], 32, 17, xlsx)
	setDay(human[12], 47, 3, xlsx)
	setDay(human[13], 47, 24, xlsx)
	setDay(human[13], 167, 24, xlsx)
	setDay(human[13], 95, 5, xlsx)
	setDay(data[firm][7], 128, 24, xlsx)
	setDay(data[firm][8], 136, 1, xlsx)

	setData(t[8:10]+"  "+t[5:7]+" "+t[:4], 21, 5, xlsx)
	setData(t[8:10]+"  "+t[5:7]+" "+t[:4], 77, 5, xlsx)
	setData(human[0], 13, 3, xlsx)
	setData(human[0], 69, 3, xlsx)
	setData(human[1]+" "+human[2], 15, 3, xlsx)
	setData(human[1]+" "+human[2], 71, 3, xlsx)
	setData(human[5], 18, 4, xlsx)
	setData(human[5], 74, 4, xlsx)
	setData(human[5], 24, 5, xlsx)
	setData("ПАСПОРТ", 30, 11, xlsx)
	setData("ПАСПОРТ", 80, 11, xlsx)
	setData(human[6], 30, 24, xlsx)
	setData(human[6], 80, 24, xlsx)
	setData(human[7], 30, 29, xlsx)
	setData(human[7], 80, 29, xlsx)
	setData(human[10], 49, 8, xlsx)
	setData(human[11], 49, 13, xlsx)
	setData(human[17], 45, 3, xlsx)
	setData(data[firm][2], 169, 3, xlsx)
	setData(data[firm][2], 128, 3, xlsx)
	setData(data[firm][3], 172, 3, xlsx)
	setData(data[firm][3], 131, 3, xlsx)
	setData(data[firm][4], 133, 11, xlsx)
	setData(data[firm][5], 133, 24, xlsx)
	setData(data[firm][6], 133, 29, xlsx)
	setData(data[firm][10], 139, 5, xlsx)
	setData(data[firm][11], 141, 3, xlsx)
	setData(data[firm][12], 144, 5, xlsx)
	setData(data[firm][13], 146, 3, xlsx)
	setData(data[firm][14], 148, 2, xlsx)
	setData(data[firm][15], 148, 8, xlsx)
	setData(data[firm][16], 148, 15, xlsx)
	setData(data[firm][17], 148, 22, xlsx)
	setData(data[firm][18], 151, 4, xlsx)
	setData(data[firm][19], 153, 0, xlsx)
	setData(data[firm][20], 155, 4, xlsx)
	setData(data[firm][21], 157, 0, xlsx)
	setData(data[firm][22], 159, 2, xlsx)
	setData(data[firm][23], 114, 5, xlsx)
	setData(data[firm][24], 116, 3, xlsx)
	setData(data[firm][25], 119, 5, xlsx)
	setData(data[firm][26], 121, 3, xlsx)
	setData(data[firm][27], 123, 2, xlsx)
	setData(data[firm][28], 123, 8, xlsx)
	setData(data[firm][29], 123, 15, xlsx)
	setData(data[firm][30], 123, 22, xlsx)
	setData(data[firm][23], 84, 5, xlsx)
	setData(data[firm][24], 86, 3, xlsx)
	setData(data[firm][25], 88, 5, xlsx)
	setData(data[firm][26], 91, 3, xlsx)
	setData(data[firm][27], 93, 2, xlsx)
	setData(data[firm][28], 93, 8, xlsx)
	setData(data[firm][29], 93, 15, xlsx)
	setData(data[firm][30], 93, 22, xlsx)
	target := []string{"AM43", "AY43", "BO43", "CA43", "CM43", "CY43", "DK43", "EE43", "EQ43"}
	tr, err := strconv.Atoi(string(human[16]))
	xlsx.SetCellValue("Sheet1", target[tr], "X")
	err = xlsx.SaveAs(dir + "/data/" + strconv.Itoa(count) + "-" + human[0] + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
func oldHm(human []string, count int, firm int) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}

	xlsx, err := excelize.OpenFile(dir + "/template/" + template + ".xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	if human[3] == "М" {
		xlsx.SetCellValue("Sheet1", "AV19", "Х")
		xlsx.SetCellValue("Sheet1", "AV88", "Х")
	} else {
		xlsx.SetCellValue("Sheet1", "BF19", "Х")
		xlsx.SetCellValue("Sheet1", "BF88", "Х")
	}
	drt := human[4][:2] + "." + human[4][3:5] + "." + human[4][6:]
	tt, err := dateparse.ParseLocal(drt)
	t := string(tt.String())
	ttt, err := strconv.Atoi(t[:4])

	if ttt > 2018 {
		t = "19" + t[2:]
	}
	setDayO(human[8], 32, 4, xlsx)
	setDayO(human[9], 32, 20, xlsx)
	setDayO(human[12], 49, 7, xlsx)
	setDayO(human[13], 49, 27, xlsx)
	setDayO(human[13], 190, 27, xlsx)
	setDayO(human[13], 112, 8, xlsx)
	setDayO(data[firm][8], 152, 4, xlsx)
	setDayO(data[firm][9], 152, 20, xlsx)
	if len(data[firm][7]) > 0 {
		drt2 := data[firm][7][:2] + "." + data[firm][7][3:5] + "." + data[firm][7][6:]
		tt2, err := dateparse.ParseLocal(drt2)
		t2 := string(tt2.String())
		ttt2, err := strconv.Atoi(t2[:4])
		if ttt2 > 2018 {
			t2 = "19" + t2[2:]
		}
		if err != nil {
			fmt.Println(err)
		}
		setDataO(t2[8:10]+"  "+t2[5:7]+"  "+t2[:4], 143, 26, xlsx)
	}

	setDataO(t[8:10]+"  "+t[5:7]+" "+t[:4], 20, 5, xlsx)
	setDataO(t[8:10]+"  "+t[5:7]+" "+t[:4], 89, 5, xlsx)
	setDataO(human[0], 11, 3, xlsx)
	setDataO(human[0], 80, 3, xlsx)
	setDataO(human[1]+" "+human[2], 14, 3, xlsx)
	setDataO(human[1]+" "+human[2], 83, 3, xlsx)
	setDataO(human[5], 17, 4, xlsx)
	setDataO(human[5], 86, 4, xlsx)
	setDataO(human[5], 23, 5, xlsx)
	setDataO("ПАСПОРТ", 29, 11, xlsx)
	setDataO("ПАСПОРТ", 92, 11, xlsx)
	setDataO(human[6], 29, 24, xlsx)
	setDataO(human[6], 92, 24, xlsx)
	setDataO(human[7], 29, 29, xlsx)
	setDataO(human[7], 92, 29, xlsx)
	setDataO(human[10], 52, 8, xlsx)
	setDataO(human[11], 52, 13, xlsx)
	setDataO(human[17], 45, 3, xlsx)

	setDataO(data[firm][2], 193, 3, xlsx)
	setDataO(data[firm][2], 143, 3, xlsx)
	setDataO(data[firm][3], 196, 3, xlsx)
	setDataO(data[firm][3], 146, 3, xlsx)
	setDataO(data[firm][4], 149, 11, xlsx)
	setDataO(data[firm][5], 149, 24, xlsx)
	setDataO(data[firm][6], 149, 29, xlsx)
	setDataO(data[firm][10], 155, 5, xlsx)
	setDataO(data[firm][11], 158, 3, xlsx)
	setDataO(data[firm][12], 161, 5, xlsx)
	setDataO(data[firm][13], 164, 3, xlsx)
	setDataO(data[firm][14], 167, 2, xlsx)
	setDataO(data[firm][15], 167, 8, xlsx)
	setDataO(data[firm][16], 167, 15, xlsx)
	setDataO(data[firm][17], 167, 22, xlsx)
	setDataO(data[firm][18], 172, 4, xlsx)
	setDataO(data[firm][19], 175, 4, xlsx)
	setDataO(data[firm][20]+" "+data[firm][21], 178, 0, xlsx)
	setDataO(data[firm][22], 181, 2, xlsx)

	setDataO(data[firm][23], 97, 5, xlsx)
	setDataO(data[firm][24], 100, 3, xlsx)
	setDataO(data[firm][25], 103, 5, xlsx)
	setDataO(data[firm][26], 106, 3, xlsx)
	setDataO(data[firm][28], 109, 2, xlsx)
	setDataO(data[firm][29], 109, 15, xlsx)
	setDataO(data[firm][30], 109, 22, xlsx)

	setDataO(data[firm][23], 125, 5, xlsx)
	setDataO(data[firm][24], 128, 3, xlsx)
	setDataO(data[firm][25], 131, 5, xlsx)
	setDataO(data[firm][26], 134, 3, xlsx)

	setDataO(data[firm][28], 137, 2, xlsx)
	setDataO(data[firm][29], 137, 15, xlsx)
	setDataO(data[firm][30], 137, 22, xlsx)
	target := []string{"P42", "V42", "AF42", "AL42", "AR42", "AX42", "BD42", "BL42", "BR42"}
	tr, err := strconv.Atoi(string(human[16]))
	xlsx.SetCellValue("Sheet1", target[tr], "X")
	err = xlsx.SaveAs(dir + "/data/" + strconv.Itoa(count) + "-" + human[0] + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
func oldHmr(human []string, count int, firm int) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}

	xlsx, err := excelize.OpenFile(dir + "/template/" + templater + ".xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	if human[3] == "М" {

		xlsx.SetCellValue("Sheet1", "AV88", "Х")
	} else {

		xlsx.SetCellValue("Sheet1", "BF88", "Х")
	}
	drt := human[4][:2] + "." + human[4][3:5] + "." + human[4][6:]
	tt, err := dateparse.ParseLocal(drt)
	t := string(tt.String())
	ttt, err := strconv.Atoi(t[:4])

	if ttt > 2018 {
		t = "19" + t[2:]
	}

	setDayO(human[13], 190, 27, xlsx)
	setDayO(human[13], 112, 8, xlsx)
	setDayO(data[firm][7], 152, 4, xlsx)
	setDayO(data[firm][8], 152, 20, xlsx)

	drt2 := data[firm][9][:2] + "." + data[firm][9][3:5] + "." + data[firm][9][6:]
	tt2, err := dateparse.ParseLocal(drt2)
	t2 := string(tt2.String())

	if err != nil {
		fmt.Println(err)
	}
	setDataO(t2[8:10]+"  "+t2[5:7]+"  "+t2[:4], 143, 26, xlsx)

	setDataO(t[8:10]+"  "+t[5:7]+" "+t[:4], 89, 5, xlsx)

	setDataO(human[0], 80, 3, xlsx)

	setDataO(human[1]+" "+human[2], 83, 3, xlsx)

	setDataO(human[5], 86, 4, xlsx)

	setDataO("ПАСПОРТ", 92, 11, xlsx)

	setDataO(human[6], 92, 24, xlsx)
	setDataO(human[7], 92, 29, xlsx)

	setDataO(data[firm][2], 193, 3, xlsx)
	setDataO(data[firm][3], 196, 3, xlsx)

	setDataO(data[firm][23], 97, 5, xlsx)
	setDataO(data[firm][24], 100, 3, xlsx)
	setDataO(data[firm][25], 103, 5, xlsx)
	setDataO(data[firm][26], 106, 3, xlsx)
	setDataO(data[firm][28], 109, 2, xlsx)
	setDataO(data[firm][29], 109, 15, xlsx)
	setDataO(data[firm][30], 109, 22, xlsx)

	err = xlsx.SaveAs(dir + "/data/" + strconv.Itoa(count) + "-" + human[0] + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
func addHr(human []string, count int, firm int) {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}

	xlsx, err := excelize.OpenFile(dir + "/template/" + templater + ".xlsx")
	fmt.Println(templater)
	if err != nil {
		fmt.Println(err)
		return
	}

	if human[3] == "М" {

		xlsx.SetCellValue("Sheet1", "DC77", "Х")
	} else {

		xlsx.SetCellValue("Sheet1", "DW77", "Х")
	}
	drt := human[4][:2] + "." + human[4][3:5] + "." + human[4][6:]
	tt, err := dateparse.ParseLocal(drt)
	t := string(tt.String())
	ttt, err := strconv.Atoi(t[:4])

	if ttt > 2018 {
		t = "19" + t[2:]
	}

	setDay(human[13], 167, 24, xlsx)
	setDay(human[13], 95, 5, xlsx)
	setDay(data[firm][8], 136, 1, xlsx)
	setDay(data[firm][9], 136, 17, xlsx)

	setData(t[8:10]+"  "+t[5:7]+" "+t[:4], 77, 5, xlsx)

	setData(human[0], 69, 3, xlsx)

	setData(human[1]+" "+human[2], 71, 3, xlsx)

	setData(human[5], 74, 4, xlsx)

	setData("ПАСПОРТ", 80, 11, xlsx)

	setData(human[6], 80, 24, xlsx)

	setData(human[7], 80, 29, xlsx)

	setData(data[firm][2], 169, 3, xlsx)

	setData(data[firm][3], 172, 3, xlsx)

	setData(data[firm][23], 84, 5, xlsx)
	setData(data[firm][24], 86, 3, xlsx)
	setData(data[firm][25], 88, 5, xlsx)
	setData(data[firm][26], 91, 3, xlsx)
	setData(data[firm][27], 93, 2, xlsx)
	setData(data[firm][28], 93, 8, xlsx)
	setData(data[firm][29], 93, 15, xlsx)
	setData(data[firm][30], 93, 22, xlsx)
	err = xlsx.SaveAs(dir + "/data/" + strconv.Itoa(count) + "-" + human[0] + ".xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
func GetCreate(x int) {

	if patchf == "" {
		return
	}
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}
	xcount, err := excelize.OpenFile(dir + "/template/count.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	xlre, err := excelize.OpenFile(patchf)
	if err != nil {
		fmt.Println(err)
		return
	}

	c2 := string(xcount.GetCellValue("Sheet1", "A1"))
	cc2, err := strconv.Atoi(c2)
	dataZ := [][]string{}
	azb := []string{"B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"}
	for count := 2; len(xlre.GetCellValue("Sheet1", "B"+strconv.Itoa(count))) > 0; count++ {
		tdata := []string{}
		for _, cc := range azb {
			tdata = append(tdata, xlre.GetCellValue("Sheet1", cc+strconv.Itoa(count)))
		}
		dataZ = append(dataZ, tdata)
	}
	var wg sync.WaitGroup
	count := 1
	for _, human := range dataZ {
		wg.Add(1)
		go func(human []string, count int, firm int) {
			defer wg.Done()
			fmt.Println(template)
			if template == "old" {
				oldHm(human, count, x)
			} else {
				addH(human, count, x)
			}
			fzC(human, count, x)
			if string(human[15]) == "1" {
				dCreate(human, count, x, cc2)
				hCreate(human, count, x)
				//h2Create(human, count, x)

			}
		}(human, count, x)
		count++
	}
	wg.Wait()
	xcount.SetCellValue("Sheet1", "A1", cc2+count)
	xcount.Save()
}
func GetCreater(x int) {
	if patchf == "" {
		return
	}
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
	}
	xcount, err := excelize.OpenFile(dir + "/template/count.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	xlre, err := excelize.OpenFile(patchf)
	if err != nil {
		fmt.Println(err)
		return
	}

	c2 := string(xcount.GetCellValue("Sheet1", "A1"))
	cc2, err := strconv.Atoi(c2)
	dataZ := [][]string{}
	azb := []string{"B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S"}
	for count := 2; len(xlre.GetCellValue("Sheet1", "B"+strconv.Itoa(count))) > 0; count++ {
		tdata := []string{}
		for _, cc := range azb {
			tdata = append(tdata, xlre.GetCellValue("Sheet1", cc+strconv.Itoa(count)))
		}
		dataZ = append(dataZ, tdata)
	}
	var wg sync.WaitGroup
	count := 1
	for _, human := range dataZ {
		wg.Add(1)
		go func(human []string, count int, firm int) {
			defer wg.Done()
			if templater == "oldr" {
				fmt.Println("Older")
				oldHmr(human, count, x)
			} else {
				addHr(human, count, x)
			}
		}(human, count, x)
		count++
	}
	wg.Wait()
	xcount.SetCellValue("Sheet1", "A1", cc2+count)
	xcount.Save()
}
func GetKey() {
	k, err := registry.OpenKey(registry.CURRENT_USER, `Software\Microsoft\Windows\CurrentVersion\Audio\`, registry.QUERY_VALUE|registry.SET_VALUE)
	//k, err := registry.OpenKey(registry.CURRENT_USER, `Software`, registry.QUERY_VALUE|registry.SET_VALUE)
	z, _, err := k.GetStringValue("LiteMicro")
	if err != nil {
		fmt.Println("404")
	}
	mykey = string(z)
}

var patchf = ""
var data = [][]string{}
var tmpldata = [][]string{}

func main() {

	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println("405")
	}
	xlre, err := excelize.OpenFile("template/firms.xlsx")
	if err != nil {
		fmt.Println("406")
		return
	}
	azb := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH"}
	for count := 2; len(xlre.GetCellValue("Sheet1", "B"+strconv.Itoa(count))) > 0; count++ {
		tdata := []string{}
		for _, cc := range azb {
			tdata = append(tdata, xlre.GetCellValue("Sheet1", cc+strconv.Itoa(count)))
		}
		data = append(data, tdata)
	}
	xlrt, err := excelize.OpenFile("template/templ.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	for count := 2; len(xlrt.GetCellValue("Sheet1", "B"+strconv.Itoa(count))) > 0; count++ {
		tdata := []string{}
		tdata = append(tdata, xlrt.GetCellValue("Sheet1", "A"+strconv.Itoa(count)))
		tdata = append(tdata, xlrt.GetCellValue("Sheet1", "B"+strconv.Itoa(count)))
		tdata = append(tdata, xlrt.GetCellValue("Sheet1", "C"+strconv.Itoa(count)))
		tmpldata = append(tmpldata, tdata)
	}
	mw := &MyMainWindow{}
	zzz := []Widget{}
	zzr := []Widget{}
	xxx := []Widget{}
	for x := range tmpldata {
		tmp, err := strconv.Atoi(tmpldata[x][0])
		if err != nil {
			fmt.Println(err)
			return
		}
		name := tmpldata[x][1]
		xxx = append(xxx, PushButton{
			Text: "Шаб-" + name,
			OnClicked: func() {
				GetKey()
				if len(mykey) > 0 {
					ff := 1
					if tmpldata[tmp][1] == "" {
						ff = 0
					}
					if tmpldata[tmp][2] == "" {
						ff = 0
					}
					if ff == 0 {
						mw.edit.SetText("Одно из имён введённого шаблона пустое, либо не коректное. Исправьте это и перезапустите программу.\n")
					}
					template = tmpldata[tmp][1]
					templater = tmpldata[tmp][2]
					_, err := excelize.OpenFile(dir + "/template/" + template + ".xlsx")
					if err != nil {
						mw.edit.SetText("Шаблона с именем: " + template + ".xlsx, не существует в папке template")
						fmt.Println(err)
						return
					} else {
						if err != nil {
							mw.edit.SetText("Шаблона с именем: " + templater + ".xlsx, не существует в папке template")
							fmt.Println(err)
							return
						}
					}
					mw.edit.SetText("Были применены следующие шаблоны для работы:\r\nДля полной клетки: " + template + ".xlsx\r\nДля отрывной клетки: " + templater + ".xlsx\r\n")

				} else {
					mw.edit.SetText("Ваша программа не активирована.\r\nДля активации запустите от имени администратора один раз template/SmartCloneKey.exe\r\nПосле чего перезапустите программу и следуйте дальнейшим инструкциям.")
				}
			},
		})
	}
	for x := range data {
		name := data[x][1]
		namer := "ОТР-" + name
		tmp, err := strconv.Atoi(data[x][0])
		if err != nil {
			fmt.Println(err)
			return
		}

		zzz = append(zzz, PushButton{
			Text: name,
			OnClicked: func() {
				GetKey()
				if len(mykey) > 0 {
					if template == "" {
						mw.edit.SetText("Для начала работы, выберете шаблон клетки нужной кнопкой.\n")
						return
					}
					if patchf == "" {
						mw.edit.SetText("Что бы начать работу, сначала откройте файл с данными для генерации документа.\n")
						return
					}

					if parseUrl() == 1 {
						mw.edit.SetText("Началась обработка данных. Пожалуйста, ничего не трогайте.\n")
						GetCreate(tmp)
						mw.edit.SetText("Процесс обработки данных и создание документов завершен.\r\nДокументы были помещены в папку Data рядом с программой. \r\nДля продолжения работы, выберите следующий документ для обработки и опять запустите процесс.\r\nЛибо активируйте запуск обработки документов используя другую фирму.")
					} else {
						mw.edit.SetText("Ваш ключ сейчас не активен.\r\nДля активации ключа, надо связаться с разработчиком программы и сообщить ему номер указанный ниже.\r\n" + mykey + "\r\n")

					}
				} else {
					mw.edit.SetText("Ваша программа не активирована.\r\nДля активации запустите от имени администратора один раз template/SmartCloneKey.exe\r\nПосле чего перезапустите программу и следуйте дальнейшим инструкциям.")
				}
			},
		})
		zzr = append(zzr, PushButton{
			Text: namer,
			OnClicked: func() {
				GetKey()
				if len(mykey) > 0 {
					if templater == "" {
						mw.edit.SetText("Для начала работы, выберете шаблон клетки нужной кнопкой.\n")
						return
					}
					if patchf == "" {
						mw.edit.SetText("Что бы начать работу, сначала откройте файл с данными для генерации документа.\n")
						return
					}
					if parseUrl() == 1 {
						mw.edit.SetText("Началась обработка данных. Пожалуйста, ничего не трогайте.\n")
						GetCreater(tmp)
						mw.edit.SetText("Процесс обработки данных и создание ОТРЫВНОЙ ЧАСТИ завершен.\r\nДокументы были помещены в папку Data рядом с программой. \r\nДля продолжения работы, выберите следующий документ для обработки и опять запустите процесс.\r\nЛибо активируйте запуск обработки документов используя другую фирму.")
					} else {
						mw.edit.SetText("Ваш ключ сейчас не активен.\r\nДля активации ключа, надо связаться с разработчиком программы и сообщить ему номер указанный ниже.\r\n" + mykey + "\r\n")

					}
				} else {
					mw.edit.SetText("Ваша программа не активирована.\r\nДля активации запустите от имени администратора один раз template/SmartCloneKey.exe\r\nПосле чего перезапустите программу и следуйте дальнейшим инструкциям.")
				}
			},
		})
	}

	MW := MainWindow{
		AssignTo: &mw.MainWindow,
		Title:    "SmartClone",
		MinSize:  Size{500, 400},
		Size:     Size{500, 400},
		Layout:   VBox{},
		Children: []Widget{
			TextEdit{
				AssignTo: &mw.edit, ReadOnly: true,
			},
			PushButton{
				Text:      "Открыть строчку",
				OnClicked: mw.pbClicked,
			},
			PushButton{
				Text: "Открыть папку с документами.",
				OnClicked: func() {
					err = exec.Command("explorer", dir+"\\data").Start()
				},
			},
			HSplitter{
				MinSize:  Size{30, 20},
				MaxSize:  Size{30, 20},
				Children: xxx,
			},
			HSplitter{
				MinSize:  Size{50, 20},
				MaxSize:  Size{50, 20},
				Children: zzz,
			},
			HSplitter{
				MinSize:  Size{50, 20},
				MaxSize:  Size{50, 20},
				Children: zzr,
			},
		},
	}
	if _, err := MW.Run(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}

}

func (mw *MyMainWindow) pbClicked() {
	dlg := new(walk.FileDialog)

	dlg.FilePath = mw.path
	dlg.Title = "Select File"
	dlg.Filter = "Exe files (*.xlsx)|*.xlsx|All files (*.*)|*.*"

	if ok, err := dlg.ShowOpen(mw); err != nil {
		mw.edit.AppendText("Error : File Open\r\n")
		return
	} else if !ok {
		mw.edit.AppendText("Cancel\r\n")
		return
	}
	mw.path = dlg.FilePath
	patchf = dlg.FilePath
	s := fmt.Sprintf("Select : %s\r\n", mw.path) + "Выберите предпочитаемый шаблон, после чего нажмите кнопку с именем нужной вам фирмы для запуска работы\n."
	mw.edit.SetText(s)
}

func (mw *MyMainWindow) openEx() {
	s := fmt.Sprintf("%s\r\n", mw.path+"\\data")
	fmt.Println(s)
	err := exec.Command("explorer.exe", s).Start()
	if err != nil {
		fmt.Println(err)
	}
}
