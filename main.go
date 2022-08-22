package main

import (
	"fmt"
	"os"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

func init() {
	//fmt.Println("init main")
}

type Item struct {
	Code  string
	Time1 int64
	Price float64
	Time2 int64
}

type Result struct {
	Code  string
	Time1 string
	Price float64
	Time2 string
	Flag string
	LastPrice float64
}

func main() {
	fmt.Println("读入文件：",os.Args[1])
	f, err := excelize.OpenFile(os.Args[1])
	if err != nil {
		fmt.Println("open file error : ",err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Get value from cell by given worksheet name and axis.
	//cell, err := f.GetCellValue("Sheet1", "B2")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//fmt.Println(cell)
	// Get all the rows in the Sheet1.
	//rows, err := f.GetRows("Sheet1")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//for i, row := range rows {
	//	fmt.Println("---",i,row)
	//	for j, colCell := range row {
	//		fmt.Print(i,"-",j,"-",colCell, "\t")
	//	}
	//
	//	fmt.Println("\n")
	//}
	//toBeCharge := "01-01-2015"
	//timeLayout := "01-02-2006"
	//loc, _ := time.LoadLocation("Local")


	data := make([]*Item,0,50000)
	   rows1, err := f.Rows("Sheet1")
	   if err != nil {
	       fmt.Println(err)
	       return
	   }
	   if rows1.Next(){
		   for rows1.Next() {

			   row, err := rows1.Columns(excelize.Options{RawCellValue:true})
			   if err != nil {
				   fmt.Println(err)
			   }
			   //fmt.Println("---->",row[2])
			   p, _ := strconv.ParseFloat(row[2], 64)
			   //fmt.Println("-----",excelDateToDate(row[1]).Unix())

			   //time1, _ := time.ParseInLocation(timeLayout, row[1], loc)
			   //time2, _ := time.ParseInLocation(timeLayout, row[3], loc)
			   item := &Item{
				   Code:  row[0],
				   Time1: excelDateToDate(row[1]).Unix(),
				   Price: p,
				   Time2: excelDateToDate(row[3]).Unix(),
			   }
			   data = append(data, item)
			   //for j, colCell := range row {
				//   fmt.Print(j,"-",colCell, "\t")
			   //}
			   //fmt.Println()
		   }
		   if err = rows1.Close(); err != nil {
			   fmt.Println(err)
		   }
	   }
	//result1用来存放每一次降价的记录
	result1 := make([]*Result,0,5000)
	//map1用来存放上一次的价格 code price
	map1 := make(map[string]float64)
    //map2用来存最后一次降价的价格
	map2 := make(map[string]float64)
	for _,v := range data{
		//fmt.Println(v.Code,"\t",v.Price,"\t",v.Time1,"\t",v.Time2)
		if p,ok:=map1[v.Code];ok{
			//map1中含有该code
			//当前记录的价格比map1中的价格大,替换进去
			if v.Price > p{
				map1[v.Code] = v.Price
			}
			//当前记录的价格比map1中的价格小,说明是一次降价
			if v.Price < p{
				map1[v.Code] = v.Price
				tm1 := time.Unix(v.Time1, 0)
				tm2 := time.Unix(v.Time2, 0)
				r := &Result{
					Code:  v.Code,
					Time1: tm1.Format("2006/01/02"),
					Price: v.Price,
					Time2: tm2.Format("2006/01/02"),
					Flag:  "",
					LastPrice: p,
				}
				if v.Time2 > v.Time1 {
					r.Flag = "1"
				}else if v.Time2 == v.Time1 {
					r.Flag = "0"
				}
				result1 = append(result1, r)
				map1[v.Code] = v.Price
				map2[v.Code]= v.Price
			}

		}else {
			//map1中没有，说明首次遇到该code，添加进去
			map1[v.Code] = v.Price
		}
	}

	exclude := make([]string,0)
	//此时map1存放的是大家最后的价格
	//此时map2存放的是大家最后一次降价降价后的价格
	for k,v := range map2{
		if v2,ok := map1[k];ok{
			if v2 > v {
				//说明最后价格又涨回去了
				exclude = append(exclude, k)
			}
		}
	}

	for _,v := range exclude{
		fmt.Println("exclude :",v)
	}

	for k,v := range result1 {
		if In(v.Code, exclude) {
			result1 = append(result1[:k], result1[k+1:]...)
		}
	}

	//for _,v := range result1{
	//	//fmt.Println(v.Code,"\t",v.Price,"\t",v.Time1,"\t",v.Time2,"\t",v.Flag)
	//}
	ef := excelize.NewFile()
	sw, _ := ef.NewStreamWriter("Sheet1")
	err = StreamWriteFirstRow(sw, result1)
	if err != nil{
		return
	}
	sw.Flush()
	ef.SaveAs("result.xlsx")

	ef2 := excelize.NewFile()
	sw2, _ := ef2.NewStreamWriter("Sheet1")
	err = StreamWriteFirstRow2(sw2, exclude)
	if err != nil{
		return
	}
	sw2.Flush()
	ef2.SaveAs("exclude.xlsx")

}

func In(val string, sli []string) bool {
	for _,v := range sli{
		if v == val {
			return true
		}
	}
	return false
}

func StreamWriteFirstRow(sw *excelize.StreamWriter, data []*Result) error {
	//firstRow := reflect.ValueOf(data).Index(0).Elem() // 第1个数据的反射Value
	//v1 := firstRow.Field(0) // 首个字段的反射Value
	//v2 := firstRow.Field(1) // 第二个个字段的反射Value
	//v3 := firstRow.Field(2) // 第三个字段的反射Value
	//v4 := firstRow.Field(3) // 第四个字段的反射Value
	//v5 := firstRow.Field(4) // 第五个字段的反射Value

	//length := firstRow.NumField() // 结构体字段数量
	style := &excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
		},
	}
	styleID, err := sw.File.NewStyle(style) // 创建样式
	if err != nil {
		return err
	}

	//插入 第一行列名
	title := make([]interface{}, 6)
	title[0] = excelize.Cell{
		Value: "报账点编码",
	}
	title[1] = excelize.Cell{
		Value: "降价后价格",
	}
	title[2] = excelize.Cell{
		Value: "降价协议开始日期",
	}
	title[3] = excelize.Cell{
		Value: "首次关联降价协议起始日期",
	}
	title[4] = excelize.Cell{
		Value: "是否存在未按压降单价报账",
	}
	title[5] = excelize.Cell{
		Value: "降价前价格",
	}
	// 第1行要显示列名
	axis, err := excelize.CoordinatesToCellName(1, 1)
	if err != nil {
		return err
	}

	err = sw.SetRow(axis, title)
	if err != nil{
		fmt.Println("set row error :%s", err.Error())
		return err
	}

    i := 2
	for _,v := range data{
		row := make([]interface{}, 6) // 创建一个切片，表示一行数据
		row[0] = excelize.Cell{
			Value: v.Code,
			StyleID: styleID,
		}
		row[1] = excelize.Cell{
			Value: v.Price,
			StyleID: styleID,
		}
		row[2] = excelize.Cell{
			Value: v.Time1,
			StyleID: styleID,
		}
		row[3] = excelize.Cell{
			Value: v.Time2,
			StyleID: styleID,
		}
		row[4] = excelize.Cell{
			Value: v.Flag,
			StyleID: styleID,
		}
		row[5] = excelize.Cell{
			Value: v.LastPrice,
			StyleID: styleID,
		}

		// 每一行都是从列号1开始；行号从2开始，因为假定第1行要显示列名
		axis, err := excelize.CoordinatesToCellName(1, i)
		if err != nil {
			return err
		}

		err = sw.SetRow(axis, row, excelize.RowOpts{Height: 16})
		if err != nil{
			fmt.Println("set row error :%s", err.Error())
			return err
		}
		i++
	}




	// 流式写入行，并指定高度
	return nil
}

func StreamWriteFirstRow2(sw *excelize.StreamWriter, data []string) error {
	//插入 第一行列名
	title := make([]interface{}, 1)
	title[0] = excelize.Cell{
		Value: "报账点编码",
	}
	// 第1行要显示列名
	axis, err := excelize.CoordinatesToCellName(1, 1)
	if err != nil {
		return err
	}

	err = sw.SetRow(axis, title)
	if err != nil{
		fmt.Println("set row error :%s", err.Error())
		return err
	}
	i := 2
	for _,v := range data{
		row := make([]interface{}, 1) // 创建一个切片，表示一行数据
		row[0] = excelize.Cell{
			Value: v,

		}


		// 每一行都是从列号1开始；行号从2开始，因为假定第1行要显示列名
		axis, err := excelize.CoordinatesToCellName(1, i)
		if err != nil {
			return err
		}

		err = sw.SetRow(axis, row, excelize.RowOpts{Height: 16})
		if err != nil{
			fmt.Println("set row error :%s", err.Error())
			return err
		}
		i++
	}

	return nil
}

func excelDateToDate(excelDate string) time.Time {
	excelTime := time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	var days, _ = strconv.Atoi(excelDate)
	return excelTime.Add(time.Second * time.Duration(days*86400))
}

func SliceDelete(seq []string, index int) {

	seq = append(seq[:index], seq[index+1:]...)

	// 输出链接后的切片

	fmt.Println(seq)

}

//func main() {
//	f := excelize.NewFile()
//	// Create a new sheet.
//	index := f.NewSheet("Sheet2")
//	// Set value of a cell.
//	f.SetCellValue("Sheet2", "A2", "Hello world.")
//	f.SetCellValue("Sheet1", "B2", 100)
//	// Set active sheet of the workbook.
//	f.SetActiveSheet(index)
//	// Save spreadsheet by the given path.
//	if err := f.SaveAs("Book1.xlsx"); err != nil {
//		fmt.Println(err)
//	}
//
//}

