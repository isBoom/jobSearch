package main

import (
	"fmt"
	"github.com/tealeg/xlsx/v3"
	"os"
	"strings"
)
func Write(res [][]string){
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("中央国家行政机关省级以下直属机构")
	if err != nil {
		fmt.Printf(err.Error())
	}
	cellName := []string{"部门代码", "部门名称", "用人司局", "机构性质", "招考职位", "职位属性","职位分布", "职位简介","职位代码","机构层级",
		"考试类别","招考人数","专业","学历","学位","政治面貌","基层工作最低年限","是否在面试阶段组织专业能力测试","面试人员比例","工作地点",
		"落户地点","备注","部门网站","咨询电话1","咨询电话2","咨询电话3"}
	row := sheet.AddRow()
	for _,c := range  cellName{
			row.AddCell().Value = c
	}
	for _, r := range res{
		row = sheet.AddRow()
		for _,c:=range r{
			row.AddCell().Value = c
		}
	}
	err = file.Save("testNew.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
}
func Read() (output [][]string) {
	xlFile, err := xlsx.OpenFile("./test.xlsx")
	if err != nil {
		panic(err)
	}
	for _, sheet := range xlFile.Sheets {
		if sheet.Name == "中央国家行政机关省级以下直属机构" {
			fmt.Println("sheet name: ", sheet.Name)
			err = sheet.ForEachRow(func(r *xlsx.Row) error {
				cUse:=0
				cells := []string{}
				//检索河南 华中
				if strings.Contains(r.GetCell(1).Value,"河南") || strings.Contains(r.GetCell(1).Value,"华中")|| strings.Contains(r.GetCell(1).Value,"中南") || strings.Contains(r.GetCell(1).Value,"郑州"){
					//不要党员
					if r.GetCell(15).Value != "中共党员" {
						//不要工作经验
						if r.GetCell(16).Value == "无限制" {
							//不要基层经验
							if r.GetCell(17).Value == "无限制" {
								//不要四级
								if !strings.Contains(r.GetCell(22).Value, "425") {
									//查相关专业
									zyName := r.GetCell(12).Value
									if strings.Contains(zyName, "计算机") || strings.Contains(zyName, "软件") || strings.Contains(zyName, "工学") || strings.Contains(zyName, "不限") {
										err = r.ForEachCell(func(c *xlsx.Cell) error {
											cUse = 1
											cells = append(cells, c.Value)
											return nil
										})
									}
								}
							}
						}
					}
				}
				if err != nil {
					return err
				}else if cUse==1{
					output = append(output, cells)
				}
				return nil
			})
			if err != nil {
				fmt.Printf("sheet.ForEachRow err=%#v\n",err)
			}
		}
	}
	return
}
func main() {
	os.Remove("./testNew.xlsx")
	Write(Read())
}
