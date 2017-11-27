import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
 
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
 
import bean.ExcelChangeTableBean;
 
public class ExcelChangeTable {
 
	public static void main(String[] args) throws IOException {
		//파일을 읽기위해 엑셀파일을 가져온다
		FileInputStream fis=new FileInputStream("C:\\guide2.xls");
		HSSFWorkbook workbook=new HSSFWorkbook(fis);
		int rowindex=0;
		int columnindex=0;
		int rows= 0;
		int cells = 0;
		StringBuffer tableBody = new StringBuffer();
		List<ExcelChangeTableBean> list = new ArrayList<>();
		//시트 수 (첫번째에만 존재하므로 0을 준다)
		//만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다
		HSSFSheet sheet=workbook.getSheetAt(0);
		//행의 수
		rows=sheet.getPhysicalNumberOfRows();
		for(rowindex=1;rowindex<rows;rowindex++){
		    //행을 읽는다
		    HSSFRow row=sheet.getRow(rowindex);
		    if(row !=null){
		        //셀의 수
		        cells=row.getPhysicalNumberOfCells();
		        ExcelChangeTableBean bean = new ExcelChangeTableBean();
		        list.add(bean);
		        for(columnindex=0;columnindex<=cells;columnindex++){
		            //셀값을 읽는다
		            HSSFCell cell=row.getCell(columnindex);
		            String value="";
		            //셀이 빈값일경우를 위한 널체크
		            if(cell==null){
		                continue;
		            }else{
		                //타입별로 내용 읽기
			                switch (cell.getCellType()){
			                case HSSFCell.CELL_TYPE_FORMULA:
			                    value=cell.getCellFormula();
			                    break;
			                case HSSFCell.CELL_TYPE_NUMERIC:
			                    value=cell.getNumericCellValue()+"";
			                    break;
			                case HSSFCell.CELL_TYPE_STRING:
			                    value=cell.getStringCellValue()+"";
			                    break;
			                case HSSFCell.CELL_TYPE_BLANK:
			                    value=cell.getBooleanCellValue()+"";
			                    break;
			                case HSSFCell.CELL_TYPE_ERROR:
			                    value=cell.getErrorCellValue()+"";
			                    break;
			                }
		            }
		            if(columnindex==0) {
		            	bean.setRefuseNation(value);
		            }
		            else if(columnindex==1) {
		            	value=value.substring(0,6);
		            	bean.setDate(value);
		            }else if(columnindex==2) {
		            	bean.setProduct(value);
		            }else if(columnindex==3) {
		            	bean.setRefuseReason(value);
		            }else if(columnindex==4) {
		            	bean.setOrijin(value);
		            }else if(columnindex==5) {
		            	bean.setGroup(value);
		            }else if(columnindex==6) {
		            	bean.setGuide(value);
		            }
		          /*  if(columnindex<5) {
		            	tableBody.append("<td>"+value+"</td>\n");
		            }*/
		            }
		        /*tableBody.append("</tr>\n");*/
		        }
		}
		/*System.out.println(tableBody.toString());*/
		tableBody.append("◆지침 : "+list.get(0).getGuide()+"\n");
		tableBody.append("<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" class=\"table_a3\">\r\n" + 
				"<tbody>\r\n" + 
				"<tr>\r\n" + 
				"<td>\r\n" + 
				"<table class=\"table2\"><colgroup><col width=\"75\" /><col width=\"80\" /><col width=\"140\" /><col width=\"*\" /><col width=\"105\" /></colgroup><thead>\r\n" + 
				"<tr>\r\n" + 
				"<th width=\"68\" scope=\"col\">거부국가</th>\r\n" + 
				"<th width=\"34\" scope=\"col\">년도</th>\r\n" + 
				"<th width=\"134\" scope=\"col\">제품</th>\r\n" + 
				"<th width=\"110\" scope=\"col\">거부사유</th>\r\n" + 
				"<th width=\"64\" scope=\"col\">원산지</th></tr></thead>\r\n" + 
				"<tbody>\r\n" + 
				"");
		tableBody.append("<tr>\n");
		tableBody.append("<td>"+list.get(0).getRefuseNation()+"</td>\n");
		tableBody.append("<td>"+list.get(0).getDate()+"</td>\n");
		tableBody.append("<td>"+list.get(0).getProduct()+"</td>\n");
		tableBody.append("<td>"+list.get(0).getRefuseReason()+"</td>\n");
		tableBody.append("<td>"+list.get(0).getOrijin()+"</td>\n");
		tableBody.append("</tr>\n");
		if(!list.get(0).getGuide().equals(list.get(1).getGuide())) {
			tableBody.append("</tbody></table></td></tr></tbody></table>\n\n\n\n\n\n\n\n");
		}
		for(int index=0;index<list.size();index++) {
			if(index-1<0) {
				continue;
			}
			else if(!list.get(index-1).getGuide().equals(list.get(index).getGuide())) {
				tableBody.append("◆지침 : "+list.get(index).getGuide()+"\n");
				tableBody.append("<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" class=\"table_a3\">\r\n" + 
						"<tbody>\r\n" + 
						"<tr>\r\n" + 
						"<td>\r\n" + 
						"<table class=\"table2\"><colgroup><col width=\"75\" /><col width=\"80\" /><col width=\"140\" /><col width=\"*\" /><col width=\"105\" /></colgroup><thead>\r\n" + 
						"<tr>\r\n" + 
						"<th width=\"68\" scope=\"col\">거부국가</th>\r\n" + 
						"<th width=\"34\" scope=\"col\">년도</th>\r\n" + 
						"<th width=\"134\" scope=\"col\">제품</th>\r\n" + 
						"<th width=\"110\" scope=\"col\">거부사유</th>\r\n" + 
						"<th width=\"64\" scope=\"col\">원산지</th></tr></thead>\r\n" + 
						"<tbody>\r\n" + 
						"");
			}
				tableBody.append("<tr>\n");
				tableBody.append("<td>"+list.get(index).getRefuseNation()+"</td>\n");
				tableBody.append("<td>"+list.get(index).getDate()+"</td>\n");
				tableBody.append("<td>"+list.get(index).getProduct()+"</td>\n");
				tableBody.append("<td>"+list.get(index).getRefuseReason()+"</td>\n");
				tableBody.append("<td>"+list.get(index).getOrijin()+"</td>\n");
				tableBody.append("</tr>\n");
				if(index>=list.size()-1) {
					break;
				}
				else if(!list.get(index).getGuide().equals(list.get(index+1).getGuide())) {
					tableBody.append("</tbody></table></td></tr></tbody></table>\n\n\n\n\n\n\n\n");
				}
			
		}
		tableBody.append("</tbody></table></td></tr></tbody></table>\n\n\n\n\n\n\n\n");
		
		System.out.println(tableBody.toString());
	}
}
