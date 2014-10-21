package tool.sql;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class SqlBuildTool {

	private static final String FILE_PATH = "/Users/wangchaochao/data/eclipse/workspace-bailian/sql/source/sqlTemp.xlsx";

	private static final String PK_MARK = "pk";

	public void buildSql(String path) throws IOException {

		List<String> lines = new ArrayList<String>();
		List<String> comms = new ArrayList<String>();

		InputStream inp = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(inp);
		XSSFSheet st = workbook.getSheetAt(0);
		XSSFRow row0 = st.getRow(0);
		XSSFCell row0cell1 = row0.getCell(1);
		String tableName = row0cell1.getStringCellValue();
		lines.add("-- Create table");
		lines.add("create table " + tableName);
		lines.add("(");

		XSSFRow row4 = st.getRow(4);
		XSSFCell row4cell1 = row4.getCell(1);
		String tableDesc = row4cell1.getStringCellValue();
		comms.add("-- Add comments to the table");
		comms.add("comment on table TABLE_TEST is '" + tableDesc + "';");
		comms.add("");
		comms.add("-- Add comments to the columns");

		int lastRowNum = st.getLastRowNum();

		String pk = null;

		for (int i = 7; i < lastRowNum; i++) {
			String line = "";
			XSSFRow row = st.getRow(i);
			XSSFCell rowcell0 = row.getCell(0);
			String colDesc = null;
			if (rowcell0 != null) {
				colDesc = rowcell0.getStringCellValue();
				comms.add("comment on column TABLE_TEST.USERID is '" + colDesc
						+ "';");
			}

			XSSFCell rowcell1 = row.getCell(1);
			String colName = null;
			if (rowcell1 != null) {
				colName = rowcell1.getStringCellValue();
			}
			line = "  " + line + colName + "     ";

			XSSFCell rowcell2 = row.getCell(2);
			String colType = null;
			if (rowcell2 != null) {
				colType = rowcell2.getStringCellValue();
			}
			line = line + colType;

			XSSFCell rowcell3 = row.getCell(3);
			String colLen = null;
			if (rowcell3 != null) {
				colLen = (int)rowcell3.getNumericCellValue() + "";
			}

			XSSFCell rowcell4 = row.getCell(4);
			String colDou = null;
			if (rowcell4 != null) {
				colDou = (int)rowcell4.getNumericCellValue() + "";
			}
			if (colDou != null) {
				line = line + "(" + colLen + "," + colDou + "),";
			} else {
				line = line + "(" + colLen + "),";
			}

			XSSFCell rowcell5 = row.getCell(5);
			String colpk = null;
			if (rowcell5 != null) {
				colpk = rowcell5.getStringCellValue();
				if (PK_MARK.equals(colpk)) {
					pk = colName;
				}
			}
			lines.add(line);
		}
		int index = lines.size() - 1;
		String lastLine = lines.get(index);
		String newLine = lastLine.substring(0, lastLine.length() - 1);
		lines.set(index, newLine);
		lines.add(")");
		lines.add("tablespace " + tableName);
		lines.add("  pctfree 10");
		lines.add("  initrans 1");
		lines.add("  maxtrans 255");
		lines.add("  storage");
		lines.add("  (");
		lines.add("    initial 64K");
		lines.add("    minextents 1");
		lines.add("    maxextents unlimited");
		lines.add("  );");
		lines.add("");

		lines.addAll(comms);
		lines.add("");
		lines.add("-- Create/Recreate primary, unique and foreign key constraints");
		lines.add("alter table " + tableName);
		lines.add("  add primary key (" + pk + ")");
		lines.add("  using index");
		lines.add("  tablespace " + tableName);
		lines.add("  pctfree 10");
		lines.add("  initrans 2");
		lines.add("  maxtrans 255");
		lines.add("  storage");
		lines.add("  (");
		lines.add("    initial 64K");
		lines.add("    minextents 1");
		lines.add("    maxextents unlimited");
		lines.add("  );");
		for (String li : lines) {
			System.out.println(li);
		}
	}

	public static void main(String[] args) throws IOException {
		SqlBuildTool tool = new SqlBuildTool();
		tool.buildSql(FILE_PATH);
	}
}
