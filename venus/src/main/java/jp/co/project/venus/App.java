package jp.co.project.venus;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import jp.co.project.venus.model.BaseDto;
import jp.co.project.venus.model.TestDto;
import jp.co.project.venus.transfer.ExcelService;
import jp.co.project.venus.transfer.TestExcelTransfer;

/**
 * Hello world!
 *
 */
public class App {
	@SuppressWarnings("unchecked")
	public static void main(String[] args) {
		System.out.println("Hello World!");
		ExcelService service = new TestExcelTransfer();
		try {
			InputStream stream = new BufferedInputStream(new FileInputStream("E:/Book1.xlsx"));
			Map<String, ?> map = service.readExcel(stream);

			if (service.hasError()) {
				service.getErrorResultList().forEach(result -> {
					System.out.println("row:" + result.getRowIdx() + " col:" + result.getColIdx() + " address:"
							+ result.getCellAddress() + " msg:" + result.getMessage());
				});
			}

			map.forEach((key, value) -> {
				System.out.println("key : " + key);
				if (value instanceof List) {
					((List<BaseDto>) value).forEach(element -> {
						TestDto dto = (TestDto) element;
						System.out.println(dto.getTest1());
						System.out.println(dto.getDecimal().toString());
					});
				}
			});

//			char c = 'a';
//			Random random = new Random();
//			List<TestDto> list = new ArrayList<>();
//			for (int i = 0; i < 10; i++) {
//				int num = random.nextInt(26);
//				TestDto dto = new TestDto();
//				dto.setTest1(String.valueOf((char)(c + num)));
//				dto.setDecimal(new BigDecimal(num));
//				list.add(dto);
//			}
//
//			String path = "E:/workspace/projectT/venus/src/main/resources/static/sample.xlsx";
//			InputStream is = new BufferedInputStream(new FileInputStream(path));
//			Map<String, Object> model = new HashMap<>();
//			model.put(ExcelMapKey.INPUT_DETAIL, list);
//			ExcelWorkbook workbook = service.makeWorkbook(is, model);
//			OutputStream os = new FileOutputStream("E:/out_sample.xlsx");
//			workbook.save(os);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Exception e) {
			// TODO 自動生成された catch ブロック
			e.printStackTrace();
		}
	}
}
