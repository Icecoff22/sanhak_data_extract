package com.example.test;

import com.aspose.cad.Image;
import com.aspose.cad.cadexceptions.ImageLoadException;
import com.aspose.cad.fileformats.cad.CadImage;
import com.aspose.cad.fileformats.cad.cadconsts.CadEntityTypeName;
import com.aspose.cad.fileformats.cad.cadobjects.CadBaseEntity;
import com.aspose.cad.fileformats.cad.cadobjects.CadBlockEntity;
import com.aspose.cad.fileformats.cad.cadobjects.CadMText;
import com.aspose.cad.fileformats.cad.cadobjects.CadText;
import com.aspose.cad.imageoptions.CadRasterizationOptions;
import com.aspose.cad.imageoptions.Jpeg2000Options;
import com.aspose.cad.imageoptions.JpegOptions;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;
import java.util.concurrent.*;

@SpringBootApplication
public class TestApplication {
	public static final String dataDir = "C:\\Users\\wjdtm\\OneDrive\\바탕 화면\\대학\\산학연계\\데이터(3차)\\[18ED17] 완주삼봉지구스마트도시 정보통신 설계용역";

	public static final String outputDir = "C:\\Users\\wjdtm\\OneDrive\\바탕 화면\\test\\3차\\[18ED17] 완주삼봉지구스마트도시 정보통신 설계용역";

	public static final String ImgDataSavePath = "C:\\Users\\wjdtm\\OneDrive\\바탕 화면\\test\\이미지\\데이터(4차)\\[14ED15] 해양기상신호표지 운영관리시스템 구축설계 기술지원";

	public static void searchCadFleInDataDir() {
		//System.out.println(dataDir.substring(46));
		System.out.println("searchCadFileInDataDir");
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("sheet1");
		try {
			System.out.println("searchCadFileInDataDir2222");
			Files.walkFileTree(Paths.get(dataDir), new SimpleFileVisitor<>() {
				int n=0;
				//String fileName = dataDir.substring(46);
				@Override
				public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
					System.out.println("visitFile");
					//System.out.println(file.toAbsolutePath().toString());
					if(!Files.isDirectory(file)&&(file.getFileName().toString().contains("XREF")||file.getFileName().toString().contains("확대평면도")||file.getFileName().toString().contains("BASE")))
						return FileVisitResult.CONTINUE;
					if (!Files.isDirectory(file) && file.getFileName().toString().contains(".dwg")) {
						//String final_fileName = fileName+ "\\"+file.getFileName().toString();
						String filePath = file.toAbsolutePath().toString();
						Workbook workbook = new XSSFWorkbook();
						Sheet sheet = workbook.createSheet("sheet1");

						ArrayList<String> fileIndex = extractCadIndex(filePath);
						String fileName = filePath.substring(46);
						int lastIndexOfFileName = fileName.lastIndexOf("\\");
						System.out.println(lastIndexOfFileName);
						String tmpFileName = fileName.substring(lastIndexOfFileName+1);
						System.out.println(tmpFileName);
						String final_fileName = tmpFileName.substring(0, tmpFileName.length()-4);//.dwg잘라내기
						System.out.println(filePath);
						System.out.println(fileName);
						System.out.println(final_fileName);

						//String fileName = file.getFileName().toString();
						//String final_fileName = fileName.substring(0, fileName.length()-4);//.dwg잘라내기
						//String ImgSavedName = ImgDataSavePath+"\\"+final_fileName+".jpeg";
						//System.out.println(final_fileName);
						//CadToJpeg(filePath, ImgSavedName);
						Row row = sheet.createRow(0);
						Cell cell = row.createCell(0);
						cell.setCellValue(fileName);
						int size = fileIndex.size();
						for(int i=0;i<size;i++){
							if(i>16382)
								row.createCell(i-16382).setCellValue(fileIndex.get(i));
							else row.createCell(i+1).setCellValue(fileIndex.get(i));
						}
						try{
							String xlsxDir = outputDir+"\\"+final_fileName+".xlsx";
							System.out.println(xlsxDir);
							FileOutputStream fileOut = new FileOutputStream(xlsxDir);
							workbook.write(fileOut);
							fileOut.close();
							workbook.close();
						}catch(IOException e){
							System.out.println("poi error");
							e.printStackTrace();
						}
					}
					return FileVisitResult.CONTINUE;
				}
			});
		} catch (IOException e) {
			System.out.println("visitFIle error");
			e.printStackTrace();
		}
		System.out.println("End");
		/*try{
			FileOutputStream fileOut = new FileOutputStream("[18ED15] 고양덕은지구 스마트도시 정보통신 설계용역.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
		}catch(IOException e){
			System.out.println("poi error");
			e.printStackTrace();
		}*/

	}
	private static ArrayList<String> extractCadIndex(String cad) {
		ArrayList<String> index = new ArrayList<>();
		try {

			CadImage cadImage = (CadImage) CadImage.load(cad);
			for (CadBlockEntity blockEntity : cadImage.getBlockEntities().getValues()) {
				for (CadBaseEntity entity : blockEntity.getEntities()) {
					if (entity.getTypeName() == CadEntityTypeName.TEXT) {
						CadText cadText = (CadText) entity;
						if(filterCadIndex(cadText.getDefaultValue())==""){
							continue;
						}
						index.add(filterCadIndex(cadText.getDefaultValue()));
					}
					else if (entity.getTypeName() == CadEntityTypeName.MTEXT) {
						CadMText cadMText = (CadMText) entity;
						if(filterCadIndex(cadMText.getText())==""){
							continue;
						}
						index.add(filterCadIndex(cadMText.getText()));
					}
				}
			}
			return index;
		} catch (Exception e) {
			e.printStackTrace();
			return new ArrayList<>();

		}
	}

	private static String filterCadIndex(String index) {
		String filtered = index.replace(" ", "");
		/*int numCnt = (int) filtered.chars().filter(c -> c >= '0' && c <= '9').count();
		if (numCnt >= filtered.length() / 2)
			return "";*/
		return filtered;
	}

	public static void CadToJpeg(String filePath, String imgPath) {
		System.out.println("CadToJpeg");


		CadRasterizationOptions rasterizationOptions = new CadRasterizationOptions();
		rasterizationOptions.setPageHeight(1680);
		rasterizationOptions.setPageWidth(1920);

		JpegOptions options = new JpegOptions();
		options.setVectorRasterizationOptions(rasterizationOptions);

		synchronized (TestApplication.class){
			try(Image image = Image.load(filePath)){
				//image.save(imgPath, options);
				image.save(imgPath, options);
			}catch (ImageLoadException e){
				System.out.println("Image Load Failed");
			}catch(Exception e){
				System.out.println(e.getMessage());
			}
		}
	}

	public static void main(String[] args) {
		searchCadFleInDataDir();

	}

}
